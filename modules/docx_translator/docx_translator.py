# modules/docx_translator/docx_translator.py
# DOCX 単体翻訳モジュール
# reader -> translate -> writer のフローで動作します

from .docx_reader import read_docx
from .docx_writer import write_docx_from_template
from core.translate_text_google import Translator as HighLevelTranslator
import os

class DOCXTranslator:
    def __init__(self, high_level_translator: HighLevelTranslator):
        self.hl = high_level_translator

    def process(self, src_path, direction='en->ja'):
        """
        - src_path: 元の docx ファイル
        - direction: 'en->ja' のような翻訳方向
        戻り値: 出力ファイルパス
        """
        # 1) 読み取り
        structure = read_docx(src_path)

        # 2) 翻訳（段落、テーブル、ヘッダ、フッタ）
        translated = {
            'paragraphs': [],
            'tables': [],
            'headers': [],
            'footers': []
        }

        # paragraphs: ("para", i, text)
        for _type, idx, text in structure.get('paragraphs', []):
            try:
                tr = self.hl.translate_text(text, direction)
            except Exception:
                tr = text
            translated['paragraphs'].append((_type, idx, tr))

        # tables: ("table", t_idx, r_idx, c_idx, p_idx, text)
        for item in structure.get('tables', []):
            _type, t_idx, r_idx, c_idx, p_idx, text = item
            try:
                tr = self.hl.translate_text(text, direction)
            except Exception:
                tr = text
            translated['tables'].append((_type, t_idx, r_idx, c_idx, p_idx, tr))

        # headers
        for _type, s_idx, p_idx, text in structure.get('headers', []):
            try:
                tr = self.hl.translate_text(text, direction)
            except Exception:
                tr = text
            translated['headers'].append((_type, s_idx, p_idx, tr))

        # footers
        for _type, s_idx, p_idx, text in structure.get('footers', []):
            try:
                tr = self.hl.translate_text(text, direction)
            except Exception:
                tr = text
            translated['footers'].append((_type, s_idx, p_idx, tr))

        # 3) 出力パス生成
        out_path = self.hl._make_output_path(src_path)

        # 4) writer による保存（元の書式・画像は preserved）
        write_docx_from_template(src_path, out_path, translated)

        return out_path
