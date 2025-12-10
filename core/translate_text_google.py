# core/translate_text_google.py
# 高レベルの翻訳ユーティリティ: 各種ファイル形式を自動判別して翻訳を行うモジュール

import os
from datetime import datetime
# deep_translator を使用するためのインポート
from deep_translator import GoogleTranslator 
# ExcelTranslator は translate_file 内で遅延インポートされる
from config.settings import OUTPUT_DIR


class Translator:
    """
    翻訳機能の中核を担うクラス。
    ファイルの種類に応じて適切なモジュールを呼び出し、翻訳処理を行う。
    """

    def __init__(self):
        """出力ディレクトリが存在しない場合は作成"""
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR, exist_ok=True)

    def _parse_direction(self, direction: str):
        """'en->ja' のような翻訳方向文字列をソース言語とターゲット言語に分解する"""
        if not direction or '->' not in direction:
            return ('auto', 'ja')
        parts = direction.split('->')
        src = parts[0].strip() if parts[0].strip() else 'auto'
        dest = parts[1].strip() if len(parts) > 1 and parts[1].strip() else 'ja'
        return (src, dest)

    def translate_text(self, text: str, direction: str = 'en->ja'):
        """
        単一のテキスト文字列を deep_translator の GoogleTranslator を使用して翻訳する。
        このメソッドは PPTX や Excel 内のテキスト処理に利用される。
        """
        if text is None:
            return text
        text = str(text)
        if not text.strip():
            return text

        src, dest = self._parse_direction(direction)
        try:
            translator = GoogleTranslator(source=src, target=dest)
            return translator.translate(text)
        except Exception:
            # エラーが発生した場合は元のテキストをそのまま返す
            return text 

    def _make_output_path(self, input_path):
        """翻訳済みファイルの出力パスを生成"""
        base = os.path.basename(input_path)
        name, ext = os.path.splitext(base)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        outname = f"{name}_translated_{ts}{ext}"
        return os.path.join(OUTPUT_DIR, outname)

    def translate_file(self, input_path: str, direction: str = 'en->ja'):
        """
        入力ファイルの拡張子をもとに、適切な翻訳モジュールを呼び出して処理を実行する。
        現在対応している形式: PPTX / XLSX / TXT / CSV / DOCX
        """
        src, dest = self._parse_direction(direction)
        ext = os.path.splitext(input_path)[1].lower()

        # PowerPoint ファイルの処理
        if ext == '.pptx':
            from modules.pptx_translator.pptx_translator import PPTXTranslator
            pptx_tr = PPTXTranslator(self)
            out = pptx_tr.process(input_path, direction=direction)
            return out

        # Excel ファイルの処理
        elif ext in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
            from modules.excel_translator.excel_translator import ExcelTranslator
            excel_tr = ExcelTranslator(self)
            out = excel_tr.process(input_path, direction=direction)
            return out

        # Word DOCX の処理
        elif ext == '.docx':
            from modules.docx_translator.docx_translator import DOCXTranslator
            docx_tr = DOCXTranslator(self)
            out = docx_tr.process(input_path, direction=direction)
            return out

        # テキスト / CSV ファイルの処理
        elif ext in ('.txt', '.csv'):
            with open(input_path, 'r', encoding='utf-8') as f:
                txt = f.read()
            translated = self.translate_text(txt, direction=direction)
            outpath = self._make_output_path(input_path)
            with open(outpath, 'w', encoding='utf-8') as f:
                f.write(translated)
            return outpath

        # 未対応の拡張子
        else:
            raise ValueError(f'未対応のファイル形式です: {ext}')
