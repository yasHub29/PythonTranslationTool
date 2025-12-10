# modules/excel_translator/excel_translator.py
# Main Excel translation workflow (writes via Excel COM to preserve shapes)

import os
from config.settings import OUTPUT_DIR
from modules.excel_translator.excel_reader import read_excel_for_translation
from modules.excel_translator.excel_writer import write_translated_excel_preserve_format

class ExcelTranslator:
    def __init__(self, translator):
        self.translator = translator

    def process(self, input_path, direction="en->ja"):
        """
        Translate Excel text (cells + sheet names) while preserving layout, formatting, and images.
        Shapes, text boxes, and arrows are retained but not translated (preserved as-is).
        Uses Excel COM to write the final file so shapes/textboxes are not lost.
        """
        # Step 1: Read workbook and collect text cells (using openpyxl)
        wb, cells_to_translate = read_excel_for_translation(input_path)
        translated_results = []

        # Step 2: Translate all cell values (collect translations)
        for sheet_name, coord, text in cells_to_translate:
            try:
                translated_text = self.translator.translate_text(text, direction=direction)
            except Exception:
                translated_text = text  # fallback to original text
            translated_results.append((sheet_name, coord, translated_text))

        # Step 3: Build sheet rename mapping (old_name -> new_name)
        sheet_renames = []
        for sheet in wb.worksheets:
            try:
                translated_name = self.translator.translate_text(sheet.title, direction=direction)
                safe_name = translated_name.strip()[:31]  # Excel limit
                if not safe_name:
                    continue
                # If duplicate among targets, append suffix (we will ensure uniqueness)
                sheet_renames.append((sheet.title, safe_name))
            except Exception:
                # skip rename if translation fails
                continue

        # Step 4: Write translated text and apply sheet renames using Excel COM
        output_path = write_translated_excel_preserve_format(
            input_path, translated_results, sheet_renames, OUTPUT_DIR
        )

        return output_path
