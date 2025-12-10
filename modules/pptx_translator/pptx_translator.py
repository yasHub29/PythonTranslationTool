# modules/pptx_translator/pptx_translator.py
# Translate each shape's text individually and keep mapping for writer.

from .pptx_reader import read_pptx
from .pptx_writer import write_pptx_from_template
from core.translate_text_google import Translator as HighLevelTranslator

class PPTXTranslator:
    def __init__(self, high_level_translator: HighLevelTranslator):
        self.hl = high_level_translator

    def process(self, src_path, direction='en->ja'):
        """
        Translate each text-containing shape individually.
        The reader provides shape paths so writer can replace text in-place.
        """
        slides = read_pptx(src_path)
        translated_slides = []

        for slide in slides:
            shape_texts = slide.get("shape_texts", [])
            translated_shape_texts = []

            for path, text in shape_texts:
                try:
                    tr = self.hl.translate_text(text, direction)
                except Exception:
                    tr = text  # fallback
                translated_shape_texts.append((path, tr))

            translated_slides.append({
                "translated_shape_texts": translated_shape_texts,
                "images": slide.get("images", [])
            })

        out_path = self.hl._make_output_path(src_path)
        write_pptx_from_template(src_path, out_path, translated_slides)
        return out_path
