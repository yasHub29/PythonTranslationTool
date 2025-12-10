# core/google_translator_api.py
# Deep-translator を使用して Google 翻訳を行うモジュール

from deep_translator import GoogleTranslator as GT

class GoogleTranslator:
    def __init__(self, source_lang="auto", target_lang="en"):
        self.source_lang = source_lang
        self.target_lang = target_lang

    def translate(self, text):
        """テキストを翻訳"""
        if not text or not text.strip():
            return text
        try:
            return GT(source=self.source_lang, target=self.target_lang).translate(text)
        except Exception as e:
            print(f"Translation error: {e}")
            return text
