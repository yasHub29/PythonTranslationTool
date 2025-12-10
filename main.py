import os
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename
from core.translate_text_google import Translator
from modules.excel_translator.excel_translator import ExcelTranslator
from modules.pptx_translator.pptx_translator import PPTXTranslator
from modules.docx_translator.docx_translator import DOCXTranslator  
from core.utils import ensure_dir

# === フォルダ設定 ===
UPLOAD_FOLDER = "uploads"
ensure_dir(UPLOAD_FOLDER)

app = Flask(__name__)
app.secret_key = "supersecretkey"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

translator = Translator()

# === 起動時にアップロードフォルダをクリーンアップ ===
def cleanup_uploads_folder():
    """前回の実行で残ったファイルを削除します。"""
    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.isfile(file_path):
            try:
                os.remove(file_path)
                print(f"[クリーンアップ] 古いファイルを削除しました: {file_path}")
            except Exception as e:
                print(f"[警告] {file_path} を削除できませんでした: {e}")

cleanup_uploads_folder()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/translate", methods=["POST"])
def translate_file():

    # === ファイル検証 ===
    if "file" not in request.files:
        flash("ファイルが選択されていません。")
        return redirect(url_for("index"))

    file = request.files["file"]

    if file.filename == "":
        flash("ファイルが選択されていません。")
        return redirect(url_for("index"))

    # === 言語選択 ===
    translate_from = request.form.get("translate_from")
    translate_to = request.form.get("translate_to")

    if not translate_from or not translate_to:
        flash("翻訳元と翻訳先の言語を選択してください。")
        return redirect(url_for("index"))

    direction = f"{translate_from}->{translate_to}"

    # === ファイル保存 ===
    filename = os.path.basename(file.filename)   # ← Your original behavior (NO secure_filename)
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(file_path)

    ext = os.path.splitext(file_path)[1].lower()
    output_path = None

    try:
        # === Excel ===
        if ext in [".xlsx", ".xls", ".xlsm", ".xltx", ".xltm"]:
            worker = ExcelTranslator(translator)
            output_path = worker.process(file_path, direction)

        # === PowerPoint ===
        elif ext in [".pptx", ".ppt"]:
            worker = PPTXTranslator(translator)
            output_path = worker.process(file_path, direction)

        # === Word DOCX ===
        elif ext == ".docx":
            worker = DOCXTranslator(translator)       
            output_path = worker.process(file_path, direction)

        else:
            flash(f"対応していないファイル形式です: {ext}")
            if os.path.exists(file_path):
                os.remove(file_path)
            return redirect(url_for("index"))

        # === 翻訳完了後、元ファイル削除 ===
        if os.path.exists(file_path):
            os.remove(file_path)

    except Exception as e:
        flash(f"翻訳中にエラーが発生しました: {e}")
        if os.path.exists(file_path):
            os.remove(file_path)
        return redirect(url_for("index"))

    return render_template("result.html", output_path=output_path)


@app.route("/download/<path:filename>")
def download_file(filename):
    return send_file(filename, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
