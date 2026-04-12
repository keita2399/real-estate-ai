"""
Flask Web UI - 不動産書類AI自動入力システム
"""
import os
import uuid
import logging
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
from dotenv import load_dotenv

from extractor import configure_gemini, extract_from_pdfs
from excel_writer import write_to_excel, preview_extracted
from create_template import create_template

load_dotenv()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB上限

UPLOAD_DIR  = Path("uploads")
OUTPUT_DIR  = Path("output")
TEMPLATE_PATH = Path("template/jyuujiku.xlsx")

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# テンプレート確認
if not TEMPLATE_PATH.exists():
    create_template(str(TEMPLATE_PATH))

configure_gemini()

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/process", methods=["POST"])
def process():
    """PDFをアップロードしてGeminiで解析、Excelを返す"""
    if "files" not in request.files:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "ファイルが選択されていません"}), 400

    session_id = uuid.uuid4().hex
    session_dir = UPLOAD_DIR / session_id
    session_dir.mkdir(parents=True)

    saved_paths = []
    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            return jsonify({"error": f"{f.filename} はPDFではありません"}), 400
        dest = session_dir / f.filename
        f.save(dest)
        saved_paths.append(str(dest))

    try:
        extracted = extract_from_pdfs(saved_paths)

        # 会社情報をマージ
        company_info = {}
        try:
            import json as _json
            company_info = _json.loads(request.form.get('company_info', '{}'))
        except Exception:
            pass

        prefix = Path(saved_paths[0]).stem
        output_path = write_to_excel(
            extracted,
            template_path=str(TEMPLATE_PATH),
            output_dir=str(OUTPUT_DIR),
            prefix=prefix,
            company_info=company_info,
        )
        output_file = Path(output_path)
        return jsonify({
            "success": True,
            "filename": output_file.name,
            "extracted": extracted,
            "filled_count": sum(1 for v in extracted.values() if v and v not in ("不明", "記載なし", "なし")),
            "total_count": len(extracted),
        })
    except Exception as e:
        logging.exception("処理中にエラーが発生しました")
        return jsonify({"error": str(e)}), 500
    finally:
        # アップロード一時ファイルを削除
        import shutil
        shutil.rmtree(session_dir, ignore_errors=True)


@app.route("/api/download/<filename>")
def download(filename):
    """生成されたExcelをダウンロード"""
    path = OUTPUT_DIR / filename
    if not path.exists():
        return jsonify({"error": "ファイルが見つかりません"}), 404
    return send_file(
        str(path.resolve()),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
