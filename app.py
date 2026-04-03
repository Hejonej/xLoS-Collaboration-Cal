import os
import uuid
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

from processor import process_files, PeriodConfig

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
DEFAULT_TEMPLATE = BASE_DIR / "static" / "default_template.xlsx"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXT = {".xlsx", ".xls"}


def _allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXT


@app.route("/")
def index():
    has_template = DEFAULT_TEMPLATE.exists()
    return render_template("index.html", has_template=has_template)


@app.route("/process", methods=["POST"])
def process():
    try:
        report_month = int(request.form.get("report_month", 0))
        if report_month < 200001 or report_month > 209912:
            return jsonify(ok=False, error="올바른 기준월을 선택해 주세요."), 400

        file_keys = ["xlos_prev", "xlos_curr", "same_prev", "same_curr"]
        labels = {
            "xlos_prev": "전기 xLoS",
            "xlos_curr": "당기 xLoS",
            "same_prev": "전기 동일LoS",
            "same_curr": "당기 동일LoS",
        }

        saved = {}
        job_id = uuid.uuid4().hex[:8]

        for key in file_keys:
            f = request.files.get(key)
            if not f or f.filename == "":
                return jsonify(ok=False, error=f"'{labels[key]}' 파일을 업로드해 주세요."), 400
            if not _allowed(f.filename):
                return jsonify(ok=False, error=f"'{labels[key]}': .xlsx 파일만 가능합니다."), 400
            safe = f"{job_id}_{key}_{secure_filename(f.filename)}"
            path = UPLOAD_DIR / safe
            f.save(str(path))
            saved[key] = str(path)

        # 템플릿: 사용자가 올렸으면 사용, 아니면 내장 템플릿
        template_file = request.files.get("template")
        if template_file and template_file.filename:
            safe = f"{job_id}_template_{secure_filename(template_file.filename)}"
            tpath = UPLOAD_DIR / safe
            template_file.save(str(tpath))
            saved["template"] = str(tpath)
        elif DEFAULT_TEMPLATE.exists():
            saved["template"] = str(DEFAULT_TEMPLATE)
        else:
            return jsonify(ok=False, error="템플릿 파일이 없습니다. 템플릿을 업로드해 주세요."), 400

        cfg = PeriodConfig(report_month)
        out_name = f"xLoS {cfg.month_label} 분석_AUTO.xlsx"
        out_path = str(OUTPUT_DIR / f"{job_id}_{out_name}")

        result = process_files(
            report_month=report_month,
            xlos_prev_path=saved["xlos_prev"],
            xlos_curr_path=saved["xlos_curr"],
            same_prev_path=saved["same_prev"],
            same_curr_path=saved["same_curr"],
            template_path=saved["template"],
            output_path=out_path,
        )

        for key in ["xlos_prev", "xlos_curr", "same_prev", "same_curr"]:
            try:
                os.remove(saved[key])
            except OSError:
                pass
        if "template" in saved and saved["template"] != str(DEFAULT_TEMPLATE):
            try:
                os.remove(saved["template"])
            except OSError:
                pass

        return jsonify(
            ok=True,
            download_id=f"{job_id}_{out_name}",
            xlos_summary=result["xlos_summary"],
            same_summary=result["same_summary"],
            validation=result.get("validation", []),
        )

    except ValueError as e:
        return jsonify(ok=False, error=str(e)), 400
    except Exception as e:
        return jsonify(ok=False, error=f"처리 중 오류 발생: {e}"), 500


@app.route("/update-template", methods=["POST"])
def update_template():
    """내장 템플릿을 새 파일로 교체"""
    f = request.files.get("template")
    if not f or f.filename == "":
        return jsonify(ok=False, error="파일을 선택해 주세요."), 400
    if not _allowed(f.filename):
        return jsonify(ok=False, error=".xlsx 파일만 가능합니다."), 400
    f.save(str(DEFAULT_TEMPLATE))
    return jsonify(ok=True, message="템플릿이 업데이트되었습니다.")


@app.route("/download/<path:filename>")
def download(filename):
    path = OUTPUT_DIR / filename
    if not path.exists():
        return "파일을 찾을 수 없습니다.", 404
    clean_name = filename.split("_", 1)[1] if "_" in filename else filename
    return send_file(str(path), as_attachment=True, download_name=clean_name)


if __name__ == "__main__":
    app.run(debug=True, port=5000, use_reloader=False)
