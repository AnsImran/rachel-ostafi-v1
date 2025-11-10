from __future__ import annotations

import io
import tempfile
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, send_file

from scripts.timesheet_agent import convert_timesheet

app = Flask(__name__)
app.config["TEMPLATE_PATH"] = Path("inputs and examples/template.xlsx")


@app.route("/", methods=["GET", "POST"])
def index():
    error = None

    if request.method == "POST":
        uploaded_file = request.files.get("timesheet")

        if not uploaded_file or uploaded_file.filename == "":
            error = "Please choose a TimesheetPortal .xlsx file to convert."
        else:
            template_path = Path(app.config["TEMPLATE_PATH"])
            if not template_path.exists():
                error = f"Template not found at {template_path}."
            else:
                try:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        tmpdir_path = Path(tmpdir)
                        source_path = tmpdir_path / "source.xlsx"
                        uploaded_file.save(source_path)

                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_path = tmpdir_path / f"Invoice_{timestamp}.xlsx"

                        convert_timesheet(source_path, template_path, output_path)
                        data = output_path.read_bytes()
                except Exception as exc:  # pragma: no cover - user feedback path
                    error = f"Conversion failed: {exc}"
                else:
                    buffer = io.BytesIO(data)
                    buffer.seek(0)
                    download_name = f"Invoice_{timestamp}.xlsx"
                    return send_file(
                        buffer,
                        as_attachment=True,
                        download_name=download_name,
                        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    return render_template(
        "index.html",
        error=error,
        template_path=app.config["TEMPLATE_PATH"],
    )


if __name__ == "__main__":
    app.run(debug=True)
