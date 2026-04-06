"""
TaskFlow – Flask Web Application
──────────────────────────────────
Browser-based GUI for the TaskFlow scheduler.
"""

from __future__ import annotations
import os, json, tempfile, datetime as dt
from pathlib import Path
from flask import (
    Flask, render_template, request, jsonify, send_file, session,
)
from scheduler import (
    load_yaml, schedule, export_gantt_xlsx, export_gantt_csv,
    export_ics, generate_calendar_links, ScheduleResult,
)

app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = os.urandom(24)

UPLOAD_DIR = Path(tempfile.mkdtemp())
_results_cache: dict[str, ScheduleResult] = {}
_last_raw_data: dict | None = None


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/schedule", methods=["POST"])
def api_schedule():
    """Accept a YAML file, run the scheduler, return JSON results."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename"}), 400

    yaml_path = UPLOAD_DIR / f.filename
    f.save(yaml_path)

    try:
        data = load_yaml(str(yaml_path))
        global _last_raw_data
        _last_raw_data = {
            "start_date": data["start"].strftime("%m/%d/%Y"),
            "end_date": data["end"].strftime("%m/%d/%Y"),
            "maximum_tasks": data["max_tasks"],
            "members": [{"name": m.name, "areas": m.areas,
                          "dates_unavailable": [[s.isoformat(), e.isoformat()] for s, e in m.unavailable]}
                         for m in data["members"]],
            "tasks": [{"id": t.id, "description": t.description, "areas": t.areas,
                       "dependencies": t.dependencies, "estimated_days": t.estimated_days}
                      for t in data["tasks"]],
        }
        result = schedule(data)

        sid = os.urandom(8).hex()
        _results_cache[sid] = result

        tasks = []
        for t in sorted(result.tasks, key=lambda t: (t.start_date or dt.date.max)):
            preds = list(result.graph.predecessors(t.id))
            succs = list(result.graph.successors(t.id))
            tasks.append({
                "id": t.id,
                "description": t.description,
                "areas": t.areas,
                "dependencies": t.dependencies,
                "assigned_to": t.assigned_to or "Unassigned",
                "start_date": t.start_date.isoformat() if t.start_date else None,
                "end_date": t.end_date.isoformat() if t.end_date else None,
                "estimated_days": t.estimated_days,
                "predecessors": preds,
                "successors": succs,
            })

        members = []
        palette = ["#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F",
                    "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC"]
        for i, m in enumerate(result.members):
            members.append({
                "name": m.name,
                "areas": m.areas,
                "color": palette[i % len(palette)],
            })
        member_colors = {m["name"]: m["color"] for m in members}
        for t in tasks:
            t["color"] = member_colors.get(t["assigned_to"], "#999999")

        cal_links = generate_calendar_links(result)
        cal_data = []
        for lnk in cal_links:
            cal_data.append({
                "task_id": lnk["task_id"],
                "description": lnk["description"],
                "assigned_to": lnk["assigned_to"],
                "start": lnk["start"].isoformat(),
                "end": lnk["end"].isoformat(),
                "google_url": lnk["google_url"],
                "outlook_url": lnk["outlook_url"],
            })

        return jsonify({
            "session_id": sid,
            "tasks": tasks,
            "members": members,
            "topo_order": result.topo_order,
            "warnings": result.warnings,
            "project_start": result.project_start.isoformat(),
            "project_end": result.project_end.isoformat(),
            "calendar_links": cal_data,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/export/<fmt>")
def api_export(fmt):
    """Export the last schedule result as xlsx, csv, or ics."""
    sid = request.args.get("sid", "")
    result = _results_cache.get(sid)
    if not result:
        return jsonify({"error": "No schedule found. Run the scheduler first."}), 404

    ext_map = {"xlsx": ".xlsx", "csv": ".csv", "ics": ".ics"}
    mime_map = {
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "csv": "text/csv",
        "ics": "text/calendar",
    }
    if fmt not in ext_map:
        return jsonify({"error": f"Unknown format: {fmt}"}), 400

    out_path = UPLOAD_DIR / f"taskflow_export{ext_map[fmt]}"
    if fmt == "xlsx":
        export_gantt_xlsx(result, str(out_path))
    elif fmt == "csv":
        export_gantt_csv(result, str(out_path))
    elif fmt == "ics":
        alarm = int(request.args.get("alarm", 30))
        export_ics(result, str(out_path), alarm_minutes=alarm)

    return send_file(
        str(out_path),
        mimetype=mime_map[fmt],
        as_attachment=True,
        download_name=f"taskflow_schedule{ext_map[fmt]}",
    )


@app.route("/api/raw_data")
def api_raw_data():
    if not _last_raw_data:
        return jsonify({"error": "No YAML loaded yet. Upload and run a file first."}), 404
    return jsonify(_last_raw_data)


def run_webapp(port=5000, debug=False):
    """Launch the Flask web app."""
    import webbrowser, threading
    url = f"http://127.0.0.1:{port}"
    print(f"\n  ⬡  TaskFlow Scheduler")
    print(f"  ────────────────────────────")
    print(f"  Open in browser: {url}")
    print(f"  Press Ctrl+C to stop\n")
    if not debug:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=debug)


if __name__ == "__main__":
    run_webapp()
