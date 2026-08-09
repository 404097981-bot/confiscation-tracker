import sqlite3
import os
from datetime import datetime, timedelta
from io import BytesIO

from flask import Flask, request, jsonify, send_file, send_from_directory
from openpyxl import Workbook

app = Flask(__name__, static_folder="static", static_url_path="")

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.db")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_name TEXT NOT NULL,
                student_id TEXT NOT NULL,
                item_name TEXT NOT NULL,
                reason TEXT NOT NULL DEFAULT '',
                confiscated_at TEXT NOT NULL,
                return_date TEXT NOT NULL DEFAULT '',
                return_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'active',
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            )
        """)
        try:
            conn.execute("ALTER TABLE records ADD COLUMN return_date TEXT NOT NULL DEFAULT ''")
        except:
            pass
        try:
            conn.execute("ALTER TABLE records ADD COLUMN reason TEXT NOT NULL DEFAULT ''")
        except:
            pass


init_db()


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/records", methods=["GET"])
def list_records():
    status = request.args.get("status")
    with get_db() as conn:
        if status:
            rows = conn.execute("SELECT * FROM records WHERE status=? ORDER BY return_at ASC", (status,)).fetchall()
        else:
            rows = conn.execute("SELECT * FROM records ORDER BY return_at ASC").fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/records", methods=["POST"])
def create_record():
    data = request.get_json()
    if not data:
        return jsonify({"error": "无效数据"}), 400

    student_name = data.get("student_name", "").strip()
    student_id = data.get("student_id", "").strip()
    item_name = data.get("item_name", "").strip()
    reason = data.get("reason", "").strip()
    confiscated_at = data.get("confiscated_at", "").strip()
    return_date = data.get("return_date", "").strip()

    if not student_name or not student_id or not item_name:
        return jsonify({"error": "请填写完整信息"}), 400
    if not confiscated_at or not return_date:
        return jsonify({"error": "请选择日期"}), 400

    return_at = return_date + "T17:40:00+08:00"

    with get_db() as conn:
        cursor = conn.execute(
            "INSERT INTO records (student_name, student_id, item_name, reason, confiscated_at, return_date, return_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (student_name, student_id, item_name, reason, confiscated_at, return_date, return_at)
        )
        record = conn.execute("SELECT * FROM records WHERE id=?", (cursor.lastrowid,)).fetchone()

    return jsonify(dict(record)), 201


@app.route("/api/records/<int:record_id>/return", methods=["PUT"])
def return_record(record_id):
    with get_db() as conn:
        conn.execute("UPDATE records SET status='returned' WHERE id=?", (record_id,))
        if conn.total_changes == 0:
            return jsonify({"error": "记录不存在"}), 404
    return jsonify({"ok": True})


@app.route("/api/records/<int:record_id>/adjust", methods=["PUT"])
def adjust_record(record_id):
    data = request.get_json()
    delta = data.get("delta", 0)
    if delta not in (-1, 1):
        return jsonify({"error": "无效的调整值"}), 400

    with get_db() as conn:
        row = conn.execute("SELECT * FROM records WHERE id=? AND status='active'", (record_id,)).fetchone()
        if not row:
            return jsonify({"error": "记录不存在或已归还"}), 404

        from datetime import datetime as dt
        try:
            return_date_str = row["return_date"] if row["return_date"] else row["return_at"][:10]
            new_return_date = dt.strptime(return_date_str, "%Y-%m-%d") + timedelta(days=delta)
        except:
            return jsonify({"error": "日期解析失败"}), 400

        new_return_date_str = new_return_date.strftime("%Y-%m-%d")
        new_return_at = new_return_date_str + "T17:40:00+08:00"
        conn.execute("UPDATE records SET return_date=?, return_at=? WHERE id=?",
                     (new_return_date_str, new_return_at, record_id))

    return jsonify({"ok": True, "return_date": new_return_date_str, "return_at": new_return_at})


@app.route("/api/records/<int:record_id>", methods=["DELETE"])
def delete_record(record_id):
    with get_db() as conn:
        conn.execute("DELETE FROM records WHERE id=?", (record_id,))
        if conn.total_changes == 0:
            return jsonify({"error": "记录不存在"}), 404
    return jsonify({"ok": True})


@app.route("/api/export")
def export_xlsx():
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM records ORDER BY created_at DESC").fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "没收记录"

    headers = ["序号", "学生姓名", "学号", "物品名称", "没收原因", "没收日期", "归还日期", "归还截止时间", "状态"]
    ws.append(headers)

    status_map = {"active": "未归还", "returned": "已归还"}

    for i, row in enumerate(rows, 1):
        ws.append([
            i,
            row["student_name"],
            row["student_id"],
            row["item_name"],
            row["reason"],
            row["confiscated_at"],
            row["return_date"],
            row["return_at"],
            status_map.get(row["status"], row["status"])
        ])

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 20
    ws.column_dimensions["I"].width = 10

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"没收记录_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
