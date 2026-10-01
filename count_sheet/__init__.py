"""BL Tools — ใบนับสต๊อก P&G (Cycle Count) + ชนสต็อก + นับซ้ำครั้งที่ 2/3 เฉพาะบรรทัดที่ Diff

วิธีเชื่อมกับ app.py:
    from count_sheet import bp as count_sheet_bp
    app.register_blueprint(count_sheet_bp)
แล้วเปิดหน้า /count-sheet
"""
from datetime import date, datetime
import io
import zipfile

from flask import Blueprint, render_template, request, send_file

from .generator import CountSheetError, generate, load_plan, process_counted

bp = Blueprint("count_sheet", __name__, url_prefix="/count-sheet",
               template_folder="templates")

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _plan_days():
    try:
        plan = load_plan()
        return sorted(plan), {d: len(s) for d, s in plan.items()}
    except Exception:
        return list(range(1, 31)), {}


def _check_xlsx(f, label):
    if not f or not f.filename:
        raise CountSheetError(f"ยังไม่ได้เลือกไฟล์ {label}")
    if not f.filename.lower().endswith(".xlsx"):
        raise CountSheetError(f"ไฟล์ {f.filename} ไม่ใช่ .xlsx")


def _friendly(e):
    if isinstance(e, (zipfile.BadZipFile, KeyError)) or "zip" in str(e).lower():
        return "เปิดไฟล์ไม่ได้ — ไฟล์เสีย หรือไม่ใช่ Excel (.xlsx) จริง ลอง Save As เป็น .xlsx ใหม่"
    return f"อ่านไฟล์ไม่ได้: {e}"


def _page(form=None, error=None, error_in=None):
    days, sku_count = _plan_days()
    form = form or {"day": 1, "date": date.today().strftime("%Y-%m-%d")}
    return render_template("count_sheet/index.html", days=days, sku_count=sku_count,
                           form=form, error=error, error_in=error_in)


@bp.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return _page()

    form = {"day": request.form.get("day", "1"), "date": request.form.get("date", "")}
    try:
        _check_xlsx(request.files.get("stock_mc"), "Stock MC")
        _check_xlsx(request.files.get("stock_fo"), "Stock FO")
        cycle = request.files.get("cycle")
        if cycle and cycle.filename:
            _check_xlsx(cycle, "Cycle")
        try:
            day = int(form["day"])
            count_date = datetime.strptime(form["date"], "%Y-%m-%d").date()
        except ValueError:
            raise CountSheetError("กรอก Day หรือวันที่นับไม่ถูกต้อง")
        data, info = generate(
            day=day, count_date=count_date,
            mc_file=io.BytesIO(request.files["stock_mc"].read()),
            fo_file=io.BytesIO(request.files["stock_fo"].read()),
            cycle_file=io.BytesIO(cycle.read()) if cycle and cycle.filename else None)
        name = f"ใบนับ_PG_{count_date:%d-%m-%Y}_Day{day}_ครั้งที่1.xlsx"
        return send_file(io.BytesIO(data), as_attachment=True, download_name=name, mimetype=XLSX)
    except CountSheetError as e:
        return _page(form, str(e), "new")
    except Exception as e:
        return _page(form, _friendly(e), "new")


@bp.route("/result", methods=["POST"])
def result():
    f = request.files.get("counted")
    try:
        _check_xlsx(f, "ใบนับที่กรอกยอดแล้ว")
        data, info = process_counted(io.BytesIO(f.read()))
        d = info["date"].replace("/", "-")
        if info["forms"]:
            parts = "_".join(f"{f['zone']}{(f['group'] or '').replace(' ', '')}ครั้งที่{f['round']}"
                             for f in info["forms"])
            name = f"ใบนับ_PG_{d}_Day{info['day']}_{parts}.xlsx"
        else:
            name = f"สรุปผลนับ_PG_{d}_Day{info['day']}.xlsx"
        return send_file(io.BytesIO(data), as_attachment=True, download_name=name, mimetype=XLSX)
    except CountSheetError as e:
        return _page(error=str(e), error_in="result")
    except Exception as e:
        return _page(error=_friendly(e), error_in="result")
