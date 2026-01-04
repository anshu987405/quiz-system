from flask import Flask, request, render_template, send_file, jsonify, redirect
from openpyxl import Workbook, load_workbook
import os, datetime

app = Flask(__name__)

# ================= CONFIG =================
ADMIN_USER = "admin"
ADMIN_PASS = "admin@123"

QUIZ_FILE = "students.xlsx"
STUDENT_MASTER = "students_master.xlsx"
PAYMENT_FILE = "fees_payments.xlsx"

# ================= CREATE FILES (SAFE) =================
def ensure_files():
    # Quiz file
    if not os.path.exists(QUIZ_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["Name","Score","Correct","Total","Date"])
        wb.save(QUIZ_FILE)

    # Student master
    if not os.path.exists(STUDENT_MASTER):
        wb = Workbook()
        ws = wb.active
        ws.append([
            "StudentID","Name","Father","Course",
            "Mobile","Address","TotalFees","Status"
        ])
        wb.save(STUDENT_MASTER)

    # Payment history
    if not os.path.exists(PAYMENT_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["StudentID","PaidAmount","Mode","Date"])
        wb.save(PAYMENT_FILE)

ensure_files()

# ================= HELPERS =================
def generate_student_id():
    wb = load_workbook(STUDENT_MASTER)
    ws = wb.active
    return "STU2026" + str(ws.max_row).zfill(3)

def get_status(row):
    return row[7] if len(row) > 7 else "ACTIVE"

def get_total_paid(student_id):
    wb = load_workbook(PAYMENT_FILE)
    ws = wb.active
    total = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] == student_id:
            total += r[1]
    return total

# ================= QUIZ SUBMIT =================
@app.route("/submit", methods=["POST"])
def submit_quiz():
    data = request.get_json()
    wb = load_workbook(QUIZ_FILE)
    ws = wb.active
    ws.append([
        data["name"],
        data["score"],
        data["correct"],
        data["total"],
        datetime.datetime.now().strftime("%d-%m-%Y %H:%M")
    ])
    wb.save(QUIZ_FILE)
    return jsonify({"status":"saved"})

# ================= ADMIN LOGIN =================
@app.route("/admin", methods=["GET","POST"])
def admin():
    if request.method == "POST":
        if request.form["username"] == ADMIN_USER and request.form["password"] == ADMIN_PASS:
            return redirect("/dashboard")
    return render_template("admin_login.html")

# ================= DASHBOARD (FIXED) =================
@app.route("/dashboard")
def dashboard():
    wb = load_workbook(STUDENT_MASTER)
    ws = wb.active

    active_students = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if get_status(r) == "ACTIVE":
            active_students.append(r)

    total_students = len(active_students)
    total_paid = sum(get_total_paid(r[0]) for r in active_students)
    total_due = sum(r[6] - get_total_paid(r[0]) for r in active_students)

    return render_template(
        "admin_dashboard.html",
        total_students=total_students,
        total_paid=total_paid,
        total_due=total_due
    )

# ================= ADD STUDENT =================
@app.route("/add-student")
def add_student():
    return render_template("student_form.html")


@app.route("/save-student", methods=["POST"])
def save_student():
    d = request.form
    sid = generate_student_id()

    total_fees = int(d["total_fees"])
    initial_paid = int(d.get("initial_paid", 0))

    # 1️⃣ Save student
    wb = load_workbook(STUDENT_MASTER)
    ws = wb.active
    ws.append([
        sid,
        d["name"],
        d["father"],
        d["course"],
        d["mobile"],
        d["address"],
        total_fees,
        "ACTIVE"
    ])
    wb.save(STUDENT_MASTER)

    # 2️⃣ Save initial payment
    if initial_paid > 0:
        wb2 = load_workbook(PAYMENT_FILE)
        ws2 = wb2.active
        ws2.append([
            sid,
            initial_paid,
            "Initial",
            datetime.datetime.now().strftime("%d-%m-%Y %H:%M")
        ])
        wb2.save(PAYMENT_FILE)

    return redirect("/students")



# ================= STUDENT LIST =================
@app.route("/students")
def students():
    wb = load_workbook(STUDENT_MASTER)
    ws = wb.active

    data = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if get_status(r) == "ACTIVE":
            paid = get_total_paid(r[0])
            due = r[6] - paid
            data.append(r + (paid, due))

    return render_template("student_list.html", data=data)

# ================= SOFT DELETE =================
@app.route("/delete-student/<sid>")
def delete_student(sid):
    wb = load_workbook(STUDENT_MASTER)
    ws = wb.active

    for row in ws.iter_rows(min_row=2):
        if row[0].value == sid:
            # agar Status column exist karta hai
            if len(row) >= 8:
                row[7].value = "DELETED"
            else:
                # agar purani file hai (no status column)
                ws.cell(row=row[0].row, column=8).value = "DELETED"
            break

    wb.save(STUDENT_MASTER)
    return redirect("/students")


# ================= ADD PAYMENT =================
@app.route("/print-receipt/<sid>")
def print_receipt(sid):
    # ================= STUDENT DETAILS =================
    wb = load_workbook(STUDENT_MASTER)
    ws = wb.active

    student = None
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] == sid:
            student = r
            break

    if not student:
        return "Student not found"

    # ================= PAYMENT CALCULATION =================
    total_paid = get_total_paid(sid)
    due = student[6] - total_paid

    # ================= RENDER RECEIPT =================
    return render_template(
        "fee_receipt.html",
        sid=sid,
        name=student[1],
        course=student[3],
        mobile=student[4],
        paid_amount=total_paid,   # total paid till now
        total_fees=student[6],
        total_paid=total_paid,
        due=due
    )


# ================= FEES DUE REPORT =================
@app.route("/fees-due")
def fees_due():
    wb = load_workbook(STUDENT_MASTER)
    ws = wb.active

    data = []

    for r in ws.iter_rows(min_row=2, values_only=True):

        # 🔐 Status safe check (old files support)
        status = r[7] if len(r) > 7 else "ACTIVE"
        if status != "ACTIVE":
            continue

        student_id = r[0]
        total_fees = int(r[6])

        paid = get_total_paid(student_id)
        due = total_fees - paid

        # ❗ Sirf jinke due > 0
        if due > 0:
            data.append((
                student_id,   # 0
                r[1],          # name
                r[3],          # course
                r[4],          # mobile
                total_fees,    # total
                paid,          # paid
                due            # due
            ))

    return render_template("fee_due.html", data=data)


# ================= RUN =================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

