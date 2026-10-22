"""Builds the course sandbox spreadsheet (anonymized, fictional data)."""
import random
from datetime import datetime, timedelta, date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

random.seed(2026)
OUT = "assets/sandbox.xlsx"  # run from the repo root

NAVY = "1F2A44"
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)

first = ["Anna", "Boris", "Vera", "Gleb", "Daria", "Egor", "Jeanne", "Zakhar", "Irina", "Kirill",
         "Liana", "Mark", "Nina", "Oleg", "Polina", "Roman", "Sofia", "Timur", "Ulyana", "Fedor",
         "Alex", "Maya", "Nick", "Eva", "Leo", "Olga", "Sergey", "Tina", "Julia", "Yan"]
last_initial = list("ABVGDEZIKLMNOPRSTFHW")
countries = ["Georgia", "Armenia", "Kazakhstan", "Serbia", "Poland", "Uzbekistan"]

# ---- Employees ----
employees = []
supervisors = []
for i in range(6):
    name = f"{random.choice(first)} {random.choice(last_initial)}."
    supervisors.append({"id": f"S-{i+1:02d}", "name": name})
for i in range(30):
    sup = supervisors[i % len(supervisors)]
    employees.append({
        "id": f"D-{i+1:03d}",
        "name": f"{first[i]} {random.choice(last_initial)}.",
        "role": "Drafter",
        "team": random.choice(["FP", "ESX", "FP", "ESX", "FF"]),
        "country": random.choice(countries),
        "sup": sup["name"],
        "email": f"drafter{i+1:03d}@example.com",
        "start": date(2023, 1, 1) + timedelta(days=random.randint(0, 1100)),
    })

wb = Workbook()

def style_header(ws, widths):
    for c, w in enumerate(widths, start=1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"

# ---- Help ----
ws = wb.active
ws.title = "Help"
info = [
    ["Course sandbox: Google Sheets Scripts"],
    [""],
    ["Before you start: File → Make a copy. Run every course script on your copy."],
    ["All names, emails and numbers are fictional. Any resemblance to real people is coincidental."],
    [""],
    ["Sheet", "What's inside", "Used in"],
    ["Orders", "200 orders over 3 months: dates, deadlines, statuses, assignees, hours, QC scores", "Modules 2, 3, 5, 6, 8"],
    ["Employees", "30 drafters: team, country, supervisor, fictional email", "Modules 5, 6, 7"],
    ["Hours", "Worked and logged hours by week", "Modules 6, 8"],
    ["Export", "A messy export: extra spaces, mixed case, duplicates, empty rows, dates as text", "Modules 2, 6"],
    ["Archive", "Empty sheet for closed orders", "Module 5"],
    ["Summary", "Empty sheet for metrics and charts", "Module 8"],
    ["Settings", "Parameters the scripts read: sheet names, thresholds, addresses", "Modules 1, 5"],
]
for r in info:
    ws.append(r)
ws["A1"].font = Font(bold=True, size=14, name="Arial")
for c in range(1, 4):
    ws.cell(row=6, column=c).font = HEADER_FONT
    ws.cell(row=6, column=c).fill = HEADER_FILL
ws.column_dimensions["A"].width = 16
ws.column_dimensions["B"].width = 80
ws.column_dimensions["C"].width = 24

# ---- Orders ----
ws = wb.create_sheet("Orders")
headers = ["Order ID", "Created", "Deadline", "Type", "Status", "Assignee ID", "Assignee",
           "Supervisor", "Hours", "QC Score", "Delivered", "Comment"]
ws.append(headers)
now = datetime(2026, 10, 19, 12, 0)
start = datetime(2026, 7, 20, 8, 0)
comments = ["", "", "", "Client asked to rush", "More photos needed", "Second review", "", "Complex floor plan", ""]
orders = []
for i in range(200):
    created = start + timedelta(minutes=random.randint(0, int((now - start).total_seconds() // 60) - 600))
    typ = random.choices(["FP", "ESX", "FF"], weights=[5, 4, 1])[0]
    deadline = created + timedelta(hours={"FP": 24, "ESX": 48, "FF": 72}[typ])
    emp = random.choice([e for e in employees if e["team"] == typ] or employees)
    hours = round(max(0.5, random.gauss({"FP": 3, "ESX": 5, "FF": 6}[typ], 1.2)), 1)
    if random.random() < 0.02:
        hours = round(hours * random.uniform(3, 5), 1)  # outliers for module 8
    age_h = (now - created).total_seconds() / 3600
    if age_h > 240:
        status = random.choices(["Done", "Problem"], weights=[95, 5])[0]
    elif age_h > 30:
        status = random.choices(["Done", "In Review", "In Progress", "Problem"], weights=[50, 15, 25, 10])[0]
    else:
        status = random.choice(["New", "In Progress", "In Review"])
    done = None
    qc = None
    if status == "Done":
        late = random.random() < 0.12
        done = deadline + timedelta(hours=random.uniform(1, 20)) if late else created + timedelta(
            hours=random.uniform(2, (deadline - created).total_seconds() / 3600 - 0.5))
        qc = random.choices([5, 4, 3, 2], weights=[45, 35, 15, 5])[0]
        done = min(done, now - timedelta(hours=10))  # never after the latest order
    orders.append([f"ORD-{10001 + i}", created, deadline, typ, status, emp["id"], emp["name"], emp["sup"],
                   hours, qc, done, random.choice(comments)])
orders.sort(key=lambda r: r[1])
for r in orders:
    ws.append(r)
style_header(ws, [12, 17, 17, 6, 13, 14, 14, 14, 7, 10, 17, 26])
for row in ws.iter_rows(min_row=2):
    for idx in (1, 2, 10):
        row[idx].number_format = "dd.mm.yyyy hh:mm"

# ---- Employees ----
ws = wb.create_sheet("Employees")
ws.append(["ID", "Name", "Role", "Team", "Country", "Supervisor", "Email", "Start Date"])
for e in employees:
    ws.append([e["id"], e["name"], e["role"], e["team"], e["country"], e["sup"], e["email"], e["start"]])
style_header(ws, [8, 14, 13, 9, 12, 14, 24, 13])
for row in ws.iter_rows(min_row=2):
    row[7].number_format = "dd.mm.yyyy"

# ---- Hours ----
ws = wb.create_sheet("Hours")
ws.append(["Week Of", "ID", "Name", "Worked Hours", "Logged Hours"])
week0 = date(2026, 8, 24)
for w in range(8):
    wk = week0 + timedelta(weeks=w)
    for e in employees:
        worked = round(random.gauss(38, 4), 1)
        logged = round(worked + random.gauss(0, 0.8), 1)
        if random.random() < 0.03:
            logged = round(worked + random.uniform(5, 12), 1)  # mismatches to catch
        ws.append([wk, e["id"], e["name"], worked, logged])
style_header(ws, [12, 8, 14, 14, 20])
for row in ws.iter_rows(min_row=2):
    row[0].number_format = "dd.mm.yyyy"

# ---- Export (messy) ----
ws = wb.create_sheet("Export")
ws.append(["id", "assignee", "status", "date", "hours"])
sample = random.sample(orders, 40)
messy = []
for r in sample:
    name = r[6]
    st = r[4]
    d = r[1].strftime(random.choice(["%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"]))
    variant = random.random()
    if variant < 0.25:
        name = "  " + name + " "
    elif variant < 0.4:
        name = name.upper()
    if random.random() < 0.3:
        st = random.choice([st.lower(), " " + st, st.upper()])
    messy.append([r[0], name, st, d, str(r[8]).replace(".", ",")])
messy += [list(m) for m in random.sample(messy, 5)]  # duplicates
random.shuffle(messy)
for i, m in enumerate(messy):
    ws.append(m)
    if i % 13 == 7:
        ws.append([None, None, None, None, None])  # empty rows
for c, w in enumerate([12, 18, 14, 12, 8], start=1):
    ws.column_dimensions[get_column_letter(c)].width = w

# ---- Archive, Summary ----
ws = wb.create_sheet("Archive")
ws.append(headers + ["Archived At"])
style_header(ws, [12, 17, 17, 6, 13, 14, 14, 14, 7, 10, 17, 26, 17])
ws = wb.create_sheet("Summary")
ws["A1"] = "Module 8 scripts write metrics and charts here"

# ---- Settings ----
ws = wb.create_sheet("Settings")
ws.append(["Parameter", "Value", "Description"])
for r in [
    ["ORDERS_SHEET", "Orders", "Sheet the scripts read orders from"],
    ["ARCHIVE_SHEET", "Archive", "Sheet closed orders are moved to"],
    ["HOT_HOURS", 4, "Hours before the deadline when an order counts as urgent"],
    ["NOTIFY_EMAIL", "", "Your email for notifications (module 5)"],
    ["SLACK_WEBHOOK", "", "Leave empty: the webhook URL lives in script properties (module 5)"],
    ["TIME_ZONE", "Asia/Tbilisi", "Time zone for dates in reports"],
]:
    ws.append(r)
style_header(ws, [22, 22, 56])

wb.save(OUT)
print("saved", OUT, "orders", len(orders), "messy", len(messy))
