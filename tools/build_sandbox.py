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

first = ["Анна", "Борис", "Вера", "Глеб", "Дарья", "Егор", "Жанна", "Захар", "Ирина", "Кирилл",
         "Лиана", "Марк", "Нина", "Олег", "Полина", "Роман", "Софья", "Тимур", "Ульяна", "Фёдор",
         "Алекс", "Майя", "Ник", "Ева", "Лев", "Ольга", "Сергей", "Тина", "Юлия", "Ян"]
last_initial = list("АБВГДЕЗИКЛМНОПРСТФХШ")
countries = ["Грузия", "Армения", "Казахстан", "Сербия", "Польша", "Узбекистан"]

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
        "role": "Составитель",
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

# ---- Справка ----
ws = wb.active
ws.title = "Справка"
info = [
    ["Учебная таблица курса «Скрипты для Google Таблиц»"],
    [""],
    ["Перед работой: Файл → Создать копию. Все скрипты курса запускаются на копии."],
    ["Все имена, email и цифры вымышлены. Совпадения с реальными людьми случайны."],
    [""],
    ["Лист", "Что внутри", "Где используется"],
    ["Заказы", "200 заказов за 3 месяца: даты, дедлайны, статусы, исполнители, часы, оценки QC", "Модули 2, 3, 5, 6, 8"],
    ["Сотрудники", "30 составителей: команда, страна, супервайзер, вымышленный email", "Модули 5, 6, 7"],
    ["Часы", "Рабочие и залогированные часы по неделям", "Модули 6, 8"],
    ["Выгрузка", "«Грязная» выгрузка: пробелы, регистр, дубли, пустые строки, даты текстом", "Модули 2, 6"],
    ["Архив", "Пустой лист для закрытых заказов", "Модуль 5"],
    ["Сводка", "Пустой лист для метрик и графиков", "Модуль 8"],
    ["Настройки", "Параметры, которые читают скрипты: имена листов, пороги, адреса", "Модули 1, 5"],
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

# ---- Заказы ----
ws = wb.create_sheet("Заказы")
headers = ["ID заказа", "Создан", "Дедлайн", "Тип", "Статус", "ID исполнителя", "Исполнитель",
           "Супервайзер", "Часы", "Оценка QC", "Сдан", "Комментарий"]
ws.append(headers)
now = datetime(2026, 10, 19, 12, 0)
start = datetime(2026, 7, 20, 8, 0)
comments = ["", "", "", "Клиент просил ускорить", "Нужны доп. фото", "Повторная проверка", "", "Сложный план этажа", ""]
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
        status = random.choices(["Готов", "Проблема"], weights=[95, 5])[0]
    elif age_h > 30:
        status = random.choices(["Готов", "На проверке", "В работе", "Проблема"], weights=[50, 15, 25, 10])[0]
    else:
        status = random.choice(["Новый", "В работе", "На проверке"])
    done = None
    qc = None
    if status == "Готов":
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

# ---- Сотрудники ----
ws = wb.create_sheet("Сотрудники")
ws.append(["ID", "Имя", "Роль", "Команда", "Страна", "Супервайзер", "Email", "Дата начала"])
for e in employees:
    ws.append([e["id"], e["name"], e["role"], e["team"], e["country"], e["sup"], e["email"], e["start"]])
style_header(ws, [8, 14, 13, 9, 12, 14, 24, 13])
for row in ws.iter_rows(min_row=2):
    row[7].number_format = "dd.mm.yyyy"

# ---- Часы ----
ws = wb.create_sheet("Часы")
ws.append(["Неделя с", "ID", "Имя", "Рабочие часы", "Залогированные часы"])
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

# ---- Выгрузка (messy) ----
ws = wb.create_sheet("Выгрузка")
ws.append(["id", "исполнитель", "статус", "дата", "часы"])
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

# ---- Архив, Сводка ----
ws = wb.create_sheet("Архив")
ws.append(headers + ["Дата архивации"])
style_header(ws, [12, 17, 17, 6, 13, 14, 14, 14, 7, 10, 17, 26, 17])
ws = wb.create_sheet("Сводка")
ws["A1"] = "Сюда скрипты модуля 8 выводят метрики и графики"

# ---- Настройки ----
ws = wb.create_sheet("Настройки")
ws.append(["Параметр", "Значение", "Описание"])
for r in [
    ["ЛИСТ_ЗАКАЗОВ", "Заказы", "Откуда скрипты берут заказы"],
    ["ЛИСТ_АРХИВА", "Архив", "Куда переносятся закрытые заказы"],
    ["ПОРОГ_ПРОСРОЧКИ_Ч", 4, "За сколько часов до дедлайна заказ считается «горящим»"],
    ["EMAIL_УВЕДОМЛЕНИЙ", "", "Ваш email для уведомлений (модуль 5)"],
    ["SLACK_WEBHOOK", "", "Адрес входящего webhook Slack (модуль 5, по желанию)"],
    ["ЧАСОВОЙ_ПОЯС", "Asia/Tbilisi", "Часовой пояс для дат в отчётах"],
]:
    ws.append(r)
style_header(ws, [22, 22, 56])

wb.save(OUT)
print("saved", OUT, "orders", len(orders), "messy", len(messy))
