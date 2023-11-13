import openpyxl

# Створити новий файл Excel
workbook = openpyxl.Workbook()

# Вибрати аркуш за замовчуванням
sheet = workbook.active

# Заголовки таблиці
header = ['Прізвище', 'Ім\'я', 'Місяць 1', 'Місяць 2', 'Місяць 3', 'Податок', 'Сума до видачі']

# Встановити заголовки в перший рядок
sheet.append(header)

# Дані про співробітників
employees_data = [
    ['Іванов', 'Іван', 1000, 1200, 900],
    ['Петров', 'Петро', 800, 950, 1100],
    ['Сидоров', 'Сидір', 1200, 1000, 950],
    ['Коваленко', 'Марія', 1100, 1050, 1150]
]

# Додати дані про співробітників до таблиці
for employee in employees_data:
    # Розрахунок податку за три місяці
    tax = sum([0.13 * (income - (income * 0.02) - (income * 0.005) - income * 0.005 - 71.1) for income in employee[2:5]])

    # Сума до видачі
    total_salary = sum(employee[2:5]) - tax

    # Додати рядок з даними до таблиці
    sheet.append(employee + [tax, total_salary])

# Зберегти файл Excel
workbook.save('розрахунок_заробітної_плати.xlsx')
