import openpyxl

water_sales_workbook = openpyxl.load_workbook("watersales.xlsx")

water_sales_worksheet = water_sales_workbook["Sheet1"]

for row in range(2, water_sales_worksheet.max_row + 2):
    value_a = water_sales_worksheet.cell(row=row, column=6).value
    value_b = water_sales_worksheet.cell(row=row, column=11).value

    if value_a is None and value_b is None:
        continue

    result = value_a * value_b

    water_sales_worksheet.cell(row=row, column=12).value = result

water_sales_worksheet.cell(row=1, column=12).value = "February 2023 Billing"

water_sales_workbook.save("watersales_1.xlsx")
