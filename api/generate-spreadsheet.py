from fastapi import FastAPI, Request
from fastapi.responses import FileResponse
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
import tempfile
import uuid
import os

app = FastAPI()

@app.post("/generate-spreadsheet")
async def generate_spreadsheet(request: Request):
    data = await request.json()

    calorie_target = data["calorie_target"]
    protein_target = data["protein_target"]
    days = data["days"]

    wb = Workbook()
    for day, meals in days.items():
        ws = wb.create_sheet(title=day)

        # Summary row at top
        ws.append(["Metric", "Target", "Actual", "Difference"])
        ws.append(["Calories (kcal)", calorie_target, "", ""])
        ws.append(["Protein (grams)", protein_target, "", ""])
        ws.append([])

        total_day_cal = 0
        total_day_protein = 0

        for meal in meals:
            ws.append([meal["meal_name"]])
            ws.append(["Meal", "Ingredient", "Quantity", "Unit", "Protein (g)", "Calories"])

            meal_cal = 0
            meal_protein = 0
            for ing in meal["ingredients"]:
                ws.append([
                    meal["meal_name"],
                    ing["name"],
                    ing["quantity"],
                    ing["unit"],
                    ing["protein"],
                    ing["calories"]
                ])
                meal_cal += ing["calories"]
                meal_protein += ing["protein"]

            ws.append(["TOTAL", "", "", "", meal_protein, meal_cal])
            ws.append([])
            total_day_cal += meal_cal
            total_day_protein += meal_protein

        ws.append(["Daily Totals", "", "", "", total_day_protein, total_day_cal])
        ws["C2"] = total_day_cal
        ws["D2"] = total_day_cal - calorie_target
        ws["C3"] = total_day_protein
        ws["D3"] = total_day_protein - protein_target

        # Optional formatting
        for row in ws.iter_rows(min_row=1, max_row=3):
            for cell in row:
                cell.font = Font(bold=True)
        for row in ws.iter_rows(min_row=1, max_col=4, max_row=3):
            for cell in row:
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, f"meal_plan_{uuid.uuid4()}.xlsx")
    wb.save(file_path)
    return FileResponse(file_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename="meal_plan.xlsx")
