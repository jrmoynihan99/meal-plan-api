import json
import tempfile
import uuid
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

def handler(request):
    # Handle CORS preflight
    if request.method == 'OPTIONS':
        return {
            'statusCode': 200,
            'headers': {
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Methods': 'POST, OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type',
            },
            'body': ''
        }
    
    # Only allow POST requests
    if request.method != 'POST':
        return {
            'statusCode': 405,
            'headers': {
                'Access-Control-Allow-Origin': '*',
            },
            'body': json.dumps({'error': 'Method not allowed'})
        }
    
    try:
        # Parse request body
        if hasattr(request, 'get_json'):
            data = request.get_json()
        else:
            # Handle different request body formats
            body = getattr(request, 'body', None) or getattr(request, 'data', None)
            if body:
                if isinstance(body, bytes):
                    body = body.decode('utf-8')
                data = json.loads(body)
            else:
                return {
                    'statusCode': 400,
                    'headers': {'Access-Control-Allow-Origin': '*'},
                    'body': json.dumps({'error': 'No data provided'})
                }
        
        calorie_target = data["calorie_target"]
        protein_target = data["protein_target"]
        days = data["days"]
        
        # Create workbook
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
            
            # Update summary
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
        
        # Save to temporary file
        temp_dir = tempfile.gettempdir()
        file_path = os.path.join(temp_dir, f"meal_plan_{uuid.uuid4()}.xlsx")
        wb.save(file_path)
        
        # Read file content and encode as base64
        with open(file_path, 'rb') as f:
            file_content = f.read()
        
        # Clean up
        os.remove(file_path)
        
        # Return file as base64 encoded response
        import base64
        file_b64 = base64.b64encode(file_content).decode('utf-8')
        
        return {
            'statusCode': 200,
            'headers': {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': 'attachment; filename="meal_plan.xlsx"',
                'Access-Control-Allow-Origin': '*',
            },
            'body': file_b64,
            'isBase64Encoded': True
        }
        
    except Exception as e:
        return {
            'statusCode': 500,
            'headers': {'Access-Control-Allow-Origin': '*'},
            'body': json.dumps({'error': str(e)})
        }
