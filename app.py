import os
from flask import Flask, render_template, request, send_file, jsonify
from datetime import datetime, timedelta
import calendar
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

class DutyScheduler:
    def __init__(self):
        self.shifts = {
            'A': '06:00-14:00',
            'B': '14:00-22:00',
            'C': '22:00-06:00',
            'G': 'General',
            'R': 'Rest'
        }
        self.shift_rotation = {'A': 'C', 'C': 'B', 'B': 'A'}
        
    def get_day_name(self, year, month, day):
        return calendar.day_abbr[calendar.weekday(year, month, day)].upper()

    def get_next_shift(self, current_shift):
        return self.shift_rotation.get(current_shift, current_shift)

    def generate_schedule(self, employees_data, year, month):
        num_days = calendar.monthrange(year, month)[1]
        schedule = {}
        
        for emp in employees_data:
            schedule[emp['name']] = {
                'code': emp['code'],
                'post': emp['post'],
                'shifts': []
            }
            
            current_shift = emp['start_shift']
            rest_day = int(emp['rest_day'])  # Convert to int
            # Convert Sunday from 0 to 6
            rest_day = 6 if rest_day == 0 else rest_day - 1
            was_rest_day = False
            
            for day in range(1, num_days + 1):
                # Check if it's a rest day
                if calendar.weekday(year, month, day) == rest_day:
                    schedule[emp['name']]['shifts'].append('R')
                    was_rest_day = True
                else:
                    if was_rest_day and current_shift != 'G':
                        current_shift = self.get_next_shift(current_shift)
                        was_rest_day = False
                    schedule[emp['name']]['shifts'].append(current_shift)
                    
        return schedule

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()
    year = int(data.get('year', datetime.now().year))
    month = int(data.get('month', datetime.now().month))
    employees_data = data.get('employees', [])
    
    scheduler = DutyScheduler()
    schedule = scheduler.generate_schedule(employees_data, year, month)
    
    return jsonify({
        'schedule': schedule,
        'month': month,
        'year': year,
        'month_name': calendar.month_name[month]
    })

@app.route('/export', methods=['POST'])
def export():
    data = request.get_json()
    schedule = data.get('schedule', {})
    month = data.get('month')
    year = data.get('year')
    
    wb = Workbook()
    ws = wb.active
    
    # Styles
    header_fill = PatternFill(start_color="4A4A4A", end_color="4A4A4A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style='thin', color="000000"),
        right=Side(style='thin', color="000000"),
        top=Side(style='thin', color="000000"),
        bottom=Side(style='thin', color="000000")
    )
    
    # Calculate last column letter
    num_days = calendar.monthrange(year, month)[1]
    last_col = get_column_letter(num_days + 3)
    
    # Main Header - extend to last column
    ws.merge_cells(f'A1:{last_col}1')
    ws['A1'] = 'BAGASSE YARD SHIFT SCHEDULE'
    ws.merge_cells(f'A2:{last_col}2')
    ws['A2'] = f"{calendar.month_name[month].upper()} {year}"
    
    # Column Headers
    ws['A3'] = 'S.R'
    ws['B3'] = 'SUPERVISOR'
    ws['C3'] = 'CODE NO.'
    
    # Days and dates
    for day in range(1, num_days + 1):
        col = get_column_letter(day + 3)
        ws[f'{col}3'] = day
        ws[f'{col}4'] = DutyScheduler().get_day_name(year, month, day)
    
    # Write schedule
    row = 5
    sr_no = 1
    
    # First write supervisors
    for emp_name, data in schedule.items():
        if data['post'].upper() == 'SUPERVISOR':
            ws[f'A{row}'] = sr_no
            ws[f'B{row}'] = emp_name
            ws[f'C{row}'] = data['code']
            
            for col, shift in enumerate(data['shifts'], start=4):
                ws.cell(row=row, column=col, value=shift)
            
            row += 1
            sr_no += 1
    
    # Then write helpers
    for emp_name, data in schedule.items():
        if data['post'].upper() == 'HELPER':
            ws[f'A{row}'] = sr_no
            ws[f'B{row}'] = emp_name
            ws[f'C{row}'] = data['code']
            
            for col, shift in enumerate(data['shifts'], start=4):
                ws.cell(row=row, column=col, value=shift)
            
            row += 1
            sr_no += 1
    
    # Apply styles
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if row <= 4:  # Header rows
                cell.fill = header_fill
                cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 4  # S.R
    ws.column_dimensions['B'].width = 16  # Name
    ws.column_dimensions['C'].width = 8  # Code
    date_column_width = 4.5
    for col in range(4, num_days + 4):
        ws.column_dimensions[get_column_letter(col)].width = date_column_width
    
    # Set print settings
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToHeight = 1
    ws.page_setup.fitToWidth = 1
    
    ws.page_margins.left = 0.1
    ws.page_margins.right = 0.1
    ws.page_margins.top = 0.2
    ws.page_margins.bottom = 0.2
    
    # Save to BytesIO
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    
    return send_file(
        excel_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'duty_schedule_{year}_{month}.xlsx'
    )

if __name__ == '__main__':
    app.run(debug=True)
