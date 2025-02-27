# Duty Schedule Generator

A web-based application for K.R Papers Pvt. Ltd that generates monthly shift schedules for employees following a rotational shift pattern.

## Developer
Nitin Shukla

## Features

- Dynamic Shift Scheduling with rotation pattern (A → C → B → A)
- Rest Day Logic - Shift changes occur only after a rest day
- Fixed General Duty Employees
- Supervisor & Helper Order maintained
- Export to Excel functionality
- Modern web-based UI
- Shift timings:
  - A Shift: 06:00-14:00
  - B Shift: 14:00-22:00
  - C Shift: 22:00-06:00
  - G: General Duty
  - R: Rest Day

## Setup Instructions

1. Install Python 3.8 or higher if not already installed

2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python app.py
   ```

4. Open a web browser and navigate to:
   ```
   http://localhost:5000
   ```

## Usage

1. Select the starting shift (A, B, or C)
2. Choose the month and year
3. Click "Generate Schedule" to view the schedule
4. Click "Export to Excel" to download the schedule as an Excel file

## Technical Details

- Backend: Python Flask
- Frontend: HTML, CSS, JavaScript
- Libraries: 
  - Flask for web server
  - Pandas for data manipulation
  - OpenPyXL for Excel file generation
  - Bootstrap for UI components
