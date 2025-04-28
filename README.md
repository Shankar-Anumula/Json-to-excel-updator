yo# Json-to-excel-updator

This Python utility reads multiple JSON files containing employee utilisation data, identifies the correct rows in corresponding Excel files based on `empid` and `date`, and updates only the relevant cells in-place. It also creates backups of Excel files and archives processed JSON files.

## Features

- Updates only specific cells in Excel (`.xlsx`) files based on `empid` and `date`.
- Reads and processes multiple JSON files from a folder.
- Supports multiple Excel files (one per team).
- Backs up Excel files before updating.
- Archives JSON files after processing.
- Preserves all existing Excel formatting and formulas.


## JSON Input Format

Each JSON file should be an array of objects with the following keys:

```json
[
  {
    "empid": "E123",
    "team": "TeamA",
    "lob": "LOB1",
    "date": "2024-04-25",
    "analysis": 2,
    "design": 1,
    "test execution": 3,
    "regression": 0,
    "demo": 1,
    "leave": 0,
    "downtime": 0
  },
  ...
]
```


# Scheduling

This repository contains a Python script that reads data from a JSON file and updates an Excel file using `pandas` and `openpyxl`. The script can be scheduled to run automatically every 6 hours on a Windows machine.

---

## Requirements

- Python 3.x
- `pandas` library
- `openpyxl` library

Install the required libraries using:
```bash
pip install pandas openpyxl
```

---

## Steps to Schedule the Script to Run Every 6 Hours (Windows)

### 1. Open Task Scheduler
1. Press `Windows + S` to open the search bar.
2. Type **Task Scheduler** and open it.

### 2. Create a Basic Task
1. In the **Task Scheduler**, click on **"Create Basic Task..."** from the right-hand panel.
2. Enter a **Name** (e.g., "Run Python Script Every 6 Hours") and optionally a description.
3. Click **Next**.

### 3. Set the Trigger
1. Select **Daily** as the trigger and click **Next**.
2. Set the **Start date and time** (e.g., 12:00 AM if you want the first execution to happen at midnight).
3. On the next screen, check the **"Repeat task every"** option and set it to **6 hours**. Set the **duration** to **Indefinitely**.
4. Click **Next**.

### 4. Set the Action
1. Select **"Start a Program"** and click **Next**.
2. In the **Program/script** field, enter the path to your Python interpreter (e.g., `C:\Python39\python.exe`).
3. In the **Add arguments (optional)** field, enter the path to your script (e.g., `C:\path\to\your_script.py`).
4. Click **Next**.

### 5. Confirm and Finish
1. Review the details and click **Finish** to create the task.

### 6. Test the Task
1. Locate your task in the **Task Scheduler Library**.
2. Right-click it and select **Run** to confirm it works.

---

## Example Script

Below is an example Python script that reads data from a JSON file and updates an Excel file:

```python
import pandas as pd

# Load data from JSON file
json_data = 'data.json'
data = pd.read_json(json_data)

# Update Excel file
excel_file = 'output.xlsx'
with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
    data.to_excel(writer, index=False, sheet_name='UpdatedData')

print(f"Data from {json_data} has been updated in {excel_file}.")
```

---

## Notes
- Ensure the paths to the Python interpreter and the script are correct when setting up the task in the Task Scheduler.
- You can modify the script to suit your specific requirements.

---

Feel free to contribute or raise an issue if you encounter any problems!

