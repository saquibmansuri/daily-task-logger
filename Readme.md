# Daily Task Logger

## Overview

Daily Task Logger is a lightweight desktop application built with Python and Tkinter for tracking daily tasks for multiple projects. Users can add tasks with project details, task descriptions, hours spent, and optional comments. Tasks can be saved to both a TXT log and an Excel file for record keeping at one place. Users can share this sheet with their TL, Manager or Client at the end of the month/billing cycle.

## Prerequisites

* Python 3.10 or higher
* Required Python packages:

  * `openpyxl`
  * `tkinter` (usually included with Python)

Install required packages using:

```bash
pip install openpyxl
```

## Setup

1. Download the project files to your local machine.
2. Ensure `main.py` and the project folder are in a single directory.
3. Run the application:

```bash
python main.py
```

This will open the GUI where you can start adding tasks.

## Saving

* Tasks are saved in `tasks.txt` and `tasks.xlsx` in the same folder as the script.
* Make sure `tasks.xlsx` is **closed** before saving to avoid permission errors.

## Scheduling the Script

To automatically run the task logger at a specific time, you can use **Windows Task Scheduler**:

1. Open Task Scheduler.
2. Click **Create Task**.
3. Under **General**, name your task.
4. Under **Triggers**, click **New** and set the schedule (daily, weekly, etc.).
5. Under **Actions**, click **New**:

   * Action: Start a program
   * Program/script: `python`
   * Add arguments: `C:\path\to\main.py`  (replace with your script path)
   * Start in: `C:\path\to\folder` (folder where script is located)
6. Click **OK** to save the task.

Now, the task logger will run automatically according to the schedule and prompt you for the details, no need to remember.

## Notes

* It will create both Txt and Excel file if not present.
* Please keep the tasks.xlsx (Excel file) closed while saving the tasks to avoid errors.
* TXT logs can remain open in editors like Notepad without issues.
* If you are using a Python virtual environment, the Windows Task Scheduler may not work directly because it might point to a different Python interpreter than the one in your virtual environment. Simplest workaround is to create a batch file that activates the virtual environment and runs the script
