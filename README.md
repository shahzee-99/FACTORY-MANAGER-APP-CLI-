# FACTORY-MANAGER-APP-CLI-
Factory Manager CLI app manages factory workers by allowing registration, tracking daily work, and calculating monthly salaries. It exports data to Excel and generates detailed salary bills, deducting loans if any. Built with Python, Pandas, and openpyxl for data management and file handling.
This Python-based Factory Manager CLI Application helps manage worker registration, track daily work, and calculate monthly salaries for factory workers. The app allows you to enter data related to a worker’s daily output, track loans taken by workers, and generate salary bills at the end of each month.

Features

Worker Registration:
Register new factory workers by storing their basic information (name, age, and phone number) in text files.

Daily Work Entry:
Enter details of the work completed by a worker for the day, including the number of designs made and any loan taken.
The work data is stored in an Excel file that keeps track of each worker's daily output throughout the month.

Monthly Salary Calculation:
Calculate the salary for a worker based on the number of designs completed in a given month.
Automatically deduct any loans taken by the worker from the final salary.
Generate a detailed salary bill, including design-wise breakdowns, total loans, and final salary after deductions.

Tech Stack

Python 3.x for the core functionality
Pandas for data manipulation and Excel file handling
Openpyxl for working with Excel files
Datetime & Calendar for managing date and time operations

How It Works

Registration:
Workers are registered by entering their name, age, and phone number.
The worker's information is saved in a text file named after the worker.

Daily Work Entry:
For each day, the worker’s name, design worked on, number of designs completed, and any loan amount is entered.
This information is stored in an Excel file named after the worker (<worker_name>_work.xlsx), which tracks their daily output.

Salary Calculation:
At the end of the month, the worker’s salary is calculated based on the total number of designs made and the respective prices.
The loan amount (if any) is deducted from the total salary.
A salary bill is generated in a text file with a breakdown of each day’s work and the final salary after deductions.

Usage
Register a Worker:
Run the app and select option 1 to register a new worker by providing the required information.

Enter Daily Work:
Select option 2 to enter a worker's daily work, including the design count and loan amount (if applicable).

Calculate Salary:
To calculate a worker's salary for a specific month, choose option 3. The app will generate a detailed salary bill and save it to a text file.

Quit the Program:
Type exit to quit the program at any time.
