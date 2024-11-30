import pandas as pd
from datetime import datetime
import calendar
import openpyxl
import pyttsx3

# Text-to-speech engine initialization
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id)
engine.setProperty("rate", 200)

# Define design prices
design_prices = {

    "tx bridge" : 15,
    "payper 18 piece" : 11,
    "honda watch" : 8,
    "smile lamba" : 9,
    "honda plastic" : 11,
    "baaz" : 20,
    "lf" : 9,
    "mubarik key chain" : 20,
    "star wali" : 8,
    "heel pad small" : 15,
    "heel pad big" : 18,
    "honda lamba" : 9,
    "naya jota upper" : 16,
    "smile emoji" : 9,
    "love lamba" : 11,
    "smile taj" : 9,
    "honda extreme" : 8,
    "honda football" : 14,
    "hi max" : 8,
    "black pad horse" : 10,
    "honda lopi" : 10,
    "tiktok" : 8,
    "facebook" : 8,
    "whatsapp" : 8,
    "honda tyre double side" : 14,
    "artmos canter" : 10,
    "baloon lamba" : 9,
    "5 design lamba key chain" : 12,
    "chand wala baloo (double side)" : 16,
    "3 dill + spider man" : 8,
    "white billi" : 10,
    "old jota upper" : 10,
    "honda par wala" : 9,
    "nexara 4" : 10,
    "apple key chain" : 8,
    "i love you lamba" : 10,
    "university key chain" : 8,
    "payper 16 piece" : 10,
    "old jota soul" : 10,
    "naya jota soul" : 15,
    "smile double side sialkot" : 14,
    "tom" : 18,
    "jerry" : 18,
    "snooker double side" : 14,
    "football double side" : 14,
    "chuzza double side" : 14,
    "pooh" : 14,
    "spiderman and superman double" : 10,
    # Add other designs and prices here...
}

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def registeration():
    """Register a new worker."""
    name = input("Enter worker's name: ").strip().capitalize()
    age = int(input("Enter worker's age: "))
    phone = input("Enter worker's contact or phone number: ")

    try:
        with open(f"{name}.txt", "w") as f:
            f.write(f"Name: {name}\nAge: {age}\nPhone Number: {phone}\n")
        print("Congratulations! Registration successful.")
        speak("Congratulations! Registration successful.")
    except Exception as e:
        print(f"Error during registration: {e}")
        speak(f"Error during registration: {e}")

def today_work():
    """Record today's work for a worker."""
    name = input("Enter worker's name: ").strip().lower()
    date_input = input("Enter the date for the work entry (YYYY-MM-DD) or press Enter for today: ").strip()
    work_date = date_input if date_input else datetime.now().strftime("%Y-%m-%d")

    try:
        datetime.strptime(work_date, "%Y-%m-%d")
    except ValueError:
        print("Invalid date format! Please use YYYY-MM-DD.")
        speak("Invalid date format! Please use YYYY-MM-DD.")
        return

    design_name = input("Enter the name of the design: ").strip().lower()
    design_count = int(input("Enter the number of designs made: "))
    has_loan = input("Has the worker taken a loan (yes/no): ").strip().lower()
    loan_amount = int(input("Enter the loan amount: ")) if has_loan == "yes" else 0

    work_data = {
        'Date': work_date,
        'Worker Name': name,
        'Design Name': design_name,
        'Design Count': design_count,
        'Loan Amount': loan_amount
    }

    excel_file = f"{name}_work.xlsx"

    try:
        try:
            existing_data = pd.read_excel(excel_file)
            updated_data = pd.concat([existing_data, pd.DataFrame([work_data])], ignore_index=True)
        except FileNotFoundError:
            updated_data = pd.DataFrame([work_data])

        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
            updated_data.to_excel(writer, index=False)

        print(f"Work for {work_date} has been successfully recorded in {excel_file}.")
        speak(f"Work for {work_date} has been successfully recorded.")
    except Exception as e:
        print(f"Error while saving work data: {e}")
        speak(f"Error while saving work data: {e}")

def calculateSalary(name, month, year):
    """Calculate the salary for a worker for a specific month."""
    start_date = f"{year}-{month:02d}-01"
    end_date = f"{year}-{month:02d}-{calendar.monthrange(year, month)[1]}"

    try:
        file_name = f"{name}_work.xlsx"
        df = pd.read_excel(file_name)

        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df.dropna(subset=['Date'], inplace=True)
        df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

        total_salary = 0
        total_loans = 0
        bill_details = []

        for _, row in df.iterrows():
            design_name = row['Design Name'].strip().lower()
            design_count = row['Design Count']
            loan_amount = row.get('Loan Amount', 0)

            if design_name in design_prices:
                design_price = design_prices[design_name]
                design_salary = design_count * design_price
                total_salary += design_salary
                bill_details.append({
                    'Date': row['Date'].strftime('%Y-%m-%d'),
                    'Design Name': design_name,
                    'Design Count': design_count,
                    'Design Price': design_price,
                    'Total for Design': design_salary
                })
            else:
                print(f"Design '{row['Design Name']}' not found. Skipping.")
            total_loans += loan_amount

        final_salary = total_salary - total_loans

        bill_file = f"{name}_salary_bill_{year}-{month:02d}.txt"
        with open(bill_file, 'w') as f:
            f.write(f"Salary Bill for {name.capitalize()}\n")
            f.write(f"Month: {calendar.month_name[month]} {year}\n")
            f.write(f"Date Range: {start_date} to {end_date}\n\n")
            f.write("Design Details:\n")
            for item in bill_details:
                f.write(f"{item['Date']} - {item['Design Name']}: {item['Design Count']} PKR {item['Design Price']} each = PKR {item['Total for Design']}\n")
            f.write(f"\nTotal Salary (before loan deduction): PKR {total_salary}\n")
            f.write(f"Total Loans: PKR {total_loans}\n")
            f.write(f"Final Salary (after loan deduction): PKR {final_salary}\n")

        print(f"Total Salary (before loan deduction): PKR {total_salary}")
        print(f"Total Loans: PKR {total_loans}")
        print(f"Final Salary (after loan deduction): PKR {final_salary}")
        print(f"Bill has been generated: {bill_file}")
        speak("Salary calculation complete. Bill has been generated.")

    except FileNotFoundError:
        print(f"No work data found for {name}.")
        speak(f"No work data found for {name}.")
    except Exception as e:
        print(f"Error during salary calculation: {e}")
        speak(f"Error during salary calculation: {e}")

if __name__ == "__main__":
    print("\n-----------------WELCOME TO FACTORY MANAGER APP-------------------")
    speak("Welcome to Factory Manager App, created by Muhammad Shahzaib.")

    while True:
        
        print("=======================================================")
        print("\nMenu:\n1. Register a new worker\n2. Enter today's work\n3. Calculate Salary\n4. Add new Design\nType 'exit' to quit\n")
        choice = input("Enter your choice: ").strip().lower()

        if choice == "1":
            registeration()
        
        elif choice == "2":
            today_work()
        
        elif choice == "3":
            name = input("Enter the name of the worker: ").strip().lower()
            month = int(input("Enter the month (numeric): "))
            year = int(input("Enter the year: "))
            calculateSalary(name, month, year)
        
        elif choice == "4":
            design_name = input("Enter the design name: ").strip().lower()
            design_price = int(input("Enter the design price: "))
            design_prices[design_name] = design_price
            print(f"Design '{design_name}' added with price {design_price}.")
            speak(f"Design '{design_name}' added.")
        
        elif choice == "exit":
            speak("Exiting Factory Manager App.")
            print("Goodbye!")
            break
        
        else:
            print("Invalid choice! Please try again.")
            speak("Invalid choice! Please try again.")