import pandas as pd
from datetime import datetime
import calendar
import openpyxl
import pyttsx3

engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id)
rate = engine.setProperty("rate", 200)

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
        # Add more designs and their prices as needed
}

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def registeration():
    # This function is responsible for the registration of the user 
    name = input("Enter worker's name: ") 
    age = int(input("Enter worker's age: "))
    phone = input("Enter worker's contact or phone number: ")

    try:
        with open(f"{name.capitalize()}.txt", "w") as f:  # It creates a file for the user, storing their provided data
            f.write(f"Name: {name.capitalize()}\nAge: {str(age)}\nPhone Number: {phone}\n")
        print("Congratulations! Your registration has been successful.")
        speak("Congratulations! Your registration has been successful.")

    except Exception as e:
        print(f"An error occurred during registration: {e}")
        speak(f"An error occurred during registration: {e}")

def today_work():
    # This function is responsible for getting worker's work for a specific date
    name = input("Enter worker's name: ").lower().strip()
    date_input = input("Enter the date for the work entry (YYYY-MM-DD) or press Enter for today: ").strip()
    # Use today's date if no date is entered
    work_date = date_input if date_input else datetime.now().strftime("%Y-%m-%d")
    
    try:
        # Validate the date format
        datetime.strptime(work_date, "%Y-%m-%d")
    except ValueError:
        print("Invalid date format! Please use YYYY-MM-DD.")
        speak("Invalid date format! Please use YYYY-MM-DD.")
        return

    design_name = input("Enter the name of the design: ")
    design_count = int(input("Enter the number of designs that the worker made: "))
    choice = input("Has the worker taken a loan (yes/no): ").lower().strip()
    loan_amount = int(input("Enter the amount of loan: ")) if choice == "yes" else 0

    # Prepare data to append to file and export to Excel
    work_data = {
        'Date': work_date,
        'Worker Name': name,
        'Design Name': design_name,
        'Design Count': design_count,
        'Loan Amount': loan_amount
    }

    excel_file = f"{name}_work.xlsx"

    try:
        # Check if file exists; if not, create a new one with the first entry
        try:
            # Read existing data
            existing_data = pd.read_excel(excel_file)
            # Append new data
            df = pd.DataFrame([work_data])
            updated_data = pd.concat([existing_data, df], ignore_index=True)

        except FileNotFoundError:
            # If file doesn't exist, create new data
            updated_data = pd.DataFrame([work_data])

        # Write the updated data to the Excel file, overwriting the previous content
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
            updated_data.to_excel(writer, index=False)

        print(f"Work for {work_date} has been successfully added to {excel_file}.")
        speak(f"Work for {work_date} has been successfully added to {excel_file}.")

    except Exception as e:
        print(f"An error occurred: {e}")
        speak(f"An error occurred: {e}")

def calculateSalary(name, month, year): # This function is used to calculate the worker's salary
    # Define a dictionary with design names and their prices


    # Define the date range for the month
    start_date = f"{year}-{month:02d}-01"
    end_date = f"{year}-{month:02d}-{calendar.monthrange(year, month)[1]}"

    try:
        # Load worker's work data from the Excel file
        file_name = f"{name}_work.xlsx"
        df = pd.read_excel(file_name)

        # Filter data for the selected month
        df['Date'] = pd.to_datetime(df['Date'])
        df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

        # Initialize variables
        total_salary = 0
        total_loans = 0
        bill_details = []  # To store details for the bill

        # Calculate salary based on design prices and work data
        for _, row in df.iterrows():
            design_name = row['Design Name']
            design_count = row['Design Count']
            loan_amount = row['Loan Amount']

            # Get the price for the design
            if design_name in design_prices:
                design_price = design_prices[design_name]

            else:
                print(f"Design '{design_name}' not found in price list. Skipping.")
                speak(f"Design '{design_name}' not found in price list. Skipping.")
                continue

            # Calculate salary for this entry
            design_salary = design_count * design_price
            total_salary += design_salary
            total_loans += loan_amount

            # Append detail to bill
            bill_details.append({
                'Date': row['Date'].strftime('%Y-%m-%d'),
                'Design Name': design_name,
                'Design Count': design_count,
                'Design Price': design_price,
                'Total for Design': design_salary
            })

        # Subtract total loan amount
        final_salary = total_salary - total_loans

        # Generate and save the bill
        bill_file = f"{name}_salary_bill_{year}-{month:02d}.txt"

        with open(bill_file, 'w') as f:
            f.write(f"Salary Bill for {name.capitalize()}\n")
            f.write(f"Month: {calendar.month_name[month]} {year}\n")
            f.write(f"Date Range: {start_date} to {end_date}\n\n")
            f.write("Design Details:\n")

            for item in bill_details:
                f.write(f"{item['Date']} - {item['Design Name']}: {item['Design Count']}  PKR {item['Design Price']} each = PKR {item['Total for Design']}\n")

            f.write("\n")
            f.write(f"Total Salary (before loan deduction): PKR {total_salary}\n")
            f.write(f"Total Loans: PKR {total_loans}\n")
            f.write(f"Final Salary (after loan deduction): PKR {final_salary}\n")

        print(f"Worker: {name.capitalize()}")
        speak(f"Worker: {name}")

        print(f"Total Salary (before loan deduction): PKR {total_salary}")
        speak(f"Total Salary (before loan deduction): PKR {total_salary}")

        print(f"Total Loans: PKR {total_loans}")
        speak(f"Total Loans: PKR {total_loans}")

        print(f"Final Salary (after loan deduction): PKR {final_salary}")
        speak(f"Final Salary (after loan deduction): PKR {final_salary}")

        print(f"Bill has been generated: {bill_file}")
        speak("Bill has been generated")

    except Exception as e:
        print(f"An error occurred during salary calculation: {e}")
        speak(f"An error occurred during salary calculation: {e}")


if __name__ == "__main__":
    
    try:
        print("\n-----------------WELCOME TO FACTORY MANAGER APP-------------------")
        speak("WELCOME TO FACTORY MANAGER APP, CREATED BY MUHAMMAD SHAHZAIB. IT'S FREE VERSION")
        while True:
            print("===========================================================") 
            print("\n-1 Register a new worker\n-2 Enter today's work\n-3 Calculate Salary after month\n-4 Add new Design\nExit for quit\n")

            choice = input("Enter what you want to do: ").lower().strip()

            if choice == "1":
                registeration()

            elif choice == "2":
                today_work()

            elif choice == "3":
                name = input("Enter the name of worker: ")
                month = int(input("Enter the month for salary calculation (numeric): "))
                year = int(input("Enter year: "))
                calculateSalary(name, month, year)
            
            elif choice == "4":
                desing_name = input("Enter the design name: ")
                desing_price = int(input("Enter the design price: "))
                design_prices.update({desing_name : desing_price})
                speak(f"design named {desing_name} has been added with price {desing_price}.")

            elif choice == "exit":
                speak("Factory Manager App is quitting....")
                break

            else:
                print("Enter a valid input [1, 2, 3, exit]")
                speak("Enter a valid input [1, 2, 3, exit]")

    except Exception as e:
        print(f"An error occurred: {e}")
        speak(f"An error occurred: {e}")
