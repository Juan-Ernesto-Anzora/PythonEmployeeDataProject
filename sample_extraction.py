import pandas as pd
import pyodbc
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

try: 
    # Read data from the Excel file
    data = pd.read_excel('./dataset/Employees.xlsx')
    logging.info("Successfully read data from Excel file.")

    # Display the data
    print("Data from Excel:")
    print(data)


    # Database connection details
    server = 'DESKTOP-RJND44F'
    database = 'EmployeeData'

    # Establish a database connection using Windows authentication
    connection_string = (
        f'DRIVER={{ODBC Driver 18 for SQL Server}};'
        f'SERVER={server};'
        f'DATABASE={database};'
        f'Trusted_Connection=yes;'
        f'TrustServerCertificate=yes;'
    )
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()
    logging.info("Connected to SQL Server successfully.")

    # Insert data into the database
    for index, row in data.iterrows():
        try:
            cursor.execute(
                "INSERT INTO Employees (No,FirstName,LastName,Gender,StartDate,Years,Department,Country,Center,MonthlySalary,AnnualSalary,JobRate,SickLeaves,UnpaidLeaves,OvertimeHours) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                row['No'], row['First Name'], row['Last Name'], row['Gender'], row['Start Date'], row['Years'], row['Department'], row['Country'], row['Center'], row['Monthly Salary'], row['Annual Salary'], row['Job Rate'], row['Sick Leaves'], row['Unpaid Leaves'], row['Overtime Hours']
            )
        except Exception as e:
            logging.error(f"Error inserting row {index + 1}: {e}")

    # Commit the transaction
    connection.commit()
    logging.info("Transaction committed successfully.")

except pyodbc.Error as e:
    logging.error(f"Database error: {e}")
except FileNotFoundError as e:
    logging.error(f"Excel file not found: {e}")
except Exception as e:
    logging.error(f"An unexpected error occurred: {e}")
finally:
    # Ensure that the database connection is closed
    try:    
        # Close the connection
        cursor.close()
        connection.close()
        logging.info("Database connection closed.")

    except Exception as e:
        logging.error(f"Error closing the connection: {e}")
print("Data successfully inserted into the database.")