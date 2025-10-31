import pymysql      # Used to connect and interact with MySQL
from openpyxl import Workbook  # Used to create and write Excel files
class connect:
    def __init__(self,host,user,password,database):  # Initialize database connection parameters
        self.host=host
        self.user=user
        self.password=password
        self.database=database
        self.con=None   # Connection object placeholder
    def mysql(self):
        self.con=pymysql.connect(
            host=self.host,
            user=self.user,
            password=self.password,
            database=self.database
        )
        return self.con  # Establish connection to MySQL database
db=connect("localhost","root","Kadarisrinuvas@321","netfilx")  # Create an object of Connect class and 
                                                                #connect to MySQL

con=db.mysql() # Open connection
obj=con.cursor()  # Create a cursor to execute SQL queries
obj.execute("select * from collection")  # Execute SQL query to fetch all data from 'collection' table and also netflix table
data =obj.fetchall()  # Fetch all the rows from the executed query
coloums=[desc[0] for desc in obj.description]    # Extract column names from cursor description
wd=Workbook ()  # Create a new Excel workbook
ws=wd.active   # Get the active worksheet
ws.title="netfilx_collection"  # Rename the sheet
ws.append(coloums)  # Write column headers to the first row in Excel
for rows in data :  # Loop through each row of data and write it to the Excel file
    ws.append(rows)
output_file="netfilc_collection.xlsx" # Define output Excel file name
wd.save(output_file) # Save the workbook to the specified file
con.close() # Close the database connection