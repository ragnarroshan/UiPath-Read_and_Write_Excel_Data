## Exercise_4-UiPath-Read_and_Write_Excel_Data
## Aim:
To automate the process of reading data from an Excel file and writing data into another Excel file or a different sheet using UiPath.

## Equipment Required:
UiPath Studio (installed on a compatible system).<br>
Microsoft Excel application.<br>
Excel file (e.g., "ExcelFile.xlsx") containing data for reading and writing purposes.<br>
Computer with:<br>
Minimum 4 GB RAM.<br>
Minimum 2.0 GHz CPU.<br>
Windows operating system.<br>
.NET Framework 4.6.1 or later.
## Procedure:
### Start a New Process:

Open UiPath Studio.<br>
Create a new process by selecting Process under the New Project tab.<br>
Name the project UiPath Read and Write Excel Data and click Create to begin.
### Install Excel Activities Package:

If the Excel activities are not already available, go to Manage Packages (click on the Project panel).<br>
Search for UiPath.Excel.Activities in the Official packages.<br>
Click Install and then Save to include the Excel activities in the project.<br>
### Add Excel Application Scope:

In the Activities panel, search for Excel Application Scope.
Drag and drop the Excel Application Scope activity into the main sequence window.<br>
In the Properties pane, specify the path to the Excel file (e.g., "C:\Path\To\Your\ExcelFile.xlsx").<br>
The Excel Application Scope will open the Excel file and keep it open while the robot interacts with it.<br>
### Read Data from Excel:

Inside the Excel Application Scope, search for Read Range in the Activities panel.<br>
Drag the Read Range activity inside the Excel Application Scope.
Set the SheetName property to "Sheet1" (or the sheet from which data needs to be read).<br>
Leave the Range field blank if you want to read all data from the sheet, or specify a range like "A1:C10" if you want specific cells.<br>
Create a variable named dtExcelData to store the output data, which will be of type DataTable.<br>
### Modify the Data (Optional):

If you wish to process the data after reading it, you can use activities like For Each Row to loop through the rows of the DataTable.<br>
Example: Use Assign activities to modify certain cell values, calculate totals, or apply filters.
### Write Data into Excel:

After processing or modifying the data, add a Write Range activity under the Excel Application Scope.<br>
Set the SheetName to a different sheet, such as "Sheet2" (or the same sheet, if preferred).<br>
Set the DataTable property to dtExcelData to write the data back to the file.<br>
Specify the Range (e.g., "A1") to indicate where the data should be written.<br>
Ensure the Add Headers option is checked if the data contains column headers that need to be written.
### Save and Run the Workflow:

Save the project by pressing CTRL+S.<br>
Click Run from the toolbar to execute the workflow.<br>
UiPath will read data from Sheet1, process it (if specified), and then write the data to Sheet2 (or the same sheet).
### Example Scenario:
Initial Excel File:

Sheet1 contains customer data: Name, Age, City.<br>
Sheet2 is initially empty.
### UiPath Workflow Actions:

The robot reads the data from Sheet1.<br>
(Optional) The data is modified (e.g., age is incremented by 1).<br>
The modified data is written to Sheet2, including the headers.
### Expected Outcome:

Sheet2 will now have the same data as Sheet1 (or the modified version), demonstrating how UiPath can read and write data efficiently.
## UiPath WorkFlow:
![alt text](<img/Screenshot 2024-09-16 220140.png>)
![alt text](<img/Screenshot 2024-09-16 220212.png>)
![alt text](<img/Screenshot 2024-09-16 220227.png>)
![alt text](<img/Screenshot 2024-09-16 220300.png>)
![alt text](<img/Screenshot 2024-09-16 220315.png>)
![alt text](<img/Screenshot 2024-09-16 220401.png>)
![alt text](<img/Screenshot 2024-09-16 220422.png>)
![alt text](<img/Screenshot 2024-09-16 220529.png>)
![alt text](<img/Screenshot 2024-09-16 220546.png>)
## Result:
The automation workflow successfully reads data from Sheet1 of the Excel file, processes it if required, and writes the data into Sheet2 (or another location) of the same Excel file. This confirms the ability of UiPath to interact with Excel for reading and writing operations.