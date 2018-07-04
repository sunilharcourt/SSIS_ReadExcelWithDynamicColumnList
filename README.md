# Read data from excel with dynamic column list using SSIS

**Problem Statement**  
I have multiple excel files lying under a single folder. I want to import into separate tables. I don't know the name of the files, the name of the first Sheet, nor the number of columns. So here is at least one way to do it using SSIS.

I have created a script task that reads metadata of excel dynamically and creates table in SQL Server and then copies data into it.

**Step 1: Add ForEach Loop Container**
1. Before adding foreach loop container, create three user variables of type String - ErrorDesc, ExcelSheetName, ImportFilePath 
1. Add for each loop container with the following settings - 
1. Enumerator - ForEach file enumerator 
1. In folder settings, paste folder that you has all excel files
1. In files settings, put *.xlsx
1. In retrieve file name, select "Fully qualified"
1. Go to variables settings of foreach loop container, and set USer::ImportFilePath (with Index 0)

**Step 2: Add a Script task inside ForEach Loop Container**
1. Double click on Script task and set ReadOnlyVariables for User::ImportFilePath, and set ReadWriteVariables for User::ErrorDesc,User::ExcelSheetName
2. Now, Click on 'Edit Script'
3. Replace the content of Main() method with script in project "Script.cs"

**Note** - Make sure you have installed Office driver for Excel connectivity 

**Suggestion** - 
1. To start with, first download and install 32 bit driver from here - https://www.microsoft.com/en-us/download/details.aspx?id=23734
2. Configure SSIS project to run in 32 bit mode
3. Run the project and check if data is imported successfully
4. If done, then install 64 bit driver and run SSIS project for 64 bit and check

Let me know if you are facing any issue in setting up things. 

