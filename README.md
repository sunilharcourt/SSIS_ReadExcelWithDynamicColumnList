# Read data from excel with dynamic column list using SSIS
SSIS does not provide any solution to pull data from multiple excels when column list is dynamic.  
To solve that problem, I have created a script task that reads metadata of excel dynamically and creates table in SQL Server and then copies data into it.
Check Wiki page to get more understanding of settings etc.
