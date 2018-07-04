
 public void Main()
     {


            string connectionString; // connection string for SQL
            OleDbConnection excelConnection; // oledb connection for Excel
            string SQLConnectionStringName; // connection string for SQL server
            SqlConnection sqlcon; //Sql connection for connecting SQL Server
            DataTable tablesInFile, ColumnList;
            int tableCount = 0;
            string currentTable;
            string CreateColumnList = "";
            string SelectColumnList = "";
            string CreateTableSQL;

            try
            {
                string fileName = Path.GetFileNameWithoutExtension(Dts.Variables["ImportFilePath"].Value.ToString());
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                        "Data Source=" + Dts.Variables["ImportFilePath"].Value + ";Extended Properties=Excel 12.0";
                SQLConnectionStringName = "Data Source=.;Initial Catalog=tempdb;Integrated Security=True;MultipleActiveResultSets=True";
                excelConnection = new OleDbConnection(connectionString);
                excelConnection.Open();
                tablesInFile = excelConnection.GetSchema("Tables");
                tableCount = tablesInFile.Rows.Count;
                if (tableCount > 0)
                {
                    currentTable = tablesInFile.Rows[0]["TABLE_NAME"].ToString();

                    Dts.Variables["ExcelSheetName"].Value = currentTable;

                    ColumnList  = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns,
                                                                        new Object[] { null, null, currentTable, null }
                                                                        );
                    foreach (DataRow r in ColumnList.Rows)
                    {
                        CreateColumnList = CreateColumnList + ", " + r["COLUMN_NAME"].ToString() + " nvarchar(max) ";
                        SelectColumnList = SelectColumnList + ", " + r["COLUMN_NAME"].ToString();
                    }

                    CreateColumnList = CreateColumnList.Substring(1);
                    SelectColumnList = " select " + SelectColumnList.Substring(1) + " FROM ["+currentTable+"]";

                    //connect sql server and create table
                    CreateTableSQL = " CREATE TABLE ["+ fileName +"_"+ currentTable + "] (" + CreateColumnList + ")";
                    sqlcon = new SqlConnection(SQLConnectionStringName);
                    sqlcon.Open();
                    SqlCommand command = new SqlCommand(CreateTableSQL, sqlcon);
                    command.ExecuteNonQuery();

                    //let's copy data now
                    OleDbCommand oledbcommand = new OleDbCommand(SelectColumnList, excelConnection);

                    OleDbDataReader oldDBReader = oledbcommand.ExecuteReader();
                    SqlBulkCopy bulkCopy = new SqlBulkCopy(SQLConnectionStringName, SqlBulkCopyOptions.TableLock);
                    bulkCopy.DestinationTableName = fileName +"_"+ currentTable ;
                    bulkCopy.WriteToServer(oldDBReader);

                    //cleanup
                    oldDBReader.Close();
                    bulkCopy.Close();
                    sqlcon.Close();
                    excelConnection.Close();

                    Dts.TaskResult = (int)ScriptResults.Success;

                }
                else
                {
                    Dts.Variables["ErrorCode"].Value = "Excel Sheet Lookup Failure. -- No Sheets Found.";
                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
            }
            catch (Exception e)
            {
                Dts.Variables["ErrorCode"].Value = "Excel Sheet Lookup Failure. -- " + e.Message;
                throw new Exception("Excel Sheet Lookup Failure.", e);
            }
        }

      
