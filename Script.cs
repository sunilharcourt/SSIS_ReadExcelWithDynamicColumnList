#region Help:  Introduction to the script task
/* The Script Task allows you to perform virtually any operation that can be accomplished in
 * a .Net application within the context of an Integration Services control flow. 
 * 
 * Expand the other regions which have "Help" prefixes for examples of specific ways to use
 * Integration Services features within this script task. */
#endregion


#region Namespaces
using System;
using System.Data;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Data.SqlClient;
#endregion

namespace ST_dada0d0f88b8402f9cc9182f44b439b4
{
    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
	public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
	{
        #region Help:  Using Integration Services variables and parameters in a script
        /* To use a variable in this script, first ensure that the variable has been added to 
         * either the list contained in the ReadOnlyVariables property or the list contained in 
         * the ReadWriteVariables property of this script task, according to whether or not your
         * code needs to write to the variable.  To add the variable, save this script, close this instance of
         * Visual Studio, and update the ReadOnlyVariables and 
         * ReadWriteVariables properties in the Script Transformation Editor window.
         * To use a parameter in this script, follow the same steps. Parameters are always read-only.
         * 
         * Example of reading from a variable:
         *  DateTime startTime = (DateTime) Dts.Variables["System::StartTime"].Value;
         * 
         * Example of writing to a variable:
         *  Dts.Variables["User::myStringVariable"].Value = "new value";
         * 
         * Example of reading from a package parameter:
         *  int batchId = (int) Dts.Variables["$Package::batchId"].Value;
         *  
         * Example of reading from a project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].Value;
         * 
         * Example of reading from a sensitive project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].GetSensitiveValue();
         * */

        #endregion

        #region Help:  Firing Integration Services events from a script
        /* This script task can fire events for logging purposes.
         * 
         * Example of firing an error event:
         *  Dts.Events.FireError(18, "Process Values", "Bad value", "", 0);
         * 
         * Example of firing an information event:
         *  Dts.Events.FireInformation(3, "Process Values", "Processing has started", "", 0, ref fireAgain)
         * 
         * Example of firing a warning event:
         *  Dts.Events.FireWarning(14, "Process Values", "No values received for input", "", 0);
         * */
        #endregion

        #region Help:  Using Integration Services connection managers in a script
        /* Some types of connection managers can be used in this script task.  See the topic 
         * "Working with Connection Managers Programatically" for details.
         * 
         * Example of using an ADO.Net connection manager:
         *  object rawConnection = Dts.Connections["Sales DB"].AcquireConnection(Dts.Transaction);
         *  SqlConnection myADONETConnection = (SqlConnection)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Sales DB"].ReleaseConnection(rawConnection);
         *
         * Example of using a File connection manager
         *  object rawConnection = Dts.Connections["Prices.zip"].AcquireConnection(Dts.Transaction);
         *  string filePath = (string)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Prices.zip"].ReleaseConnection(rawConnection);
         * */
        #endregion


		/// <summary>
        /// This method is called when this script task executes in the control flow.
        /// Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
        /// To open Help, press F1.
        /// </summary>
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

        #region ScriptResults declaration
        /// <summary>
        /// This enum provides a convenient shorthand within the scope of this class for setting the
        /// result of the script.
        /// 
        /// This code was generated automatically.
        /// </summary>
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

	}
}