using System.Data;
using System.Data.OleDb;


namespace RpaHelper
{
    /// <summary>
    /// Excel toolset to work with workbooks, worksheets, etc. Uses OLEDB.
    /// </summary>
    class Excel
    {
        /// <summary>
        /// Read Excel Worksheet by Sheet Name.
        /// </summary>
        public static DataTable ReadWorksheet(string workbookPath, string sheetName, bool hasHeaders)
        {
            //Connect to Workbook
            OleDbConnection oleExcelConnection = ConnectToWorksheet(workbookPath, hasHeaders);
            
            //Find Worksheet
            string? foundSheetName = FindWorksheetByName(oleExcelConnection, sheetName);
            if (foundSheetName is null)
            {
                oleExcelConnection.Close();
                throw new Exception($@"Workbook has no sheet ""{sheetName}""");
            }
            
            //Read Worksheet to DataTable
            return GetDataTable(oleExcelConnection, foundSheetName);
        }

        /// <summary>
        /// Read Excel Worksheet by Sheet Index.
        /// </summary>
        public static DataTable ReadWorksheet(string workbookPath, int sheetIndex, bool hasHeaders)
        {
            //Connect to Workbook
            OleDbConnection oleExcelConnection = ConnectToWorksheet(workbookPath, hasHeaders);
            
            //Find Worksheet
            string? foundSheetName = FindWorksheetByIndex(oleExcelConnection, sheetIndex);
            if (foundSheetName is null)
            {
                oleExcelConnection.Close();
                throw new Exception($@"Workbook has no sheet with index: {sheetIndex.ToString()}");
            }
            
            //Read Worksheet to DataTable
            return GetDataTable(oleExcelConnection, foundSheetName);
        }

        /// <summary>
        /// Returns an OleDbConnection to workbook.
        /// </summary>
        private static OleDbConnection ConnectToWorksheet(string workbookPath, bool hasHeaders)
        {
            //Basic Checks
            if (String.IsNullOrEmpty(workbookPath))
                throw new Exception($@"Workbook path cannot be null or empty!");
            if (!File.Exists(workbookPath))
                throw new Exception($@"File ""{workbookPath}"" not found!");
            
            //Prepare Connection parameters
            string headersProperty = hasHeaders ? "Yes" : "No";
            string Provider = "Microsoft.ACE.OLEDB.12.0";

            string ExcelVersion;
            switch (Path.GetExtension(workbookPath).ToLower())
            {
                case ".xlsx":
                    ExcelVersion = "Excel 12.0";
                    break;
                case ".xls":
                    ExcelVersion = "Excel 8.0";
                    break;
                case ".xlsm":
                    ExcelVersion = "Excel 12.0 Macro";
                    break;
                default:
                    throw new Exception($@"File ""{workbookPath}"" is not an Excel file!");
            }
            
            //Construct Connection String
            string ConnectionString = $@"
                Provider={Provider};
                Data Source={workbookPath};
                Extended Properties=""{ExcelVersion};HDR={headersProperty};IMEX=1""";

            //Open Connection
            OleDbConnection oleExcelConnection = new OleDbConnection(ConnectionString);
            oleExcelConnection.Open();

            return oleExcelConnection;
            
        }

        /// <summary>
        /// Search worksheet by Sheet Name.
        /// </summary>
        private static string? FindWorksheetByName(OleDbConnection oleExcelConnection, string sheetName)
        {
            
            if (String.IsNullOrEmpty(sheetName))
            {
                oleExcelConnection.Close();
                throw new Exception($@"Worksheet name cannot be null or empty!");
            }
            string? foundSheetName;
            //Read List of Worksheets
            DataTable dtTablesList = oleExcelConnection.GetSchema("Tables");

            //Worksheet Name always ends with '$' sign when using OLEDB
            string oleSheetName = sheetName + "$";

            //Check if WorkSheet is present in Workbook
            if (dtTablesList != null && dtTablesList.Rows.Count > 0)
            {
                foundSheetName = (
                    from r in dtTablesList.AsEnumerable()
                    where r.Field<string>("TABLE_NAME") == oleSheetName
                    select r.Field<string>("TABLE_NAME")
                    ).FirstOrDefault();
            }
            else
            {
                foundSheetName = null;
            }

            return foundSheetName;
        }

        /// <summary>
        /// Search worksheet by Sheet Index.
        /// </summary>
        private static string? FindWorksheetByIndex(OleDbConnection oleExcelConnection, int sheetIndex)
        {
            //Basic check
            if (sheetIndex < 0)
            {
                oleExcelConnection.Close();
                throw new Exception($@"Worksheet index should be greater than 0!");
            }
            string? foundSheetName;
            
            //Read List of Worksheets
            DataTable dtTablesList = oleExcelConnection.GetSchema("Tables");

            if (dtTablesList != null && (dtTablesList.Rows.Count - 1) >= sheetIndex)
            {
                foundSheetName = dtTablesList.Rows[sheetIndex].Field<string>("TABLE_NAME");
            }
            else
            {
                foundSheetName = null;
            }
            
            return foundSheetName;
        }

        /// <summary>
        /// Retrieve worksheet data as DataTable.
        /// </summary>
        private static DataTable GetDataTable(OleDbConnection oleExcelConnection, string foundSheetName)
        {
            using (OleDbCommand oleExcelCommand = new OleDbCommand())
            {
                oleExcelCommand.CommandText = $@"Select * From [{foundSheetName}]";
                oleExcelCommand.Connection = oleExcelConnection;
                using (OleDbDataAdapter oleExcelDataAdapter = new OleDbDataAdapter())
                {
                    DataTable result = new DataTable();
                    oleExcelDataAdapter.SelectCommand = oleExcelCommand;
                    oleExcelDataAdapter.Fill(result);
                    oleExcelConnection.Close();
                    return result;
                }
            }
        }
        //Other classes should go here.....
    }
}
