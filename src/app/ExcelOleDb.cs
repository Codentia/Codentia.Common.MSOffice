using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace Codentia.Common.MSOffice
{
    /// <summary>
    /// This static class exposes methods used to interrogate Microsoft Excel WorkBooks (XLS Files) and WorkSheets using
    /// the Jet (OleDb) interface.
    /// </summary>
    public static class ExcelOleDb
    {
        /// <summary>
        /// Fetch an array of the WorkSheets found within the specified Excel WorkBook.
        /// </summary>
        /// <param name="format">OfficeFileFormat to load as</param>
        /// <param name="xlsPath">Full or relative path to the file</param>
        /// <returns>string array</returns>
        public static string[] GetWorkSheetNames(OfficeFileFormat format, string xlsPath)
        {
            OleDbConnection conn = GetOleDbConnection(format, xlsPath);
            DataTable schemaData = conn.GetSchema("Tables");
            conn.Close();
            conn.Dispose();

            List<string> worksheetList = new List<string>();
            for (int i = 0; i < schemaData.Rows.Count; i++)
            {
                string worksheet = Convert.ToString(schemaData.Rows[i]["TABLE_NAME"]).Replace("$", string.Empty).Trim();

                if (!worksheet.EndsWith("_") && !worksheet.StartsWith("_") && !worksheetList.Contains(worksheet) && !worksheet.EndsWith("FilterDatabase"))
                {
                    worksheetList.Add(worksheet);
                }
            }

            string[] worksheets = new string[worksheetList.Count];
            worksheetList.CopyTo(worksheets);

            return worksheets;
        }

        /// <summary>
        /// Retrieve the contents of the specified worksheet within the specified workbook as a DataTable.
        /// </summary>
        /// <param name="format">OfficeFileFormat to load as</param>
        /// <param name="xlsPath">Full or relative path to the file</param>
        /// <param name="worksheetName">Name of the worksheet to open and return</param>
        /// <returns>DataTable of worksheet</returns>
        public static DataTable GetWorkSheet(OfficeFileFormat format, string xlsPath, string worksheetName)
        {
            DataTable returnData = new DataTable();
            OleDbConnection conn = GetOleDbConnection(format, xlsPath);

            try
            {
                OleDbDataAdapter oleAdapter = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}$]", worksheetName), conn);
                oleAdapter.Fill(returnData);
                oleAdapter.Dispose();
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to open the specified worksheet", ex);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            returnData.TableName = worksheetName;

            return returnData;
        }

        /// <summary>
        /// Load all worksheets within the specified workbook and return as a DataSet containing named DataTables.
        /// </summary>
        /// <param name="format">OfficeFileFormat to load as</param>
        /// <param name="xlsPath">Full or relative path to the file</param>
        /// <returns>DataSet of workbook</returns>
        public static DataSet GetWorkBook(OfficeFileFormat format, string xlsPath)
        {
            DataSet xlsData = new DataSet();

            string[] worksheets = GetWorkSheetNames(format, xlsPath);

            for (int i = 0; i < worksheets.Length; i++)
            {
                DataTable sheet = GetWorkSheet(format, xlsPath, worksheets[i]);
                xlsData.Tables.Add(sheet);
            }

            return xlsData;
        }

        /*
        public static void Insert(OfficeFileFormat format, string xlsPath, string worksheet, DataTable worksheetData)
        {
            throw new System.NotImplementedException();
            *//*
string connectionString = @"Provider=Microsoft.Jet.
   OLEDB.4.0;Data Source=Book1.xls;Extended
   Properties=""Excel 8.0;HDR=YES;""";

DbProviderFactory factory =
   DbProviderFactories.GetFactory("System.Data.OleDb");

using (DbConnection connection = factory.CreateConnection())
{
    connection.ConnectionString = connectionString;

    using (DbCommand command = connection.CreateCommand())
    {
        command.CommandText = "INSERT INTO [Cities$]
         (ID, City, State) VALUES(4,\"Tampa\",\"Florida\")";

        connection.Open();

        command.ExecuteNonQuery();
    }
}        
             */ 
/*        }

        public static void Update(OfficeFileFormat format, string xlsPath, string worksheet, DataTable updatedWorksheetData, string indexColumn)
        {
            throw new System.NotImplementedException();
        }

        public static void Delete(OfficeFileFormat format, string xlsPath, string worksheet, string indexColumn, object indexValue)
        {
            throw new System.NotImplementedException();
        }
*/

        /// <summary>
        /// Validate the given path and open an OleDbConnection if valid
        /// </summary>
        /// <param name="format">OfficeFileFormat to load as</param>
        /// <param name="xlsPath">Full or relative path to the Excel workbook to be opened</param>
        /// <returns>OleDbConnection object</returns>
        private static OleDbConnection GetOleDbConnection(OfficeFileFormat format, string xlsPath)
        {
            OleDbConnection conn = null;
            string connString = string.Empty;

            switch (format)
            {
                case OfficeFileFormat.Excel97_2003:
                ////    connString = string.Format("Provider=Microsoft.Jet.Oledb.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"", xlsPath);
                    connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"", xlsPath);
                    break;
                case OfficeFileFormat.Excel2007:
                    connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES\";", xlsPath);
                    break;
            }

            if (!string.IsNullOrEmpty(xlsPath))
            {
                try
                {
                    conn = new OleDbConnection(connString);
                    conn.Open();
                }
                catch (Exception ex)
                {
                    throw new Exception("Unable to open the specified file - NOTE that 'enable 32bit' is required in IIS7/64bit!", ex);
                }
            }
            else
            {
                throw new Exception("Unable to open the specified file");
            }

            return conn;
        }
    }
}
