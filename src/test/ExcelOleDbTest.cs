using System;
using System.Data;
using Codentia.Test.Helper;
using NUnit.Framework;

namespace Codentia.Common.MSOffice.Test
{
    /// <summary>
    /// This class acts as the unit testing fixture for the static class ExcelOleDb
    /// <see cref="ExcelOleDb"/>
    /// </summary>
    [TestFixture]
    public class ExcelOleDbTest
    {
        /// <summary>
        /// Perform any activities required prior to unit testing (e.g. prepare test data)
        /// </summary>
        [TestFixtureSetUp]
        public void TextFixtureSetUp()
        {
        }

        /// <summary>
        /// Scenario: Invalid path or a path to a non-existant file specified
        /// Expected: Exception (Unable to open the specified file)
        /// </summary>
        [Test]
        public void _001_GetWorkSheetNames_InvalidFilePath()
        {
            // null
            Assert.That(delegate { ExcelOleDb.GetWorkSheetNames(OfficeFileFormat.Excel97_2003, null); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file"));

            // empty 
            Assert.That(delegate { ExcelOleDb.GetWorkSheetNames(OfficeFileFormat.Excel97_2003, string.Empty); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file"));

            // non-existant
            Assert.That(delegate { ExcelOleDb.GetWorkSheetNames(OfficeFileFormat.Excel97_2003, @"Z:\ThisXLSDoesNotExist.xls"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file - NOTE that 'enable 32bit' is required in IIS7/64bit!"));
        
            // not am xls                
            Assert.That(delegate { ExcelOleDb.GetWorkSheetNames(OfficeFileFormat.Excel97_2003, @"TestData\TextFile1.txt"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file - NOTE that 'enable 32bit' is required in IIS7/64bit!"));
        }

        /// <summary>
        /// Scenario: Work Sheet names retrieved for a workbook containing a single sheet
        /// Expected: string array of length 1 returned containing the expected name
        /// </summary>
        [Test]
        public void _002_GetWorkSheetNames_XLS_SingleSheet()
        {
            string[] names = ExcelOleDb.GetWorkSheetNames(OfficeFileFormat.Excel97_2003, @"TestData\xls1");
            Assert.That(names.Length, Is.EqualTo(1), "Expected 1");
            Assert.That(names[0], Is.EqualTo("Sheet1"), "Incorrect value");
        }

        /// <summary>
        /// Scenario: Work sheet names retrieved for a workbook containing more than one sheets
        /// Expected: string array of appropriate length containing the correct names
        /// </summary>
        [Test]
        public void _003_GetWorkSheetNames_XLS_ManySheets()
        {
            string[] names = ExcelOleDb.GetWorkSheetNames(OfficeFileFormat.Excel97_2003, @"TestData\xls2");
            Assert.That(names.Length, Is.EqualTo(3), "Expected 3");
            Assert.That(names[0], Is.EqualTo("first"), "Incorrect value");
            Assert.That(names[1], Is.EqualTo("second"), "Incorrect value");
            Assert.That(names[2], Is.EqualTo("third"), "Incorrect value");
        }

        /// <summary>
        /// Scenario: Invalid path or a path to a non-existant file specified
        /// Expected: Exception (Unable to open the specified file)
        /// </summary>
        [Test]
        public void _004_GetWorkSheet_InvalidFilePath()
        {
            // null
            Assert.That(delegate { ExcelOleDb.GetWorkSheet(OfficeFileFormat.Excel97_2003, null, "Sheet 1"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file"));

            // empty 
            Assert.That(delegate { ExcelOleDb.GetWorkSheet(OfficeFileFormat.Excel97_2003, string.Empty, "Sheet 1"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file"));

            // non-existant
            Assert.That(delegate { ExcelOleDb.GetWorkSheet(OfficeFileFormat.Excel97_2003, @"Z:\ThisXLSDoesNotExist.xls", "Sheet 1"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file - NOTE that 'enable 32bit' is required in IIS7/64bit!"));

            // not am xls                
            Assert.That(delegate { ExcelOleDb.GetWorkSheet(OfficeFileFormat.Excel97_2003, @"TestData\TextFile1.txt", "Sheet 1"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file - NOTE that 'enable 32bit' is required in IIS7/64bit!"));
        }     

        /// <summary>
        /// Scenario: Attempt made to retrieve a worksheet from a valid workbook - sheet does not exist
        /// Expected: Exception (Unable to open the specified worksheet)
        /// </summary>
        [Test]
        public void _007_GetWorkSheet_XLS_NonExistant()
        {
            Assert.That(delegate { ExcelOleDb.GetWorkSheet(OfficeFileFormat.Excel97_2003, @"TestData\xls1.xls", "wibble"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified worksheet"));
        }

        /// <summary>
        /// Scenario: Existing worksheet within a valid workbook retrieved
        /// Expected: DataTable accurately representing the data within the XLS sheet
        /// </summary>
        [Test]
        public void _008_GetWorkSheet_XLS_Valid()
        {
            DataTable method = ExcelOleDb.GetWorkSheet(OfficeFileFormat.Excel97_2003, @"TestData\xls1.xls", "Sheet1");
            DataTable check = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls1.xls", "Sheet1", true);

            Assert.That(SqlHelper.CompareDataTables(check, method), Is.True, "Data does not match");
        }

        /// <summary>
        /// Scenario: Invalid path or a path to a non-existant file specified
        /// Expected: Exception (Unable to open the specified file)
        /// </summary>
        [Test]
        public void _009_GetWorkBook_InvalidFilePath()
        {
            // null
            Assert.That(delegate { ExcelOleDb.GetWorkBook(OfficeFileFormat.Excel97_2003, null); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file"));

            // empty 
            Assert.That(delegate { ExcelOleDb.GetWorkBook(OfficeFileFormat.Excel97_2003, string.Empty); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file"));

            // non-existant
            Assert.That(delegate { ExcelOleDb.GetWorkBook(OfficeFileFormat.Excel97_2003, @"Z:\ThisXLSDoesNotExist.xls"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file - NOTE that 'enable 32bit' is required in IIS7/64bit!"));
            
            // not am xls              
            Assert.That(delegate { ExcelOleDb.GetWorkBook(OfficeFileFormat.Excel97_2003, @"TestData\TextFile1.txt"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified file - NOTE that 'enable 32bit' is required in IIS7/64bit!"));
        }

        /// <summary>
        /// Scenario: Workbook containing a single sheet opened
        /// Expected: DataSet containing a single (appropriately named) DataTable which in turn contains the data from the worksheet
        /// </summary>
        [Test]
        public void _011_GetWorkBook_XLS_SingleSheet()
        {
            DataSet dsMethod = ExcelOleDb.GetWorkBook(OfficeFileFormat.Excel97_2003, @"TestData\xls1.xls");
            Assert.That(dsMethod.Tables.Count, Is.EqualTo(1), "Incorrect count");
            Assert.That(dsMethod.Tables[0].TableName, Is.EqualTo("Sheet1"), "Incorrect TableName");

            DataTable check = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls1.xls", "Sheet1", true);
            Assert.That(SqlHelper.CompareDataTables(check, dsMethod.Tables[0]), Is.True, "Data does not match");
        }

        /// <summary>
        /// Scenario: Workbook containing more than one sheet opened
        /// Expected: DataSet contains one DataTable for each sheet in the workbook, each with the appropriate data
        /// </summary>
        [Test]
        public void _012_GetWorkBook_XLS_ManySheets()
        {
            DataSet dsMethod = ExcelOleDb.GetWorkBook(OfficeFileFormat.Excel97_2003, @"TestData\xls2.xls");
            Assert.That(dsMethod.Tables.Count, Is.EqualTo(3), "Incorrect count");
            Assert.That(dsMethod.Tables[0].TableName, Is.EqualTo("first"), "Incorrect TableName");
            Assert.That(dsMethod.Tables[1].TableName, Is.EqualTo("second"), "Incorrect TableName");
            Assert.That(dsMethod.Tables[2].TableName, Is.EqualTo("third"), "Incorrect TableName");

            DataTable check1 = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls2.xls", "first", true);
            DataTable check2 = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls2.xls", "second", true);
            DataTable check3 = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls2.xls", "third", true);

            Assert.That(SqlHelper.CompareDataTables(check1, dsMethod.Tables[0]), Is.True, "Data does not match");
            Assert.That(SqlHelper.CompareDataTables(check2, dsMethod.Tables[1]), Is.True, "Data does not match");
            Assert.That(SqlHelper.CompareDataTables(check3, dsMethod.Tables[2]), Is.True, "Data does not match");
        }

        /// <summary>
        /// Scenario: Work Sheet names retrieved for a workbook containing a single sheet
        /// Expected: string array of length 1 returned containing the expected name
        /// </summary>
        [Test]
        public void _013_GetWorkSheetNames_XLSX_SingleSheet()
        {
            string[] names = ExcelOleDb.GetWorkSheetNames(OfficeFileFormat.Excel2007, @"TestData\xls1.xlsx");
            Assert.That(names.Length, Is.EqualTo(1), "Expected 1");
            Assert.That(names[0], Is.EqualTo("Sheet1"), "Incorrect value");
        }

        /// <summary>
        /// Scenario: Work sheet names retrieved for a workbook containing more than one sheets
        /// Expected: string array of appropriate length containing the correct names
        /// </summary>
        [Test]
        public void _014_GetWorkSheetNames_XLSX_ManySheets()
        {
            string[] names = ExcelOleDb.GetWorkSheetNames(OfficeFileFormat.Excel2007, @"TestData\xls2.xlsx");
            Assert.That(names.Length, Is.EqualTo(3), "Expected 3");
            Assert.That(names[0], Is.EqualTo("first"), "Incorrect value");
            Assert.That(names[1], Is.EqualTo("second"), "Incorrect value");
            Assert.That(names[2], Is.EqualTo("third"), "Incorrect value");
        }

        /// <summary>
        /// Scenario: Attempt made to retrieve a worksheet from a valid workbook - sheet does not exist
        /// Expected: Exception (Unable to open the specified worksheet)
        /// </summary>
        [Test]
        public void _015_GetWorkSheet_XLSX_NonExistant()
        {
            Assert.That(delegate { ExcelOleDb.GetWorkSheet(OfficeFileFormat.Excel2007, @"TestData\xls1.xlsx", "wibble"); }, Throws.InstanceOf<Exception>().With.Message.EqualTo("Unable to open the specified worksheet"));
        }

        /// <summary>
        /// Scenario: Existing worksheet within a valid workbook retrieved
        /// Expected: DataTable accurately representing the data within the XLS sheet
        /// </summary>
        [Test]
        public void _016_GetWorkSheet_XLSX_Valid()
        {
            DataTable method = ExcelOleDb.GetWorkSheet(OfficeFileFormat.Excel2007, @"TestData\xls1.xlsx", "Sheet1");
            DataTable check = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls1.xls", "Sheet1", true);

            Assert.That(SqlHelper.CompareDataTables(check, method), Is.True, "Data does not match");
        }

        /// <summary>
        /// Scenario: Workbook containing a single sheet opened
        /// Expected: DataSet containing a single (appropriately named) DataTable which in turn contains the data from the worksheet
        /// </summary>
        [Test]
        public void _017_GetWorkBook_XLSX_SingleSheet()
        {
            DataSet dsMethod = ExcelOleDb.GetWorkBook(OfficeFileFormat.Excel2007, @"TestData\xls1.xlsx");
            Assert.That(dsMethod.Tables.Count, Is.EqualTo(1), "Incorrect count");
            Assert.That(dsMethod.Tables[0].TableName, Is.EqualTo("Sheet1"), "Incorrect TableName");

            DataTable check = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls1.xls", "Sheet1", true);
            Assert.That(SqlHelper.CompareDataTables(check, dsMethod.Tables[0]), Is.True, "Data does not match");
        }

        /// <summary>
        /// Scenario: Workbook containing more than one sheet opened
        /// Expected: DataSet contains one DataTable for each sheet in the workbook, each with the appropriate data
        /// </summary>
        [Test]
        public void _018_GetWorkBook_XLSX_ManySheets()
        {
            DataSet dsMethod = ExcelOleDb.GetWorkBook(OfficeFileFormat.Excel2007, @"TestData\xls2.xlsx");
            Assert.That(dsMethod.Tables.Count, Is.EqualTo(3), "Incorrect count");
            Assert.That(dsMethod.Tables[0].TableName, Is.EqualTo("first"), "Incorrect TableName");
            Assert.That(dsMethod.Tables[1].TableName, Is.EqualTo("second"), "Incorrect TableName");
            Assert.That(dsMethod.Tables[2].TableName, Is.EqualTo("third"), "Incorrect TableName");

            DataTable check1 = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls2.xls", "first", true);
            DataTable check2 = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls2.xls", "second", true);
            DataTable check3 = ExcelHelper.LoadXLSSheetToDataTable(@"TestData\xls2.xls", "third", true);

            Assert.That(SqlHelper.CompareDataTables(check1, dsMethod.Tables[0]), Is.True, "Data does not match");
            Assert.That(SqlHelper.CompareDataTables(check2, dsMethod.Tables[1]), Is.True, "Data does not match");
            Assert.That(SqlHelper.CompareDataTables(check3, dsMethod.Tables[2]), Is.True, "Data does not match");
        }
    }
}
