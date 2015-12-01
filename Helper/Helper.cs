# region Comments
/* 29-jun-2011
    1. Added dispose in finally
    2. Added a property for ScreenShotDirectory
    3. Added Global Variable for Temporary Folder
    4. Added check for existence of ScreenShotDirectory
 * */
/* 01-july-2011
   1. Added GetTestData and GetExcelConnection functions
 * */
/* 20 -july-2011
    1. Added new Class TestDataManagement
    
 * */
/* 25 -july-2011
    1. Added new Class ReporterManagement
    2. Added to Source Control
    
 * */
/* 28 -july-2011
    1. Added new Class Reporter
    
    
 * */
/* 04 -Aug-2011
    1. Rectified Infinite looping
    
    
 * */
/* 04 -Aug-2011 4:13 PM
    1. Added LogManagement class
    
    
 * */
/* 09 -Aug-2011 
    1. Added Property for TesdataFile and TestCase
    
    
 * */
/* 26 -Aug-2011 11:32 AM
    1. Updated with review comments from Team
    
    
 * */
/* 26 -Aug-2011 03:32 PM
    1. Changed GetExcelConnection method from public to internal
    
    
 * */
/* 05 -Jan-2012 03:32 PM
    1. deleted con.open() from all functions.
    
    
 * */
/* 27 -Jan-2012 03:32 PM
    1. deleted testcase from GetReporter Object.
    
    
 * */
/* 13 -Feb-2012 03:22 PM
    1. Added ComSupportFunctions
    2. Created new ResultTable Property
    
    
 * */
/* 22 -March-2012 10:23 AM
    1. Changed Access Modifier of Reporter class to internal
    2. Changed CompareData and GenerateReport methods according to changeset 1
    
    
 * */
/* 26 -March-2012 3:34 PM
    1. Introduced Versioning
    
    
 * */

/* 29 -March-2012 10:00 AM
    1. Changed Get Verification Data and Get verification Data Form returns datatable
    
    
 * */
/* 06 -April-2012 3:19 PM
    1. Added new method CreateCustomLog in log managament
    
    
 * */
/* 09 -April-2012 11:14 AM
   1. Changed Compare Data Form and Compare Data methods ( If Expected Data is Null then it is not printed in the reports)
    
    
 * */
/* 11 -May-2012 12:09 PM
   1. Changed ComCreateEmptyActualDataTable method
    
    
 * */


/* 14 -May-2012 12:09 PM
   1. Added comments and logging for CompareData,GetTestData,GetVerificationDataForm,UpdateReporterSheet
    
    
 * */

/* 15 -May-2012 12:09 PM
   1. Completed adding comments for all methods for Class TestDataManagement
 * 2. Also added logging 
    
    
 * */
/* 15 -May-2012 12:09 PM
   1. Removed reference to System.Configuration
   2. Made HelperLogPath internal method of class LogManagement
    
 * */
/* 16 -May-2012 12:09 PM
   1. Added Comments to class LogManagement
   2. Added new methods ComClearActualData,ComClearExpectedlData,ComClearTemplateData and ComClearResultData
    
 * */

/* 01 -June-2012 
   1. Added new line to function CreateCustomLog
    
 * */
/* 04 -June-2012 
   1. Added function FormatTime in class LogManagement
    
 * */
/* 10-Sep-2013
    1. Added Condition to check Actual Result also dont have a value when Expected Data doesn't have a value in compare data
    2. Added Condition to check Actual Result also dont have a value when Expected Data doesn't have a value in compareDataForm

    
 * */
/* 03-Oct-2013
    1. Added connection string to connect csv

    
 * */
/* 07-Oct-2013
    1. Changed GetTestData and GetVerificationData to get data from csv 
    2. Added tostring() conversion in CompareData and CompareDataForm functions
    
 * */
/* 23-Jan-2014
    1. Changed GetTestData and GetVerificationDataForm to get data from database 
    
/* 28-Feb-2014
 *  1. Added Fucntion CheckAreequal for comparing data based on their datatypes -- Prasanna
 * */
/* 5-Mar-2014
 *  1. Added Fucntion for sending emails in Reporter Class -- Prasanna
 * */
/* 8-Apr-2014
* 1. Added Condition IsColumnPresent to check whether column by that name exists before adding it to datatable
*/

#region 15-July2014
/*
 * Added code for capturing the [RowCount Mismatch] in functions CompareData and CompareData
 */
#endregion
#region 18-July2014
/*
 * Added code in function comparedata and CompareData and Comparedataform to handle scenario where actual row count is > expected rowcount
 */
# endregion
#endregion Main
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices.ComTypes;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Drawing.Design;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Data.Sql;
using System.Data.SqlClient;


namespace Helper
{

    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class GetTestDataErrorManagement
    {
        public string ScreenShotDirectory { get; set; }
        private string TemporaryFolder = "Temp";


        private string ScreenShotFileName(string scenario)
        {

            string d = DateTime.Now.ToString();
            d.Trim();
            string v = d.Replace("/", "_");
            string v1 = v.Replace(":", "_");
            string v2 = v1.Replace(" ", "_");
            return scenario + "_" + v2 + "_ErrorImage.jpg";
        }
        public void SaveClipBoardImagetoFile(string scenario, string errorMessage)
        {
            try
            {
                if (Directory.Exists(ScreenShotDirectory) == false)
                {
                    Directory.CreateDirectory(ScreenShotDirectory);
                }
                string lastCharacter = ScreenShotDirectory.Substring(ScreenShotDirectory.Length - 1, 1);
                if (lastCharacter != @"\")
                {
                    ScreenShotDirectory = ScreenShotDirectory + @"\";
                }

                Directory.CreateDirectory(ScreenShotDirectory + TemporaryFolder);
                string fileName = ScreenShotFileName(scenario);

                if (Clipboard.GetDataObject() != null)
                {
                    System.Windows.Forms.IDataObject data = Clipboard.GetDataObject();

                    if (data == null) new ArgumentNullException("data");

                    if (data.GetDataPresent(DataFormats.Bitmap))
                    {

                        Bitmap image = (Bitmap)data.GetData(DataFormats.Bitmap, true);
                        try
                        {


                            image.Save(ScreenShotDirectory + TemporaryFolder + @"\" + fileName, System.Drawing.Imaging.ImageFormat.Jpeg);

                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Error in function SaveClipBoardImagetoFile" + System.Environment.NewLine + ex.Message);
                        }
                        finally
                        {
                            image.Dispose();
                        }

                    }
                    else
                    {
                        throw new Exception("The Data in Clipboard is not as image format");
                    }
                }
                else
                {
                    throw new Exception("The Clipboard was empty");
                }

                WriteNoteOnImage(fileName, scenario, errorMessage);
            }
            catch (Exception ex)
            {
                throw new Exception("Error in function SaveClipBoardImagetoFile" + System.Environment.NewLine + ex.Message);
            }
        }
        private void WriteNoteOnImage(string fileName, string scenerio, string error)
        {

            Bitmap bitMap = new Bitmap(ScreenShotDirectory + TemporaryFolder + @"\" + fileName);
            Graphics g = Graphics.FromImage(bitMap);
            try
            {
                Pen pen = new Pen(Brushes.BlueViolet, 5);
                g.DrawRectangle(pen, 400, 400, 600, 250);
                StringFormat strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Center;
                g.DrawString(error, new Font("Tahoma", 11, FontStyle.Bold), Brushes.DarkRed, new RectangleF(500, 450, 500, 500), strFormat);
                g.Save();
                string fp = "";
                fp = ScreenShotDirectory + fileName;
                bitMap.Save(fp);


            }
            catch (Exception ex)
            {
                throw new Exception("Error in function WriteNoteOnImage " + System.Environment.NewLine + ex.Message);
            }
            finally
            {
                bitMap.Dispose();
                g.Dispose();
                string[] filePaths = Directory.GetFiles(ScreenShotDirectory + TemporaryFolder);
                foreach (string files in filePaths)
                    File.Delete(files);
                Directory.Delete(ScreenShotDirectory + TemporaryFolder);
            }


        }


    }
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class TestDataManagement
    {
        LogManagement log = new LogManagement();
        private string _excelConnectionPrefix = @"Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DriverId=790;ReadOnly=0;Dbq=";
        private string csvFile = "";
        private string csvFolder = "";
        string tableType = "NonUniform";
        private string _testDataFile = "";
        private string _testCase = "";
        private DataTable _template = new DataTable();
        private DataTable _expectedData = new DataTable();
        private DataTable _structure = new DataTable();
        private DataTable _data = new DataTable();
        private DataTable _resultTable;
        public DataTable ResultTable
        {
            get { return _resultTable; }
            set { _resultTable = value; }
        }
        private DataTable _actualData = new DataTable();
        public DataTable Data
        {
            get { return _data; }
            set { _data = value; }
        }
        public DataTable Template
        {
            get { return _template; }
            set { _template = value; }
        }
        public DataTable ExpectedData
        {
            get { return _expectedData; }
            set { _expectedData = value; }
        }
        public DataTable ActualData
        {
            get { return _actualData; }
            set { _actualData = value; }
        }
        public DataTable Structure
        {
            get { return _structure; }
            set { _structure = value; }
        }
        public string TestDataFile
        {
            get { return _testDataFile; }
            set { _testDataFile = value; }
        }
        public string TestCase
        {
            get { return _testCase; }
            set { _testCase = value; }
        }
        public DataRow dataRow;
        /// <summary>
        ///  This method compares data present in the Actual and Expected Data tables and store the outcome [Pass/Fail] in a result table
        /// </summary>
        public void CompareData()
        {

            try
            {
                int actualRowcount = 0;

                Console.WriteLine("Inside Compare Data");
                string filePath = LogManagement.HelperLogPath();

                log.CreateCustomLog(filePath, "[Starting CompareData]");

                DataTable dt = new DataTable();


                dt.Columns.Add("FieldName");
                dt.Columns.Add("ExpectedValue");
                dt.Columns.Add("ActualValue");
                dt.Columns.Add("Result");
                dt.Columns.Add("DateTime");

                actualRowcount = ActualData.Rows.Count;
                // This is to handle conditions where actual rows is greater and the code fails to find the corresponding row in expected
                if (actualRowcount > ExpectedData.Rows.Count)   
                    actualRowcount = ExpectedData.Rows.Count;

                log.CreateCustomLog(filePath, "\t" + "Added default columns");
                if (ExpectedData.Rows.Count != ActualData.Rows.Count)
                {
                    log.CreateCustomLog(filePath, "Actual data and Expected data rows are not matching");

                    DataRow drMisMatch = dt.NewRow();


                    drMisMatch["FieldName"] = "RowCount Mismatch";
                    drMisMatch["ExpectedValue"] = 2;
                    drMisMatch["ActualValue"] = 1;
                    drMisMatch["DateTime"] = DateTime.Now.ToLocalTime().ToString();
                    drMisMatch["Result"] = "Fail";
                    dt.Rows.Add(drMisMatch);
                    log.CreateCustomLog(filePath, "Added row mismatch record for comparsion");
                }

                //for (int j = 0; j < ActualData.Rows.Count; j++)
                for (int j = 0; j < actualRowcount; j++)
                {

                    foreach (DataColumn column in ActualData.Columns)
                    {
                        DataRow dr = dt.NewRow();

                        
                        string actualValue = "";
                        log.CreateCustomLog(filePath, "\t" + "Column Name: " + column.ColumnName);
                        if (ActualData.Rows[j][column.ColumnName] is System.DBNull == false)
                        {
                            actualValue = (string)ActualData.Rows[j][column.ColumnName].ToString();
                        }
                        else
                        {
                            actualValue = "";
                        }

                        string expectedValue = "NoData";

                        if (IsColumnPresent(column.ColumnName, ExpectedData))
                        {
                            if (ExpectedData.Rows[j][column.ColumnName] != DBNull.Value && ExpectedData.Rows[j][column.ColumnName] != null && ExpectedData.Rows[j][column.ColumnName].ToString().Length != 0)
                            {
                                log.CreateCustomLog(filePath, " Expected is surely  ");
                                if (ExpectedData.Rows.Count >= j + 1)
                                {

                                    expectedValue = (string)ExpectedData.Rows[j][column.ColumnName].ToString();

                                }


                                dr["FieldName"] = column.ColumnName;
                                dr["ExpectedValue"] = expectedValue;
                                dr["ActualValue"] = actualValue;
                                dr["DateTime"] = DateTime.Now.ToLocalTime().ToString();
                                string result = "Fail";
                                //  if (actualValue.ToLower().Trim() == expectedValue.ToLower().Trim())
                                if (CheckAreEqual(actualValue.ToLower().Trim(), expectedValue.ToLower().Trim()))
                                {
                                    result = "Pass";
                                }
                                dr["Result"] = result;



                                dt.Rows.Add(dr);
                                log.CreateCustomLog(filePath, "\t" + "Compared Colum: " + column.ColumnName);
                                Console.WriteLine(expectedValue);
                                Console.WriteLine(actualValue);
                            }
                            else
                            {
                                expectedValue = "";

                                if (actualValue.Length > 0 && expectedValue.Length > 0)

                                //if (  (actualValue != null) && (actualValue != "") && ( actualValue != DBNull.Value))
                                {
                                    log.CreateCustomLog(filePath, " Actaul is surly is surely not null ");
                                    dr["FieldName"] = column.ColumnName;
                                    dr["ExpectedValue"] = expectedValue;
                                    dr["ActualValue"] = actualValue;
                                    dr["DateTime"] = DateTime.Now.ToLocalTime().ToString();
                                    string result = "Fail";
                                    // if (actualValue.ToLower().Trim() == expectedValue.ToLower().Trim())
                                    if (CheckAreEqual(actualValue.ToLower().Trim(), expectedValue.ToLower().Trim()))
                                    {
                                        result = "Pass";
                                    }
                                    dr["Result"] = result;
                                    dt.Rows.Add(dr);
                                    log.CreateCustomLog(filePath, "\t" + "Compared Colum: " + column.ColumnName);
                                    Console.WriteLine(expectedValue);
                                    Console.WriteLine(actualValue);
                                }
                                else
                                {
                                    log.CreateCustomLog(filePath, " Do nothing as both expected and actaula ere blank ");
                                }

                            }
                        }

                    }

                }
                _resultTable = dt;
                log.CreateCustomLog(filePath, "[Completed execution of CompareData]");
            }
            catch (Exception ex)
            {
                //todo add logging comments
                throw new Exception(" Error in function CompareData: " + System.Environment.NewLine + ex.Message);
            }
        }
        public void CompareDataForm()
        {
            int actualRowcount = 0;

           
            string filePath = LogManagement.HelperLogPath();

            log.CreateCustomLog(filePath, "[Starting CompareData]");

            DataTable dt = new DataTable();


            dt.Columns.Add("FieldName");
            dt.Columns.Add("ExpectedValue");
            dt.Columns.Add("ActualValue");
            dt.Columns.Add("Result");
            dt.Columns.Add("DateTime");

            actualRowcount = ActualData.Rows.Count;
            // This is to handle conditions where actual rows is greater and the code fails to find the corresponding row in expected
            if (actualRowcount > ExpectedData.Rows.Count)
                actualRowcount = ExpectedData.Rows.Count;

            if (ExpectedData.Rows.Count != ActualData.Rows.Count)
            {
                log.CreateCustomLog(filePath, "Actual data and Expected data rows are not matching");

                DataRow drMisMatch = dt.NewRow();


                drMisMatch["FieldName"] = "RowCount Mismatch";
                drMisMatch["ExpectedValue"] = 2;
                drMisMatch["ActualValue"] = 1;
                drMisMatch["DateTime"] = DateTime.Now.ToLocalTime().ToString();
                drMisMatch["Result"] = "Fail";
                dt.Rows.Add(drMisMatch);
                log.CreateCustomLog(filePath, "Added row mismatch record for comparsion");
            }

            for (int j = 0; j < actualRowcount; j++)
            {

                foreach (DataColumn column in ActualData.Columns)
                {
                    DataRow dr = dt.NewRow();

                    string actualValue = "";
                    if (ActualData.Rows[j][column.ColumnName] is System.DBNull == false)
                    {
                        actualValue = (string)ActualData.Rows[j][column.ColumnName].ToString();
                    }
                    else
                    {
                        actualValue = "";
                    }
                    string expectedValue = "NoData";
                    if (ExpectedData.Rows[j][column.ColumnName] != DBNull.Value && ExpectedData.Rows[j][column.ColumnName] != null && ExpectedData.Rows[j][column.ColumnName].ToString().Length != 0)
                    {

                        expectedValue = (string)ExpectedData.Rows[j][column.ColumnName].ToString();


                        dr["FieldName"] = column.ColumnName;
                        dr["ExpectedValue"] = expectedValue;
                        dr["ActualValue"] = actualValue;
                        dr["DateTime"] = DateTime.Now.ToLocalTime().ToString();
                        string result = "Fail";
                        //  if (actualValue.ToLower().Trim() == expectedValue.ToLower().Trim())
                        if (CheckAreEqual(actualValue.ToLower().Trim(), expectedValue.ToLower().Trim()))
                        {
                            result = "Pass";
                        }
                        dr["Result"] = result;
                        dt.Rows.Add(dr);
                    }
                    else
                    {
                        expectedValue = "";
                        if (actualValue.Length > 0 && expectedValue.Length > 0)
                        {

                            log.CreateCustomLog(filePath, " Actaul is surley is surely not null ");
                            dr["FieldName"] = column.ColumnName;
                            dr["ExpectedValue"] = expectedValue;
                            dr["ActualValue"] = actualValue;
                            dr["DateTime"] = DateTime.Now.ToLocalTime().ToString();
                            string result = "Fail";
                            if (actualValue.ToLower().Trim() == "")
                            {
                                result = "Pass";
                            }
                            dr["Result"] = result;



                            dt.Rows.Add(dr);
                            log.CreateCustomLog(filePath, "\t" + "Compared Colum: " + column.ColumnName);
                            Console.WriteLine(expectedValue);
                            Console.WriteLine(actualValue);
                        }
                    }

                }


            }
            _resultTable = dt;
            log.CreateCustomLog(filePath, "[Completed execution of CompareData]");
        }
        public bool IsColumnPresent(string colname, DataTable dtbl)
        {
            try
            {
                bool IsColumnPresent = false;
                string icolName = "";
                for (int ic = 0; ic < dtbl.Columns.Count; ic++)
                {

                    icolName = dtbl.Columns[ic].Caption.ToString();
                    if (colname.Trim().ToLower() == icolName.Trim().ToLower())
                    {
                        IsColumnPresent = true;
                        break;
                    }
                }
                return IsColumnPresent;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }
        internal OdbcConnection GetExcelConnection(string dataFile)
        {
            string _conString = "";
            try
            {

                Console.WriteLine("Inside GetExcel Connection");
                string extension = Path.GetExtension(dataFile);
                string test = dataFile.Replace(@"\", ",");
                string[] words = test.Split(new[] { ',' });
                int count = words.Length;
                csvFolder = null;
                csvFile = words[count - 1];
                for (int s = 0; s < words.Length - 1; s++)
                {
                    if (s == 0)
                    {
                      //  csvFolder = words[s];
                        csvFolder = words[s] + @"\";
                    }
                    else
                    {
                        csvFolder = csvFolder + @"\" + words[s];
                    }
                }
                Console.WriteLine("Csv folder " + csvFolder);
                if (extension.ToLower().Trim() == ".xls" || extension.ToLower().Trim() == ".xlsx")
                {
                    _excelConnectionPrefix = @"Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DriverId=790;ReadOnly=0;Dbq=";
                    _conString = _excelConnectionPrefix + dataFile;
                }
                else if (extension.ToLower().Trim() == ".csv")
                {
                    //_excelConnectionPrefix = @"Driver={Microsoft Text Driver (*.txt; *.csv)};Extensions=asc,csv,tab,txt;Dbq=";

                    _excelConnectionPrefix = @"Driver={Microsoft Text Driver (*.txt; *.csv)};extensions=csv;Dbq=";
                    _conString = _excelConnectionPrefix + csvFolder;
                    //"Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" _& csvFileFolder & ";Extended Properties=""Text;HDR=No;FMT=Delimited"""
                }
                //string test = dataFile.Replace(".csv", "");

                Console.WriteLine("CSV conn string " + _conString);

                var _con = new OdbcConnection(_conString);
                _con.Open();
                Console.WriteLine(_con.State.ToString());
                return _con;
            }
            catch (Exception ex)
            {
                //todo add logging comments
                throw new Exception(" Error in function GetExcelConnection: " + System.Environment.NewLine + ex.Message);
            }

        }
        /// <summary>
        ///  This method retrieves the data from the data sheet of the Excel data file for a particular test case
        /// </summary>
        /// <param name="testDataFile">File name containing the test data</param>
        /// <param name="testCase">Test case whose data needs to be retrieved. Leave blank if all records need to be retrieved</param>
        public void GetTestData(string testDataFile, string testCase)
        {
            OdbcConnection con = null;
            SqlConnection con2 = new System.Data.SqlClient.SqlConnection();
            DataSet dsat = new DataSet();
            _testDataFile = testDataFile;
            _testCase = testCase;
            try
            {
                string extension = Path.GetExtension(testDataFile);
                if (extension.ToLower().Trim() == ".xls" || extension.ToLower().Trim() == ".xlsx")
                {

                    string filePath = LogManagement.HelperLogPath();
                    log.CreateCustomLog(filePath, "[Starting GetTestData]");
                    con = GetExcelConnection(testDataFile);
                    log.CreateCustomLog(filePath, "\t" + "Opened excel file:" + testDataFile);
                    var _command = new OdbcCommand();

                    _data.Clear();
                    _structure.Clear();
                    log.CreateCustomLog(filePath, "\t" + "Cleared old data from structure and data sheet");
                    _command.Connection = con;
                    if (testCase.Length == 0)
                        _command.CommandText = "SELECT * FROM [Data$] WHERE InputData='Y'";
                    else
                        _command.CommandText = "SELECT * FROM [Data$] WHERE TestCase='" + testCase + "' AND InputData='Y'";
                    var dt = new OdbcDataAdapter(_command);

                    //_data.RowChanged += new DataRowChangeEventHandler(Row_Changed);
                    dt.Fill(_data);
                    log.CreateCustomLog(filePath, "\t" + "Got Data from Data Sheet");
                    var _command1 = new OdbcCommand();
                    _command1.Connection = con;
                    _command1.CommandText = "SELECT * FROM [Structure$]";
                    log.CreateCustomLog(filePath, "\t" + "Got Data from Structure Sheet");
                    var dt1 = new OdbcDataAdapter(_command1);
                    dt1.Fill(_structure);
                    Data = _data;
                    log.CreateCustomLog(filePath, "[Completed execution of GetTestData]");
                }
                else if (extension.ToLower().Trim() == ".csv")
                {
                    string filePath = LogManagement.HelperLogPath();
                    Console.WriteLine("Inside CSV Gettest data");
                    log.CreateCustomLog(filePath, "[Starting GetTestData]");
                    con = GetExcelConnection(testDataFile);
                    Console.WriteLine("outide get csv conection");
                    log.CreateCustomLog(filePath, "\t" + "Opened csv file:" + testDataFile);
                    var _command = new OdbcCommand();

                    _data.Clear();
                    log.CreateCustomLog(filePath, "\t" + "Cleared old data from data sheet");
                    _command.Connection = con;
                    Console.WriteLine("csvfile" +   csvFile );
                    _command.CommandText = "SELECT * FROM " + csvFile;

                    var DT1 = new OdbcDataAdapter(_command);
                    DT1.Fill(_data);
                    Data = _data;
                }
                else
                {
                    string filePath = LogManagement.HelperLogPath();
                    log.CreateCustomLog(filePath, "[Starting GetTestData]");
                    con2.ConnectionString = "Data Source=WDPS022C;Initial Catalog=UIAutomationK2;User ID=sa;Password=Pass@123";
                    con2.Open();
                    _data.Clear();
                    _structure.Clear();
                    SqlCommand commandat = new SqlCommand();
                    commandat.Connection = con2;
                    commandat.CommandType = CommandType.StoredProcedure;
                    commandat.CommandText = "ReturnDataSet";
                    commandat.Parameters.Add(new SqlParameter("@ScreenName", testDataFile));
                    commandat.Parameters.Add(new SqlParameter("@testcase", testCase));

                    SqlDataAdapter adapter2 = new SqlDataAdapter(commandat);
                    adapter2.Fill(dsat);
                    _structure = dsat.Tables[0];
                    _data = dsat.Tables[1];
                    Data = _data;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(" Error in function GetTestData: " + testDataFile + "    " + ex.Message);
                //todo add logging comments
                throw new Exception(" Error in function GetTestData: " + System.Environment.NewLine + ex.Message);
            }
            finally
            {
                if (con != null && con.State.ToString() == "Open")
                {
                    con.Close();
                    con.Dispose();
                }
                if (con2 != null && con2.State.ToString() == "Open")
                {
                    con2.Close();
                    con2.Dispose();
                }
            }

        }
        /// <summary>
        ///  todo
        /// </summary>
        /// <param name="expectedDataFile"></param>
        /// <param name="primaryKey"></param>
        /// <returns></returns>
        public DataTable GetVerificationData(string expectedDataFile, string primaryKey)
        {

            _testDataFile = expectedDataFile;
            _testCase = primaryKey;
            OdbcConnection _verificationCon = null;
            try
            {
                string extension = Path.GetExtension(expectedDataFile);
                if (extension.ToLower().Trim() == ".xls" || extension.ToLower().Trim() == ".xlsx")
                {

                    DataTable dtTableType = new DataTable();

                    _verificationCon = GetExcelConnection(expectedDataFile);

                    OdbcCommand _command = new OdbcCommand();

                    _command.Connection = _verificationCon;
                    _command.CommandText = "SELECT * FROM [Template$]";
                    OdbcDataAdapter dt = new OdbcDataAdapter(_command);


                    dt.Fill(_template);

                    OdbcCommand _command1 = new OdbcCommand();

                    _command1.Connection = _verificationCon;
                    _command1.CommandText = "SELECT * FROM [Data$] where PrimaryKey='" + primaryKey + "'";
                    OdbcDataAdapter dt1 = new OdbcDataAdapter(_command1);
                    dt1.Fill(_expectedData);
                    ComCreateEmptyActualDataTable();

                    tableType = (string)_template.Rows[0]["TableType"];
                    return _template;

                }
                else if (extension.ToLower().Trim() == ".csv")
                {
                    _verificationCon = GetExcelConnection(expectedDataFile);

                    OdbcCommand _command = new OdbcCommand();
                    _command.Connection = _verificationCon;
                    _command.CommandText = "SELECT * FROM " + csvFile;
                    OdbcDataAdapter dt = new OdbcDataAdapter(_command);
                    dt.Fill(_expectedData);
                    return _expectedData;
                }
                return _template;




            }
            catch (Exception ex)
            {
                //todo add logging comments
                throw new Exception(" Error in function GetVerificationData: " + System.Environment.NewLine + ex.Message);
            }
            finally
            {
                _verificationCon.Close();
                _verificationCon.Dispose();
            }

        }
        /// <summary>
        /// This method updates the column in the data sheet with the value passed
        /// </summary>
        /// <param name="testDataFile">Path of the TestData excel file</param>
        /// <param name="testCase">Test case Id of the row containing data</param>
        /// <param name="columnName">The column whose vlaue is to be updated</param>
        /// <param name="data">Value </param>
        public void UpdateTestData(string testDataFile, string testCase, string columnName, string data)
        {
            OdbcConnection con = GetExcelConnection(testDataFile);

            try
            {


                var _command = new OdbcCommand();


                _command.Connection = con;
                string space = " ";
                if (columnName.Contains(space))
                {
                    columnName = "[" + columnName + "]";
                }

                _command.CommandText = "Update [Data$] set " + columnName + "='" + data + "' where TestCase='" + testCase + "'";


                _command.ExecuteNonQuery();


            }
            catch
            {
                throw;
            }
            finally
            {
                con.Close();
                con.Dispose();
            }

        }
        /// <summary>
        ///  This method updates the column in the data sheet with the value passed
        /// </summary>
        /// <param name="expectedDataFile">Path of the template excel file</param>
        /// <param name="columnName">Name of the column that needs to be updated</param>
        /// <param name="data">Value for the columnn</param>
        ///<example>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.UpdateTemplateData();
        ///</code>
        ///</example>
        public void UpdateTemplateData(string expectedDataFile, string columnName, string data)
        {
            OdbcConnection con = GetExcelConnection(expectedDataFile);

            try
            {


                var _command = new OdbcCommand();
                _command.Connection = con;
                string space = " ";
                if (columnName.Contains(space))
                {
                    columnName = "[" + columnName + "]";
                }



                if (TestCase.Length == 0)
                    _command.CommandText = "Update [Data$] set " + columnName + "='" + data + "'";
                else
                    _command.CommandText = "Update [Data$] set " + columnName + "='" + data + "' where PrimaryKey='" + TestCase + "'";


                _command.ExecuteNonQuery();


            }
            catch
            {
                throw;
            }
            finally
            {
                con.Close();
                con.Dispose();
            }

        }
        /// <summary>
        ///  This method updates the column in the template sheet with the value passed
        /// </summary>
        /// <param name="expectedDataFile">Path of the excel file</param>
        /// <param name="columnName">Name of the column that needs to be updated</param>
        /// <param name="data">Value for the columnn</param>
        ///<example>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.UpdateTemplate();
        ///</code>
        ///</example>
        public void UpdateTemplate(string expectedDataFile, string columnName, string data)
        {
            OdbcConnection con = GetExcelConnection(expectedDataFile);

            try
            {


                var _command = new OdbcCommand();


                _command.Connection = con;
                string space = " ";
                if (columnName.Contains(space))
                {
                    columnName = "[" + columnName + "]";
                }

                _command.CommandText = "Update [Template$] set " + columnName + "='" + data + "'";


                _command.ExecuteNonQuery();


            }
            catch
            {
                throw;
            }
            finally
            {
                con.Close();
                con.Dispose();
            }

        }

        /// <summary>
        ///  This method returns the verification data for the exceldata file passed. It reads data from both structure and data sheets.
        /// </summary>
        /// <param name="expectedDataFile">Path of the excel file</param>
        /// <param name="primarykey">Testcase ID</param>
        /// <returns>Data table containing data from the structure sheet</returns>
        ///<example>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///Return testData.GetVerificationDataForm();
        ///</code>
        ///</example>
        public DataTable GetVerificationDataForm(string expectedDataFile, string primarykey)
        {
            OdbcConnection _verificationCon = null;
            SqlConnection _verificationCon2 = new System.Data.SqlClient.SqlConnection();
            DataSet dsat = new DataSet();
            string filePath = LogManagement.HelperLogPath();

            log.CreateCustomLog(filePath, "[Starting GetVerificationDataForm]");


            try
            {


                DataTable dtTableType = new DataTable();
                string extension = Path.GetExtension(expectedDataFile);
                if (extension.ToLower().Trim() == ".xls" || extension.ToLower().Trim() == ".xlsx")
                {
                    _verificationCon = GetExcelConnection(expectedDataFile);
                    log.CreateCustomLog(filePath, "\t" + "Opened expected data file:" + expectedDataFile);

                    OdbcCommand _command = new OdbcCommand();

                    _command.Connection = _verificationCon;
                    _command.CommandText = "SELECT * FROM [Structure$]";
                    OdbcDataAdapter dt = new OdbcDataAdapter(_command);


                    dt.Fill(_template);


                    OdbcCommand _command1 = new OdbcCommand();

                    _command1.Connection = _verificationCon;
                    _command1.CommandText = "SELECT * FROM [Data$] where TestCase='" + primarykey + "'";
                    OdbcDataAdapter dt1 = new OdbcDataAdapter(_command1);
                    dt1.Fill(_expectedData);
                    log.CreateCustomLog(filePath, "\t" + "Read expected data successfully");
                    ComCreateEmptyActualDataTable();
                    log.CreateCustomLog(filePath, "[Completed execution of GetVerificationDataForm]");
                    return _template;
                }
                else
                {
                    _verificationCon2.ConnectionString = "Data Source=WDPS022C;Initial Catalog=UIAutomationK2;User ID=sa;Password=Pass@123";
                    _verificationCon2.Open();

                    log.CreateCustomLog(filePath, "\t" + "Opened expected data file:" + expectedDataFile);
                    SqlCommand commandat = new SqlCommand();
                    commandat.Connection = _verificationCon2;
                    commandat.CommandType = CommandType.StoredProcedure;
                    commandat.CommandText = "ReturnDataSet";
                    commandat.Parameters.Add(new SqlParameter("@ScreenName", expectedDataFile));
                    commandat.Parameters.Add(new SqlParameter("@testcase", primarykey));
                    SqlDataAdapter adapter2 = new SqlDataAdapter(commandat);
                    adapter2.Fill(dsat);
                    _template = dsat.Tables[0];
                    _expectedData = dsat.Tables[1];
                    log.CreateCustomLog(filePath, "\t" + "Read expected data successfully");
                    ComCreateEmptyActualDataTable();
                    log.CreateCustomLog(filePath, "[Completed execution of GetVerificationDataForm]");
                    return _template;
                }


            }
            catch (Exception ex)
            {
                //todo add logging comments
                throw new Exception(" Error in function GetVerificationDataForm: " + System.Environment.NewLine + ex.Message);
            }
            finally
            {
                if (_verificationCon != null && _verificationCon.State.ToString() == "Open")
                {
                    _verificationCon.Close();
                    _verificationCon.Dispose();
                }
                if (_verificationCon2 != null && _verificationCon2.State.ToString() == "Open")
                {
                    _verificationCon2.Close();
                    _verificationCon2.Dispose();
                }

            }

        }
        /// <summary>
        ///  This method updates the column values in the excel document containing custom columns
        /// </summary>
        /// <param name="columnFileName">Path of the excel file containing custom calls</param>
        /// <param name="columnName">Column name whose value needs to be updated</param>
        /// <param name="columnValue">New value for the column that is being updated</param>
        ///<example>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.UpdateReporterSheet();
        ///</code>
        ///</example>
        public void UpdateReporterSheet(string columnFileName, string columnName, string columnValue)
        {
            OdbcConnection con = null;
            string filePath = LogManagement.HelperLogPath();
            try
            {
                log.CreateCustomLog(filePath, "[Starting UpdateReporterSheet]");

                if (File.Exists(columnFileName) == false)
                {


                    throw new Exception("File not found :" + columnFileName);
                }
                log.CreateCustomLog(filePath, "\t" + "Found custom report file");

                //string conString = _excelConnectionPrefix + columnFilePath;

                con = GetExcelConnection(columnFileName);
                log.CreateCustomLog(filePath, "\t" + "Opened expected data file:" + columnFileName);


                var command = new OdbcCommand();


                command.Connection = con;
                string space = " ";
                if (columnName.Contains(space))
                {
                    columnName = "[" + columnName + "]";
                }
                command.CommandText = "Update [Data$] set ColumnValue ='" + columnValue + "' where CustomColumn ='" + columnName + "'";
                command.ExecuteNonQuery();
                log.CreateCustomLog(filePath, "[Updated custom report file]");



            }
            catch (Exception ex)
            {

                throw new Exception("Error in Function UpdateReporterSheet" + System.Environment.NewLine + ex.Message);
            }
            finally
            {
                con.Close();
                con.Dispose();
            }

        }

        /// <summary>
        ///  This method updates the speficied column in a given excel sheet with a GUID value
        /// </summary>
        /// <param name="testDataFile">File name containing the column</param>
        /// <param name="columnName">Name of the column</param>
        /// <param name="testCase">Testcase ID to identify the unqiue row </param>
        ///<example>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.UpdateColumnValueWithGUID();
        ///</code>
        ///</example>
        public void UpdateColumnValueWithGUID(string testDataFile, string columnName, string testCase)
        {
            string filePath = LogManagement.HelperLogPath();
            log.CreateCustomLog(filePath, "[Starting UpdateColumnValueWithGUID]");

            var con = GetExcelConnection(testDataFile);

            try
            {


                var _command = new OdbcCommand();

                _command.Connection = con;

                string guidValue = (String)System.Guid.NewGuid().ToString("N").ToUpper();

                log.CreateCustomLog(filePath, "\t" + "New GUID generated");

                string space = " ";
                if (columnName.Contains(space))
                {
                    columnName = "[" + columnName + "]";
                }

                _command.CommandText = "Update [Data$] set " + columnName + "='" + guidValue + "' where TestCase='" + testCase + "'";


                _command.ExecuteNonQuery();

                log.CreateCustomLog(filePath, "[Updated column with the new GUID value]");
            }
            catch (Exception ex)
            {

                throw new Exception("Error in Function UpdateColumnValueWithGUID" + System.Environment.NewLine + ex.Message);
            }
            finally
            {
                con.Close();
                con.Dispose();
            }


        }
        /// <summary>
        ///  This method is used to read the column value for a particular test case
        /// </summary>
        /// <param name="testDataFile"></param>
        /// <param name="testCase">Path of the excel file containing </param>
        /// <param name="columnName">Name of the column whose value is to be retrieved</param>
        /// <example>
        /// <returns>Value of the column name from the Excel file</returns>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.GetColumnValue();
        ///</code>
        ///</example>
        public string GetColumnValue(string testDataFile, string testCase, string columnName)
        {

            try
            {
                string filePath = LogManagement.HelperLogPath();
                log.CreateCustomLog(filePath, "[Starting GetColumnValue]");

                GetTestData(testDataFile, testCase);
                string id = "";
                DataTable table = Data;
                id = (String)table.Rows[0][columnName];

                if (id == "")
                {
                    throw new Exception("No value found");
                }
                else
                {
                    log.CreateCustomLog(filePath, "[Read  Columnvalue]");
                    return id;
                }

            }
            catch (FileNotFoundException)
            {

                throw new Exception("File " + testDataFile + "  not found");
            }

            catch (Exception ex)
            {

                throw new Exception("Error in Function GetColumnValue" + System.Environment.NewLine + ex.Message);
            }

        }
        /// <summary>
        /// This method allows a COM client (QTP/VBScript) to create a dotnet compliant Actual Data table
        /// </summary>
        ///<example>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.ComCreateEmptyActualDataTable();
        ///</code>
        ///</example>
        public void ComCreateEmptyActualDataTable()
        {
            DataTable dt = new DataTable();
            string colNamesArray = "";
            string colNametoAdd = "";
            string filePath = LogManagement.HelperLogPath();
            log.CreateCustomLog(filePath, "[Starting ComCreateEmptyActualDataTable]");
            try
            {
                for (int k = 0; k < Template.Rows.Count; k++)
                {
                    if (Convert.IsDBNull(Template.Rows[k]["Fieldname"]) == false)
                    {
                        colNametoAdd = (string)Template.Rows[k]["Fieldname"];

                        if (colNamesArray.Contains(colNametoAdd) == false)
                        {
                            colNamesArray = colNamesArray + colNametoAdd + ";";
                            dt.Columns.Add(colNametoAdd);
                            log.CreateCustomLog(filePath, "\t" + "Added column: " + colNametoAdd);

                        }
                    }
                }
                _actualData = dt;
                log.CreateCustomLog(filePath, "[Created ComCreateEmptyActualDataTable]");
            }
            catch (Exception e)
            {
                throw new Exception("CreateEmptyActualDataTable" + System.Environment.NewLine + e.Message);
            }

        }



        private bool CheckAreEqual(string expval, string actval)
        {
            string filePath = LogManagement.HelperLogPath();
            try
            {
                if ((ParseString(expval) == dataType.System_String) && (ParseString(actval) == dataType.System_String))
                {

                    log.CreateCustomLog(filePath, " [CheckAreEqual]  Inside only string for  " + expval);
                    if (expval == actval)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else if (
                    (ParseString(expval) == (dataType.System_Double) || ParseString(expval) == (dataType.System_Int32) || ParseString(expval) == (dataType.System_Int64)
                    )
                        &&
                    (ParseString(actval) == (dataType.System_Double) || ParseString(actval) == (dataType.System_Int32) || ParseString(actval) == (dataType.System_Int64)
                    )
                    )
                // We need to compare decimal formats by double or int    
                {
                    log.CreateCustomLog(filePath, " [CheckAreEqual]  Inside only  numeric  " + expval);

                    Double numexp = Double.Parse(expval);
                    Double numact = Double.Parse(actval);
                    if (Math.Round(numexp, 8) == Math.Round(numact, 8))
                    {
                        return true;

                    }
                    else
                    {
                        return false;
                    }
                }
                else if (ParseString(expval) == (dataType.System_DateTime) && ParseString(actval) == (dataType.System_DateTime))
                {
                    log.CreateCustomLog(filePath, " [CheckAreEqual]  Inside only Date   " + expval);
                    DateTime dtexp = DateTime.Parse(expval);
                    DateTime dtact = DateTime.Parse(actval);
                    if (dtexp == dtact)
                    {
                        return true;

                    }
                    else
                    {
                        return false;
                    }
                }
                else if (ParseString(expval) == (dataType.System_Boolean) && ParseString(actval) == (dataType.System_Boolean))
                {
                    log.CreateCustomLog(filePath, " [CheckAreEqual]  Inside only Bool   " + expval);
                    Boolean boolexp = Boolean.Parse(expval);
                    Boolean boolact = Boolean.Parse(actval);
                    if (boolexp = boolact)
                    {
                        return true;

                    }
                    else
                    {
                        return false;
                    }


                }
                else
                {
                    log.CreateCustomLog(filePath, " [CheckAreEqual]  Inside none of above cases   " + expval);

                    return false;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("error in CheckAreEqual: " + ex.Message);
            }
        }

        enum dataType
        {
            System_Boolean = 0,
            System_Int32 = 1,
            System_Int64 = 2,
            System_Double = 3,
            System_DateTime = 4,
            System_String = 5
        }

        private dataType ParseString(string str)
        {

            bool boolValue;
            Int32 intValue;
            Int64 bigintValue;
            double doubleValue;
            DateTime dateValue;

            // Place checks higher if if-else statement to give higher priority to type.

            if (bool.TryParse(str, out boolValue))
                return dataType.System_Boolean;
            else if (Int32.TryParse(str, out intValue))
                return dataType.System_Int32;
            else if (Int64.TryParse(str, out bigintValue))
                return dataType.System_Int64;
            else if (double.TryParse(str, out doubleValue))
                return dataType.System_Double;
            else if (DateTime.TryParse(str, out dateValue))
                return dataType.System_DateTime;
            else return dataType.System_String;

        }
        /// <summary>
        ///  This method creates a new data row for the data table created for COM clients (QTP/VBScript)
        /// </summary>
        ///<example>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.ComCreateNewDataRow();
        ///</code>
        ///</example>
        public void ComCreateNewDataRow()
        {
            try
            {
                string filePath = LogManagement.HelperLogPath();
                log.CreateCustomLog(filePath, "[Starting ComCreateNewDataRow]");
                DataRow dr = _actualData.NewRow();
                dataRow = dr;
                log.CreateCustomLog(filePath, "[Created new datarow]");
            }
            catch (Exception e)
            {
                throw new Exception("CreateNewDataRow" + System.Environment.NewLine + e.Message);

            }

        }
        /// <summary>
        /// This method updates the columnvalue for the column in the newly created datarow
        /// </summary>
        /// <param name="columnName">Name of the column whose value is to be updated</param>
        /// <param name="columnValue">Value with which the columns needs to be updated</param>
        ///<example>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.ComAddActualDatatoRow();
        ///</code>
        ///</example>

        public void ComAddActualDatatoRow(string columnName, string columnValue)
        {
            try
            {
                string filePath = LogManagement.HelperLogPath();
                log.CreateCustomLog(filePath, "[Starting ComAddActualDatatoRow]");
                dataRow[columnName] = columnValue;
                log.CreateCustomLog(filePath, "[Updated columnvalue for the newly created datarow " + columnName + "]");
            }
            catch (Exception e)
            {
                throw new Exception("ComAddActualDatatoRow" + System.Environment.NewLine + e.Message);

            }
        }
        /// <summary>
        ///  This method adds the newly created data row to the Data table exposed to the COM Client
        ///</summary>
        ///<example>
        ///<code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.ComAddRowToActualData();
        ///</code>
        ///</example>
        public void ComAddRowToActualData()
        {
            try
            {

                string filePath = LogManagement.HelperLogPath();
                _actualData.Rows.Add(dataRow);
                log.CreateCustomLog(filePath, "[Added row to the COM table]");
            }
            catch (Exception e)
            {
                throw new Exception("ComAddRowToActualData" + System.Environment.NewLine + e.Message);

            }
        }
        /// <summary>
        /// This method is used to add row to the Actual data table. This is supposed to be used only by com clients like QTP/VBSCRIPT
        /// </summary>
        ///<example>
        /// <code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.ComClearActualData();
        ///</code>
        ///</example>
        public void ComClearActualData()
        {
            try
            {
                string filePath = LogManagement.HelperLogPath();
                log.CreateCustomLog(filePath, "[Staring ComClearActualData]");

                _actualData.Clear();
                log.CreateCustomLog(filePath, "[Cleared COM Actual Data]");
            }
            catch (Exception e)
            {
                throw new Exception("ComClearActualData" + System.Environment.NewLine + e.Message);

            }
        }
        /// <summary>
        /// This method is used to delete all records from the Template data table. This is supposed to be used only by com clients like QTP/VBSCRIPT
        /// </summary>
        ///<example>
        /// <code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.ComClearTemplateData();
        ///</code>
        ///</example>
        public void ComClearTemplateData()
        {
            try
            {
                string filePath = LogManagement.HelperLogPath();
                log.CreateCustomLog(filePath, "[Staring ComClearTemplateData]");

                _template.Clear();
                log.CreateCustomLog(filePath, "[Cleared COM Template Data]");
            }
            catch (Exception e)
            {
                throw new Exception("ComClearTemplateData" + System.Environment.NewLine + e.Message);

            }
        }
        /// <summary>
        /// This method is used to delete all records from the Result data table. This is supposed to be used only by com clients like QTP/VBSCRIPT
        /// </summary>
        ///<example>
        /// <code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.ComClearResultData();
        ///</code>
        ///</example>
        public void ComClearResultData()
        {
            try
            {
                string filePath = LogManagement.HelperLogPath();
                log.CreateCustomLog(filePath, "[Staring ComClearResultData]");

                _resultTable.Clear();
                log.CreateCustomLog(filePath, "[Cleared COM Result Data]");
            }
            catch (Exception e)
            {
                throw new Exception("ComClearResultData" + System.Environment.NewLine + e.Message);

            }
        }
        /// <summary>
        /// This method is used to delete all records from the Expected data table. This is supposed to be used only by com clients like QTP/VBSCRIPT
        /// </summary>
        ///<example>
        /// <code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///testData.ComClearExpectedData();
        ///</code>
        ///</example>
        public void ComClearExpectedData()
        {
            try
            {
                string filePath = LogManagement.HelperLogPath();
                log.CreateCustomLog(filePath, "[Staring ComClearExpectedData]");

                _expectedData.Clear();
                log.CreateCustomLog(filePath, "[Cleared COM Expected Data]");
            }
            catch (Exception e)
            {
                throw new Exception("ComClearExpectedData" + System.Environment.NewLine + e.Message);

            }
        }
    }
    /// <summary>
    /// This class contain methods and properties to generate the verification report
    /// </summary>
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class ReportsManagement
    {
        LogManagement log = new LogManagement();
        TestDataManagement data = new TestDataManagement();

        private string _reportPath;
        public string ReportPath
        {
            get { return _reportPath; }
            set { _reportPath = value; }
        }
        private DataTable _resultTable;
        /// <summary>
        /// <example>
        /// <code>
        /// </code>
        /// </example>
        /// </summary>
        public DataTable ResultTable
        {
            get { return _resultTable; }
            set { _resultTable = value; }
        }




        internal Reporter GetReporterObject(string columnFileName)
        {

            OdbcConnection con = null;
            try
            {
                if (File.Exists(columnFileName) == false)
                {
                    throw new Exception("File not found :" + columnFileName);
                }

                con = data.GetExcelConnection(columnFileName);
                // con.Open();

                DataTable datas = new DataTable();
                var command = new OdbcCommand();


                command.Connection = con;
                command.CommandText = "SELECT * FROM [Data$]";
                var dt = new OdbcDataAdapter(command);
                dt.Fill(datas);


                Reporter objReport = new Reporter(datas.Rows.Count);

                for (int j = 0; j < datas.Rows.Count; j++)
                {
                    objReport.customColumns[j] = (String)datas.Rows[j]["CustomColumn"];
                    objReport.customValues[j] = (String)datas.Rows[j]["ColumnValue"];
                }


                return objReport;


            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
                con.Dispose();
            }






        }

        /// <summary>
        /// This method is used to generate the verification report in the .csv format
        /// </summary>
        /// <param name="columnFileName">Name of the excel file containing the custom columns</param>
        ///<example> This code sample is for clients developed in C#
        /// <code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///Helper.ReportsManagement rptManager = new Helper.ReportsManagement();
        ///testData.ComClearExpectedData();
        ///testData.ComClearActualData();
        ///testData.ComClearTemplateData();
        ///testData.GetVerificationDataForm(testDataFile, testcase);
        ///testData.ActualData = GetActualDataForm(testDataFile);
        ///testData.UpdateReporterSheet(customcolumnfile, "testcase", testcase);
        ///testData.CompareDataForm();
        ///rptManager.ResultTable = testData.ResultTable;
        ///rptManager.ReportPath = resultFilePath;
        ///rptManager.GenerateReport(customcolumnfile);
        ///testData.ComClearResultData();
        /// </code>
        /// </example>
        ///<example> This code sample is for clients developed in COM client (VBScript/QTP)
        /// <code>
        ///Helper.TestDataManagement testData= new Helper.TestDataManagement();
        ///Helper.ReportsManagement rptManager = new Helper.ReportsManagement();
        ///testData.ExpectedData.Clear();
        ///testData.ActualData.Clear();
        ///testData.Template.Clear();
        ///testData.GetVerificationDataForm(testDataFile, testcase);
        ///testData.ActualData = GetActualDataForm(testDataFile);
        ///testData.UpdateReporterSheet(customcolumnfile, "testcase", testcase);
        ///testData.CompareDataForm();
        ///rptManager.ResultTable = testData.ResultTable;
        ///rptManager.ReportPath = resultFilePath;
        ///rptManager.GenerateReport(customcolumnfile);
        ///testData.ResultTable.Clear();
        /// </code>
        /// </example>

        public void GenerateReport(string columnFileName)
        {

            string filePath = LogManagement.HelperLogPath();
            log.CreateCustomLog(filePath, "[Starting GenerateReport]");

            Reporter objReport = GetReporterObject(columnFileName);
            log.CreateCustomLog(filePath, "\t" + "Created reporter Object from the columnfile: " + columnFileName);
            int k = _resultTable.Columns.Count;
            int s = k;

            for (int i = 0; i <= objReport.customColumns.Length - 1; i++)
            {
                _resultTable.Columns.Add(objReport.customColumns[i]);

            }
            log.CreateCustomLog(filePath, "\t" + "Added custom columns, count=" + _resultTable.Columns.Count.ToString());
            Console.WriteLine(_resultTable.Columns.Count.ToString());
            for (int j = 0; j < _resultTable.Rows.Count; j++)
            {
                for (int m = 0; m < objReport.customColumns.Length; m++)
                {
                    _resultTable.Rows[j][k] = objReport.customValues[m];
                    k = k + 1;
                }
                k = s;

            }
            log.CreateCustomLog(filePath, "\t" + "Added values for custom columns");

            log.CreateCustomLog(filePath, "\t" + "Starting report generation");

            using (StreamWriter writer = new StreamWriter(ReportPath, true))
            {
                if (writer.BaseStream.Length == 0)
                {
                    foreach (DataColumn column in ResultTable.Columns)
                    {


                        writer.Write('\u0022' + column.ColumnName + '\u0022' + ",");

                    }
                    writer.WriteLine();
                }
                for (int i = 0; i < ResultTable.Rows.Count; i++)
                {


                    foreach (DataColumn column in ResultTable.Columns)
                    {
                        writer.Write('\u0022' + (string)ResultTable.Rows[i][column.ColumnName] + '\u0022' + ",");

                    }
                    writer.WriteLine();
                }

            }
            log.CreateCustomLog(filePath, "[Generated the report file]");




        }

        public void sendmail(string attachmentFilenames, string ListTo)
        {
            try
            {

                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
                string[] recipients = ListTo.Split(';');
                string[] attachments = attachmentFilenames.Split(';');
                foreach (string recipient in recipients)
                {
                    message.To.Add(recipient);
                }
                message.Subject = "Automation Script Execution Summary Report";
                message.From = new System.Net.Mail.MailAddress("noreply@bugnet-vm1.com");
                message.Body = "Please find the summary of Automation execution report attached";
                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("mail2.weatherford.com");
                smtp.Port = 25;

                foreach (string attachmentFilename in attachments)
                {
                    if (System.IO.File.Exists(attachmentFilename))
                    {
                        var attachment = new System.Net.Mail.Attachment(attachmentFilename);
                        message.Attachments.Add(attachment);
                    }
                }

                smtp.Send(message);
            }
            catch (Exception ex)
            {
                throw new Exception("Error in Sending Mails.." + ex.Message);
            }
        }
    }
    internal class Reporter
    {
        public string[] customColumns;
        public string[] customValues;

        public Reporter(int columnSize)
        {
            customColumns = new string[columnSize];
            customValues = new string[columnSize];
        }



    }
    /// <summary>
    /// This Class contains methods required to perform logging in text file
    /// </summary>
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]

    public class LogManagement
    {
        public string logFileName = "";
        public void CreateExecutionLog(string fileName)
        {
            logFileName = fileName;
            string date = DateTime.Now.ToLongTimeString();
            string text = "Test Execution started on" + date + "\n" + "\n";
            File.AppendAllText(fileName, text);
        }
        public void ActionStartLog(string action)
        {
            string date = DateTime.Now.ToLongTimeString();
            File.AppendAllText(logFileName, "\n Execution of " + " " + action + "is started on" + " " + date + "\n" + "\n");
        }
        public void ActionEndLog(string action)
        {
            string date = DateTime.Now.ToLongTimeString();
            File.AppendAllText(logFileName, "\n Execution of " + " " + action + "is completed on" + " " + date + "\n" + "\n");
        }
        /// <summary>
        /// This method is used to perform logging in a text file
        /// </summary>
        /// <param name="fileName">Name of the file where the logging is to be done</param>
        /// <param name="logData">Data to be logged</param>


        public void CreateCustomLog(string fileName, string logData)
        {
            logFileName = fileName;
            File.AppendAllText(fileName, System.DateTime.Now.ToLocalTime().ToString() + ":" + logData + Environment.NewLine);

        }

        internal static string HelperLogPath()
        {
            try
            {
                string logDirectoryName = "HelperLogs";
                if (!Directory.Exists(System.IO.Directory.GetCurrentDirectory() + @"\" + logDirectoryName))
                {
                    Directory.CreateDirectory(System.IO.Directory.GetCurrentDirectory() + @"\" + logDirectoryName);
                }

                return System.IO.Directory.GetCurrentDirectory() + @"\" + logDirectoryName + @"\Log.txt";
            }
            catch (Exception e)
            {
                throw new Exception("HelperLogPath" + System.Environment.NewLine + e.Message);

            }
        }
        /// <summary>
        ///  This method is used to return the elapsed time in second, minutes and hours
        /// </summary>
        /// <param name="ts">The Timespan</param>
        /// <returns>The formatted elapsed time </returns>
        static string FormatTime(TimeSpan ts)
        {
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
            ts.Hours, ts.Minutes, ts.Seconds,
            ts.Milliseconds / 10);

            return elapsedTime;
        }

    }


}