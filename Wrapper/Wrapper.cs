#region comments
/*  Version 1.0.0.1
 * Author:Sneha
 *  17-Sep-2012 - Added keyword 'ManualIntervention'
 *  17-Sep-2012 - Added function fn_BalloonToolTip
 *  27-Sep-2012 - Added UseWhite logic using AppConfig 
*/

/*  Version 1.0.0.2
 * Author:Sneha
 *  06-Nov-2012 - Added keyword rightclick
*/
/* Version 1.0.0.3
 * Author:Prasanna
 * 16-Jan-2013 --Added Fucntion 'GetOSArchitecture' For Launching application without modifying appPath in Excel Scripts.
 * */

/* Version 1.0.0.4
 * Author: Sneha
 * 23-Aug-2013 - Added keyword 'updatestructure' 
 * 23-Aug-2013 - Added function 'UpdateStructure' for updating structure sheet at runtime.
 * 06-Sep-2013 - Added keyword 'verifydatagrid' for verifying grid which has multiple control types in a dataitem and cell doesnt have a value. It is being used in K2.
 * 24-Sep-2013 - Added keyword 'verifyformdatak2' to verify control types specific to K2 which needs clicking somewhere on scree to make the control type enable.
 * 07-Oct-2013 - Added keywords 'setexpectednactualdata', 'comparecsv', and 'upadtereportersheet'
 * 20-Nov-2013 - updated launchapplication keyword for handling dynamic wait. Arg 3 will be the window name of application.
 * 26-Nov-2013 - Added keyword 'verifywordfile' and function 'worddocverification'
 * 26-Nov-2013 - Updated keyword 'launchapplication'
 * * Author: Prasanna
 * 28-Nov-2013 - Added Try catch Statements for all keywords so that failng scripts are terminated to move to next script in sequence..
 *               along with application forced termination
 *  Author: Sneha
 *  28-Nov-2013: Added keyword 'writeingrid' and function 'writeGridContent'
 *  Author: Ashok
 *  05-Dec-2013: Added extensive Logging
 * */

#region
// 09-Dec-2014 - Added counter to count pass fail test steps and remove the message box
//15-Dec-2014 - Updated function which takes arguments from command line and if argument is not present then it executes normal sequence file.
//24-Mar-2015 - Added condition to automatically combine path with testinputdatafolder if No path is specified for Launch Keyword arg1 
//               ( indexer case of WellFlo paths like M:\NF_02.wflx" were used for launching model So now arg1 will use jsut file name 
//                 paht will be constucted relative to testinputdata folder from app.config
#endregion
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Configuration;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
//using White.Core;
//using White.Core.UIItems;
//using White.Core.UIItems.Finders;
//using White.Core.UIItems.WindowItems;
using System.Windows.Automation;
using System.Data.Odbc;
using Helper;
using System.Management.Instrumentation;
using System.Management;
using AutoItX3Lib;
using System.IO;
using System.Diagnostics;
using System.Data;
using System.Web;
using System.Web.UI;
using System.Xml.Linq;
using System.Xml;
using System.Xml.Xsl;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Imaging;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using Microsoft.Win32;


namespace WrapperEXE
{

    class Wrapper
    {
       
        static int sNo = 0;
        static string _scriptName, _strPath, _htmlReportsPath, _reportsPath, _executionPath, _resultsFile, _finalReports, _resultsSummaryFile = null;
        static int _attempt = 50;
        static string _deleteFilesPath = "";


        static void Main(string[] args)
        {


           

            string _sequcenfile = ConfigurationManager.AppSettings["sequencefile"];
            string connprefix = @"Driver={Microsoft Excel Driver (*.xls)};DriverId=790;ReadOnly=0;Dbq=";
            string connstring = connprefix + _sequcenfile;
            _strPath = ConfigurationManager.AppSettings["testinputdata"];
            _reportsPath = ConfigurationManager.AppSettings["ReportFilePath"];
            try
            {
                _htmlReportsPath = ConfigurationManager.AppSettings["HtmlReportsPath"];
            }
            catch
            {
                _htmlReportsPath = _reportsPath;
            }
            if (_htmlReportsPath == null)
            {
                _htmlReportsPath = _reportsPath;
            }
            _resultsFile = ConfigurationManager.AppSettings["ResultFile"];
            _resultsSummaryFile = ConfigurationManager.AppSettings["ResultSummaryFile"];
            string _deleteFilesPath = ConfigurationManager.AppSettings["delefilespath"];
            string processName = ConfigurationManager.AppSettings["apppath"];
            _attempt = int.Parse(ConfigurationManager.AppSettings["attempts"]);

            int pascount = 0;
            int failcount = 0;

            try
            {
                //Try to clean up older test results and log files if they exist on specified paths 
                Console.WriteLine("testinputdata:=" + _strPath);

                if (File.Exists(_strPath + "log.csv"))
                {
                    File.Delete(_strPath + "log.csv");
                }

                if (File.Exists(_reportsPath))
                {
                    File.Delete(_reportsPath);
                }

                if (File.Exists(_resultsFile))
                {
                    File.Delete(_resultsFile);
                }

                if (File.Exists(_resultsSummaryFile))
                {
                    File.Delete(_resultsSummaryFile);
                }


            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message);
            }


            try
            {
                string dateTime = DateTime.Now.ToString();
                dateTime = dateTime.Replace(":", "_");
                dateTime = dateTime.Replace("-", "_");
                string executionReportPath = _reportsPath + "ExecutionLog_" + dateTime + "\\";
                string finalHtmlReportPath = _htmlReportsPath + dateTime + "\\";

                if (Directory.Exists(executionReportPath) == false)
                {
                    Directory.CreateDirectory(executionReportPath);
                }
                if (Directory.Exists(finalHtmlReportPath) == false)
                {
                    Directory.CreateDirectory(finalHtmlReportPath);
                }
                _executionPath = executionReportPath;
                _htmlReportsPath = finalHtmlReportPath;
                if (args.Length != 0)
                {

                    _scriptName = args[0];


                    if (File.Exists(_scriptName) == false)
                    {

                        Console.WriteLine("File does not exist:  " + _scriptName);
                        failcount++;
                        Console.WriteLine("Finished Test (Errors:" + failcount + "," + " Warnings:" + pascount + ")");
                        return;

                    }
                    else
                    {
                        string sourceFilecn = _reportsPath + "Log.csv";
                        string scriptNameWithoutPath = Path.GetFileNameWithoutExtension(_scriptName);
                        if (driveFromExcel(_scriptName) == false)
                        {
                            //if (File.Exists(sourceFilecn))
                            //{
                            //    File.Copy(sourceFilecn, executionReportPath + scriptNameWithoutPath + "ereportFail.csv");
                            //    File.Delete(sourceFilecn);
                            //}
                            //string endDateTime = DateTime.Now.ToString();
                            //dateTime = endDateTime.Replace(":", "_");
                            //dateTime = endDateTime.Replace("-", "_");
                            //var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(_scriptName);
                            //var filename = Path.GetFileName(_finalReports);
                            //if (File.Exists(_finalReports))
                            //{
                            //    string reportsName = _finalReports.Replace(filename, fileNameWithoutExtension + ".csv");
                            //    string finalReportName = finalHtmlReportPath + Path.GetFileName(reportsName);
                            //    File.Copy(_finalReports, finalReportName);
                            //    File.Delete(_finalReports);
                            //}
                            TerminateProcessByForce(processName);

                            Thread.Sleep(5000);
                        }
                        else
                        {
                            //if (File.Exists(sourceFilecn))
                            //{
                            //    File.Copy(sourceFilecn, executionReportPath + scriptNameWithoutPath + "ereportPass.csv");
                            //    File.Delete(sourceFilecn);
                            //}
                            //string endDateTime = DateTime.Now.ToString();
                            //dateTime = endDateTime.Replace(":", "_");
                            //dateTime = endDateTime.Replace("-", "_");
                            //var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(_scriptName);
                            //if (File.Exists(_finalReports))
                            //{
                            //    var filename = Path.GetFileName(_finalReports);
                            //    string reportsName = _finalReports.Replace(filename, fileNameWithoutExtension + ".csv");
                            //    string finalReportName = finalHtmlReportPath + Path.GetFileName(reportsName);
                            //    File.Copy(_finalReports, finalReportName);
                            //    File.Delete(_finalReports);
                            //}
                        }
                    }
                }

                else
                {

                    if (File.Exists(_sequcenfile) == false)
                    {
                        Console.WriteLine("Sequence File does not exist:  " + _sequcenfile);
                        failcount++;
                        Console.WriteLine("Finished Test (Errors:" + failcount + "," + " Warnings:" + pascount + ")");
                        return;
                    }

                    else
                    {

                        OdbcConnection oconn = new OdbcConnection(connstring);
                        oconn.Open();
                        string cmdtxt = "Select * from [master$]";
                        OdbcCommand ocmd = new OdbcCommand(cmdtxt);
                        ocmd.Connection = oconn;
                        OdbcDataReader oreader = null;
                        oreader = ocmd.ExecuteReader();

                        while (oreader.Read())
                        {
                            string sourceFile = _reportsPath + "Log.csv";

                            if (oreader["execute"].ToString().ToLower() == "y")
                            {
                                Console.WriteLine("excute flag value=" + oreader["execute"].ToString().ToLower());
                                _scriptName = oreader["scriptname"].ToString();
                                string scriptNameWithoutPath = Path.GetFileNameWithoutExtension(_scriptName);
                                if (driveFromExcel(_scriptName) == false)
                                {
                                    if (File.Exists(sourceFile))
                                    {
                                        File.Copy(sourceFile, executionReportPath + scriptNameWithoutPath + "ereportFail.csv");
                                        File.Delete(sourceFile);
                                    }
                                    string endDateTime = DateTime.Now.ToString();
                                    dateTime = endDateTime.Replace(":", "_");
                                    dateTime = endDateTime.Replace("-", "_");
                                    var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(_scriptName);
                                    var filename = Path.GetFileName(_finalReports);
                                    if (File.Exists(_finalReports))
                                    {
                                        string reportsName = _finalReports.Replace(filename, fileNameWithoutExtension + ".csv");
                                        string finalReportName = finalHtmlReportPath + Path.GetFileName(reportsName);
                                        File.Copy(_finalReports, finalReportName);
                                        File.Delete(_finalReports);
                                    }
                                    TerminateProcessByForce(processName);
                                    Thread.Sleep(5000);
                                    continue;
                                }
                                else
                                {
                                    if (File.Exists(sourceFile))
                                    {
                                        File.Copy(sourceFile, executionReportPath + scriptNameWithoutPath + "ereportPass.csv");
                                        File.Delete(sourceFile);
                                    }
                                    string endDateTime = DateTime.Now.ToString();
                                    dateTime = endDateTime.Replace(":", "_");
                                    dateTime = endDateTime.Replace("-", "_");
                                    var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(_scriptName);
                                    if (File.Exists(_finalReports))
                                    {
                                        var filename = Path.GetFileName(_finalReports);
                                        string reportsName = _finalReports.Replace(filename, fileNameWithoutExtension + ".csv");
                                        string finalReportName = finalHtmlReportPath + Path.GetFileName(reportsName);
                                        File.Copy(_finalReports, finalReportName);
                                        File.Delete(_finalReports);
                                    }

                                }


                            }

                        }
                        oreader.Close();
                    }
                }
                string htmlReport = "";
                Console.WriteLine("trying to write final report");
                htmlReport = ExportDatatableToHtml();
                if (Directory.Exists(finalHtmlReportPath))
                {
                File.WriteAllText(finalHtmlReportPath + @"\finalReport.html", htmlReport);
                }
                else
                {
                    Console.WriteLine("No direcroty by name "+finalHtmlReportPath+"was found No results html written");
                }
            
                string executionReport = "";
                executionReport = createExecutionHtmlReport();
                if (Directory.Exists(executionReportPath))
                {
                File.WriteAllText(executionReportPath + "ExecutionLog.html", executionReport);
                }
                else
                {
                    Console.WriteLine("No direcroty by name " + executionReportPath + "was found No logs html  written");
                }

            }

            catch (Exception ex)
            {

                throw new Exception(ex.Message);

            }

            Thread.Sleep(10000);

        }


        #region Generic functions
        #region driveFromExcel
        public static Boolean driveFromExcel(string excelfilePath)
        {
            try
            {

                DataTable dtResultSummary = new DataTable();
                dtResultSummary.Columns.Add("SNO");
                dtResultSummary.Columns.Add("SCRIPTNAME");
                dtResultSummary.Columns.Add("COMMENT");
                dtResultSummary.Columns.Add("ACTION");
                dtResultSummary.Columns.Add("ARGUMENT");
                dtResultSummary.Columns.Add("RESULT");
                dtResultSummary.Columns.Add("TIMETAKEN");
                dtResultSummary.Columns.Add("MESSAGE");
                string keyWord = ""; string stpexecute = ""; string stepFrom = ""; string comment = "";
                string arg1 = ""; string arg2 = ""; string arg3 = ""; string arg4 = ""; string arg5 = "";
                Helper.TestDataManagement driverData = new Helper.TestDataManagement();
                string _epath = ConfigurationManager.AppSettings["logfile"];
                string apppath = ConfigurationManager.AppSettings["apppath"];
                string reportsSectionFile = ConfigurationManager.AppSettings["reportssectionfile"];
                string strfile = ConfigurationManager.AppSettings["driverfile"];


                Boolean returndriveFromExcel = true;
                Boolean _usewhite = Boolean.Parse(ConfigurationManager.AppSettings["usewhite"]);
                Boolean _detlog = Boolean.Parse(ConfigurationManager.AppSettings["Detaillog"]);

                //*************** 1. Launch Wellflo application ************************************

                //WPF_App.WpfAction wpfaction = new WPF_App.WpfAction();
                UIAutomation_App.UIAutomationAction uiautomation = new UIAutomation_App.UIAutomationAction();
                uiautomation._Attempts = _attempt;
                uiautomation.UseWhite = _usewhite;
                uiautomation._eLogPtah = _epath;
                uiautomation._reportsPath = _reportsPath;
                uiautomation._reportsSectionPath = reportsSectionFile;
                uiautomation.UseDetaillog = _detlog;
                Helper.TestDataManagement testdataobj = new Helper.TestDataManagement();
                Helper.LogManagement logg = new Helper.LogManagement();
                WellFloUI.MSUIAutomation wellflocomui = new WellFloUI.MSUIAutomation();
                uiautomation._testDataPath = _strPath;
                AutoItX3Lib.AutoItX3 autoit = new AutoItX3Lib.AutoItX3();
                Helper.ReportsManagement rpt = new Helper.ReportsManagement();

                driverData.GetTestData(excelfilePath, "Tmaster");

                int pascount = 0;
                int failcount = 0;

                for (int i = 0; i < driverData.Structure.Rows.Count; i++)
                {
                    System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
                    stopwatch3.Start();
                    DataRow dr = dtResultSummary.NewRow();
                    comment = driverData.Structure.Rows[i]["Comment"].ToString();
                    keyWord = driverData.Structure.Rows[i]["Keyword"].ToString();
                    stpexecute = driverData.Structure.Rows[i]["Execute"].ToString();
                    stepFrom = driverData.Structure.Rows[i]["StepFrom"].ToString();
                    arg1 = driverData.Structure.Rows[i]["arg1"].ToString();
                    arg2 = driverData.Structure.Rows[i]["arg2"].ToString();
                    arg3 = driverData.Structure.Rows[i]["arg3"].ToString();
                    arg4 = driverData.Structure.Rows[i]["arg4"].ToString();
                    arg5 = driverData.Structure.Rows[i]["arg5"].ToString();
                    dr["SCRIPTNAME"] = _scriptName;
                    dr["SNO"] = (sNo + 1).ToString();
                    dr["ACTION"] = keyWord;
                    if (comment == null)
                    {
                        dr["COMMENT"] = " ";
                    }
                    else
                    {
                        dr["COMMENT"] = comment;
                    }
                    if (arg1 == null)
                    {
                        dr["ARGUMENT"] = " ";
                    }
                    else
                    {
                        dr["ARGUMENT"] = arg1;
                    }

                    if (stpexecute.ToLower() == "y")
                    {
                        logg.CreateCustomLog(_epath, "Performing Keyword:=====>" + comment + "====================================");
                        switch (keyWord.ToLower())
                        {
                            #region Generic keywords
                            #region launchapplication
                            case "launchapplication":
                                {
                                    string repeat = new string('=', 50);
                                    logg.CreateCustomLog(_epath, repeat + " Launching Application " + arg1 + DateTime.Now.ToString() + repeat);
                                    try
                                    {

                                        if (stepFrom == "0")
                                        {

                                            if (GetOSArchitecture() == "32")
                                            {
                                                logg.CreateCustomLog(_epath, "Os is 32 bit OS ");
                                                arg1 = arg1.Replace("Program Files (x86)", "Program Files");
                                                logg.CreateCustomLog(_epath, "app path is " + arg1);
                                            }
                                            else
                                            {
                                                logg.CreateCustomLog(_epath, "Os is 64 bit OS ");
                                                if (arg1.Contains("x86") == false)
                                                {
                                                    arg1 = arg1.Replace("Program Files", "Program Files (x86)");
                                                }
                                                logg.CreateCustomLog(_epath, "app path is " + arg1);
                                            }

                                            //we are using realtive path wrt test folder for launch
                                            if (arg1.Contains("\\") == false)
                                            {
                                                arg1 = Path.Combine(ConfigurationManager.AppSettings["testinputdata"], arg1);
                                            }
                                            #region OriginalLaunch
                                            if (File.Exists(arg1) == false)
                                            {
                                                Console.WriteLine("The Argument passed from excel script -arg1" + arg1 + "Does not exist..Please verifty if paths and configurations are correct");
                                                return false;
                                            }
                                            System.Diagnostics.Process p = new System.Diagnostics.Process();
                                            p.StartInfo.FileName = arg1;
                                            Console.WriteLine("Argument is " + arg1);
                                            p.Start();
                                            int processId = p.Id;
                                            uiautomation._processId = processId;
                                            while (!p.MainWindowTitle.Contains(arg3))
                                            {
                                                Console.WriteLine("Window Title----> {0}", p.MainWindowTitle);
                                                Thread.Sleep(10);
                                                p.Refresh();

                                            }
                                            Console.WriteLine("End Window Title {0}", p.MainWindowTitle);
                                            #endregion OriginalLaunch



                                            int lastbackslashpos = arg1.LastIndexOf("\\") + 1;
                                            System.Console.WriteLine("lastbackslashpos: " + lastbackslashpos);
                                            int arg1Len = arg1.Length;
                                            System.Console.WriteLine("arg1 Length: " + arg1Len);
                                            int diff = arg1.Length - lastbackslashpos;
                                            System.Console.WriteLine("diff Length: " + diff);
                                            arg1 = arg1.Substring(lastbackslashpos, diff);
                                            if (arg1.Contains(".exe") == true)
                                            {
                                                arg1 = arg1.Replace(".exe", "");
                                            }

                                            System.Console.WriteLine("Exe or Process name to attach is: " + arg1);
                                            if (arg2 == null || arg2 == "") //use Process name from arg1 itself no need to supply from excel
                                            {
                                                //White.Core.Application app = wpfaction.GetWPFApp(arg1); // ---> Launch the Application 
                                                //uiautomation._application = app;
                                            }
                                            else //read arg2 from Script File 
                                            {
                                                //White.Core.Application app = wpfaction.GetWPFApp(arg2); // ---> Launch the Application 
                                                //uiautomation._application = app;
                                            }



                                        }
                                        else if (stepFrom == "1")
                                        {
                                            //White.Core.Application app = wpfaction.GetWPFApp(arg2); // ---> Launch the Application 
                                            //uiautomation._application = app;
                                            //wpfaction._application = app;

                                        }
                                        else
                                        {
                                            logg.CreateCustomLog(_epath, "[Wraper]:Invalid stepfrom value");
                                        }
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        repeat = new string('=', 50);
                                        logg.CreateCustomLog(_epath, repeat + " Application Launched " + DateTime.Now.ToString() + repeat);
                                        pascount++;

                                    }

                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in Keyword Launch:" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated " + ex.Message;
                                        logg.CreateCustomLog(_epath, repeat + " Failed to Launch Application" + DateTime.Now.ToString() + repeat);
                                        failcount++;
                                    }


                                    break;
                                }
                            #endregion
                            #region launchindirect
                            case "launchindirect":
                                {
                                    try
                                    {

                                        if (stepFrom == "0")
                                        {

                                            if (GetOSArchitecture() == "32")
                                            {
                                                logg.CreateCustomLog(_epath, "Os is 32 bit OS ");
                                                arg1 = arg1.Replace("Program Files (x86)", "Program Files");
                                                logg.CreateCustomLog(_epath, "app path is " + arg1);
                                            }
                                            else
                                            {
                                                logg.CreateCustomLog(_epath, "Os is 64 bit OS ");
                                                if (arg1.Contains("x86") == false)
                                                {
                                                    arg1 = arg1.Replace("Program Files", "Program Files (x86)");
                                                }
                                                logg.CreateCustomLog(_epath, "app path is " + arg1);
                                            }



                                            #region New Launch
                                            System.Diagnostics.Process p = new System.Diagnostics.Process();
                                            p.StartInfo.FileName = arg1;

                                            p.Start();

                                            Process[] myprocess = Process.GetProcessesByName(arg2);

                                            while (myprocess.Length == 0)
                                            {
                                                logg.CreateCustomLog(_epath, "Launching Main Process ");
                                                myprocess = Process.GetProcessesByName(arg2);
                                                Thread.Sleep(1000);

                                            }
                                            logg.CreateCustomLog(_epath, "Launched: " + myprocess[0].ProcessName);
                                            Process actualProcess = myprocess[0];

                                            while (!actualProcess.MainWindowTitle.Contains(arg3))
                                            {
                                                Console.WriteLine("Window Title----> {0}", actualProcess.MainWindowTitle);
                                                Thread.Sleep(10);
                                                actualProcess.Refresh();

                                            }
                                            Console.WriteLine("End Window Title {0}", actualProcess.MainWindowTitle);
                                            uiautomation._processId = actualProcess.Id;
                                            #endregion



                                            int lastbackslashpos = arg1.LastIndexOf("\\") + 1;
                                            System.Console.WriteLine("lastbackslashpos: " + lastbackslashpos);
                                            int arg1Len = arg1.Length;
                                            System.Console.WriteLine("arg1 Length: " + arg1Len);
                                            int diff = arg1.Length - lastbackslashpos;
                                            System.Console.WriteLine("diff Length: " + diff);
                                            arg1 = arg1.Substring(lastbackslashpos, diff);
                                            if (arg1.Contains(".exe") == true)
                                            {
                                                arg1 = arg1.Replace(".exe", "");
                                            }

                                            System.Console.WriteLine("Exe or Process name to attach is: " + arg1);
                                            if (arg2 == null || arg2 == "") //use Process name from arg1 itself no need to supply from excel
                                            {
                                                //White.Core.Application app = wpfaction.GetWPFApp(arg1); // ---> Launch the Application 
                                                //uiautomation._application = app;
                                            }
                                            else //read arg2 from Script File 
                                            {
                                                //White.Core.Application app = wpfaction.GetWPFApp(arg2); // ---> Launch the Application 
                                                //uiautomation._application = app;
                                            }



                                        }
                                        else if (stepFrom == "1")
                                        {
                                            //White.Core.Application app = wpfaction.GetWPFApp(arg2); // ---> Launch the Application 
                                            //uiautomation._application = app;
                                            //uiautomation._processId = app.Process.Id;
                                            //wpfaction._application = app;

                                        }
                                        else
                                        {
                                            logg.CreateCustomLog(_epath, "[Wraper]:Invalid stepfrom value");
                                        }
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                    }

                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in Keyword Launchindirect:" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated " + ex.Message;
                                    }
                                    break;
                                }
                            #endregion
                            #region closewindow
                            case "closewindow":
                                {
                                    try
                                    {
                                        Thread.Sleep(2000);
                                        SendKeys.SendWait("{ENTER}");
                                        SendKeys.SendWait("{ENTER}");
                                        SendKeys.SendWait("%{F4}");
                                        SendKeys.SendWait("{ENTER}");
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;

                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in keyword  Function CloseWindow" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;

                                    }
                                    break;
                                }
                            #endregion
                            #region inputformdata
                            case "inputformdata":
                                {
                                    string repeat = new string('=', 50);
                                    logg.CreateCustomLog(_epath, "Strpath value is :" + _strPath);
                                    logg.CreateCustomLog(_epath, repeat + " Adding Data from Excel Path:=" + _strPath + arg1 + " for test case " + arg2 + DateTime.Now.ToString() + repeat);
                                    try
                                    {
                                        uiautomation.AddData(_strPath + arg1, arg2);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;

                                    }
                                    catch (Exception ex)
                                    {
                                        failcount++;
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in keyword  Function inputformdata" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();


                                    }
                                    break;
                                }
                            #endregion
                            #region comapredoc
                            case "comparedoc":
                                {
                                    try
                                    {
                                        TestDataManagement testData = new Helper.TestDataManagement();

                                        testData.GetTestData(_strPath + arg1, "T1");
                                        DataTable expectedData = testData.Data;
                                        int tableNumber = int.Parse(arg3);
                                        //DataTable actualData = verifyReport(_strPath + arg2, tableNumber);
                                        //DataTable ResultTable = CompareData(expectedData, actualData);
                                        //Helper.ReportsManagement rppt1 = new Helper.ReportsManagement();
                                        //rppt1.ResultTable = ResultTable;
                                        //rppt1.ReportPath = arg5;
                                        //rppt1.GenerateReport(_strPath + arg4);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in keyword  Function inputformdata" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;

                                    }
                                    break;
                                }
                            #endregion
                            #region verifyfileexistence
                            case "verifyfileexistence":
                                {
                                    try
                                    {
                                        string expectedData = _strPath + arg1;
                                        DataTable ResultTable = verifyFileExistence(expectedData);
                                        Helper.ReportsManagement rppt1 = new Helper.ReportsManagement();
                                        rppt1.ResultTable = ResultTable;
                                        rppt1.ReportPath = arg5;
                                        rppt1.GenerateReport(_strPath + arg4);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in keyword  Function inputformdata" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;

                                    }
                                    break;
                                }
                            #endregion
                            #region closewindowwithtitle
                            case "closewindowwithtitle":
                                {
                                    try
                                    {
                                        //uiautomation._globalWindow = wpfaction.GetWPFWindow(arg1);
                                        uiautomation.closeGlobalWindow();
                                        Thread.Sleep(2000);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in closewindowwithtitle" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region resetuiwindow
                            case "resetuiwindow":
                                {
                                    try
                                    {
                                        uiautomation.uiAutomationWindow = null;
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in resetuiwindow" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region verifydatatable
                            case "verifydatatable":
                                {
                                    try
                                    {
                                        uiautomation.VerifyData(_strPath + arg1, arg2);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in verifydatatable" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;

                                    }
                                    break;


                                }
                            #endregion
                            #region verifyformdata
                            case "verifyformdata":
                                {
                                    try
                                    {
                                        _finalReports = arg5;
                                        string repeat = new string('=', 50);
                                        logg.CreateCustomLog(_epath, arg1);
                                        logg.CreateCustomLog(_epath, arg2);
                                        logg.CreateCustomLog(_epath, arg3);
                                        logg.CreateCustomLog(_epath, arg4);
                                        logg.CreateCustomLog(_epath, arg5);
                                        logg.CreateCustomLog(_epath, "1-5 arguments");
                                        logg.CreateCustomLog(_epath, repeat + "Verifying Data from Excel " + arg1 + "for test case " + arg2 + DateTime.Now.ToString() + repeat);
                                        uiautomation.verifyDataForm(_strPath + arg1, arg2, arg3, _strPath + arg4, arg5);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in verifydatatable" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region compareXML
                            case "comparexml":
                                {
                                    try
                                    {
                                        _finalReports = arg4;
                                        arg3 = ConfigurationManager.AppSettings["testinputdata"] + arg3;  //arr[0];
                                        logg.CreateCustomLog(_epath, "Path used for comparexml keyword" + arg1);

                                        if (arg1.Trim().Length == 0 || arg2.Trim().Length == 0 || arg3.Trim().Length == 0 || arg4.Trim().Length == 0)
                                            logg.CreateCustomLog(_epath, "Need values for arg1,arg2,arg3 and arg4");
                                        else
                                        {

                                            CompareXML(arg1, arg2, arg3, arg4, logg, _epath);
                                            stopwatch3.Stop();
                                            dr["RESULT"] = "Success";
                                            dr["MESSAGE"] = "Completed";
                                            dr["TimeTaken"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  comparexml " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message;
                                    }
                                    break;
                                }
                            #endregion
                            #region verifydataiepanetable
                            case "verifydataiepanetable":
                                {
                                    try
                                    {
                                        uiautomation.verifyDataIEPaneTable(_strPath + arg1, arg2, arg3, _strPath + arg4, arg5);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "verifydataiepanetable" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }

                                    break;
                                }
                            #endregion
                            #region manualintervention
                            case "manualintervention":
                                {
                                    try
                                    {
                                        fn_BalloonToolTip(arg1, arg2);
                                        Console.WriteLine("Please Press Enter to proceed to next step");
                                        Console.ReadLine();
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in manualintervention" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region verifywpfdatagrid
                            case "verifywpfdatagrid":
                                {
                                    try
                                    {
                                        uiautomation.VerifyDataGridContent(_strPath + arg1, arg2, arg3, _strPath + arg4, arg5);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in verifywpfdatagrid" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region verifydatagrid

                            case "verifydatagrid":
                                {
                                    try
                                    {
                                        uiautomation.VerifyDataGrid2Content(_strPath + arg1, arg2, arg3, _strPath + arg4, arg5);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in verifydatagrid" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;

                                }
                            #endregion
                            #region verify_plotdata
                            case "verify_plotdata":
                                {
                                    try
                                    {
                                        verify_plotdata(arg1, arg2);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in verify_plotdata" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region infraselectrows

                            case "infraselectrows":
                                {
                                    try
                                    {
                                        selectspecfedrows(Int32.Parse(arg1), Int32.Parse(arg2));
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in infraselectrows" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;

                                }
                            #endregion
                            #region clickbuttonwindowtitle
                            case "clickbuttonwindowtitle":
                                {
                                    try
                                    {
                                        wellflocomui.GetAppWindow(arg1);
                                        wellflocomui.ClickButton(arg2);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in clickbuttonwindowtitle" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }

                                    break;
                                }
                            #endregion
                            #region clickbuttonwindowautoid
                            case "clickbuttonwindowautoid":
                                {
                                    try
                                    {
                                        //SearchCriteria search = SearchCriteria.ByAutomationId(arg1);
                                        //Window wnd = wpfaction._application.GetWindow(search, White.Core.Factory.InitializeOption.NoCache);
                                        //var btn = wnd.Get<White.Core.UIItems.Button>(arg2);
                                        //btn.Click();
                                        //dr["RESULT"] = "Success";
                                        //dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        //pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in clickbuttonwindowautoid" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region clickbuttonchildwindowtitle
                            case "clickbuttonchildwindowtitle":
                                {
                                    try
                                    {
                                        specialApply(arg1, arg2, arg3);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in clickbuttonchildwindowtitle" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region clickcordinates
                            case "clickcordinates":
                                {
                                    try
                                    {
                                        string repeat = new string('=', 50);
                                        logg.CreateCustomLog(_epath, repeat + " Clicking coordinates " + arg1 + ";" + arg2 + DateTime.Now.ToString() + repeat);
                                        AutoItX3Lib.AutoItX3 at1 = new AutoItX3Lib.AutoItX3();
                                        int x1 = Convert.ToInt32(arg1);
                                        int y1 = Convert.ToInt32(arg2);
                                        at1.MouseClick("LEFT", x1, y1, 1);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  clickcordinates" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region createoutputimage
                            case "createoutputimage":
                                {
                                    try
                                    {
                                        createOutputImage(arg1);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;

                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  createoutputimage" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;

                                    }
                                    break;
                                }
                            #endregion
                            #region clickcordinatesdbl
                            case "clickcordinatesdbl":
                                {
                                    try
                                    {
                                        AutoItX3Lib.AutoItX3 at1 = new AutoItX3Lib.AutoItX3();
                                        int x1 = Convert.ToInt32(arg1);
                                        int y1 = Convert.ToInt32(arg2);
                                        at1.MouseClick("LEFT", x1, y1, 2);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  clickcordinates" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region rightclick
                            case "rightclick":
                                {
                                    try
                                    {
                                        AutoItX3Lib.AutoItX3 at1 = new AutoItX3Lib.AutoItX3();
                                        int x1 = Convert.ToInt32(arg1);
                                        int y1 = Convert.ToInt32(arg2);
                                        at1.MouseClick("RIGHT", x1, y1, 1);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  rightclick" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }

                                    break;
                                }
                            #endregion
                            #region wait
                            case "wait":
                                {
                                    try
                                    {
                                        logg.CreateCustomLog(_epath, "Performing Keyword:===== Wait ====================================");
                                        Console.WriteLine("Waiting in Driver script for " + arg1 + "seconds");
                                        Thread.Sleep(Int32.Parse(arg1) * 1000);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  wait" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region updatestructure
                            case "updatestructure":
                                {
                                    try
                                    {
                                        UpdateStructure(arg1, arg2, arg3, arg4, arg5);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  updatestructure" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }

                                    break;
                                }
                            #endregion
                            #region setexpectednactualdata
                            case "setexpectednactualdata":
                                {
                                    try
                                    {
                                        testdataobj.ActualData.Clear();
                                        testdataobj.ExpectedData.Clear();
                                        testdataobj.GetTestData(arg1, arg2);
                                        testdataobj.ActualData = testdataobj.Data;
                                        testdataobj.ExpectedData = testdataobj.GetVerificationData(arg3, arg4);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  setexpectednactualdata" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }

                                    break;
                                }
                            #endregion
                            #region comparecsv
                            case "comparecsv":
                                {
                                    try
                                    {
                                        testdataobj.CompareData();
                                        logg.CreateCustomLog(_epath, "Compare Data Finished");
                                        rpt.ResultTable = testdataobj.ResultTable;
                                        rpt.ReportPath = arg2;
                                        logg.CreateCustomLog(_epath, "Trying to create Report");
                                        rpt.GenerateReport(_strPath + arg1);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  comparecsv" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }


                                    break;
                                }
                            #endregion
                            #region upadtereportersheet
                            case "upadtereportersheet":
                                {
                                    try
                                    {
                                        testdataobj.UpdateReporterSheet(_strPath + arg1, "testcase", arg2);
                                        testdataobj.UpdateReporterSheet(_strPath + arg1, "section", arg3);
                                        testdataobj.UpdateReporterSheet(_strPath + arg1, "webtable", arg4);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  upadtereportersheet" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }

                                    break;
                                }
                            #endregion
                            #region copydata
                            case "copydata":
                                {
                                    try
                                    {
                                        Excel.Application xlApp = new Excel.Application();
                                        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(_strPath + arg1);
                                        Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                                        Excel.Range xlRange = xlWorksheet.UsedRange;
                                        xlRange.Copy(System.Type.Missing);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;

                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  CopyData" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region verifywordfile
                            case "verifywordfile":
                                {
                                    try
                                    {
                                        worddocverification(arg1, arg2);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  verifywordfile " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }

                                    break;

                                }

                            #endregion
                            #region writeingrid
                            case "writeingrid":
                                {
                                    try
                                    {
                                        writeGridContent(arg1, arg2);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  writeingrid " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }
                            #endregion
                            #region Compare Excel
                            case "compareexcel":
                                {
                                    try
                                    {
                                        CompareExcel(arg1, arg2);
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        pascount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  writeingrid " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        failcount++;
                                    }
                                    break;
                                }

                            #endregion

                            #region Lowis related keywords
                            #region deletefile
                            case "deletefile":
                                {
                                    try
                                    {


                                        DeleteFile(arg1, logg, _epath);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();

                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  DeleteFile " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                    }
                                    break;
                                }
                            #endregion
                            #region deletefiles
                            case "deletefiles":
                                {
                                    try
                                    {
                                        testdataobj.GetTestData(_deleteFilesPath, "");

                                        DataTable dtDelete = testdataobj.Data;
                                        for (int row = 0; i < dtDelete.Rows.Count; row++)
                                        {
                                            DeleteFile((string)dtDelete.Rows[row][0], logg, _epath);
                                        }
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();

                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  DeleteFiles " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                    }
                                    break;
                                }
                            #endregion
                            #region searchgridrow
                            case "searchgridrow":
                                {
                                    try
                                    {
                                        SearchGridRow(arg1, arg2, arg3, uiautomation._processId);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();

                                    }
                                    catch (Exception ex)
                                    {


                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  searchgridrow " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                    }
                                    break;
                                }
                            #endregion
                            #region addrowtodatagrid
                            case "addrowtodatagrid":
                                {
                                    try
                                    {
                                        AddRowToDataGrid(arg1, arg2, uiautomation._processId);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed"; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();

                                    }
                                    catch (Exception ex)
                                    {


                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  addrowtodatagrid " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message;
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                    }
                                    break;
                                }
                            #endregion
                            #region clickmenus
                            case "clickmenus":
                                {
                                    try
                                    {
                                        ClickMenus(arg1, arg2, uiautomation._processId);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed";
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();

                                    }
                                    catch (Exception ex)
                                    {


                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  clickmenus " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message;
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                    }
                                    break;
                                }
                            #endregion
                            #region clickmenus
                            case "clickbeamdesign":
                                {
                                    try
                                    {
                                        ClickBeamDesign(arg1, arg2, arg3, uiautomation._processId);
                                        stopwatch3.Stop();
                                        dr["RESULT"] = "Success";
                                        dr["MESSAGE"] = "Completed";
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();


                                    }
                                    catch (Exception ex)
                                    {

                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  clickbeamdesign " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message;
                                        dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                    }
                                    break;
                                }
                            #endregion
                            #region cleanupcstore
                            case "cleanupcstore":
                                {
                                    try
                                    {
                                        cleanupcstore(arg1);

                                        dr["RESULT"] = "Success";
                                        dr["Message"] = "Completed";
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in  cleanupcstore " + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();

                                    }
                                    break;
                                }

                            #endregion
                            #region menuiteam
                            case "menuiteam":
                                {
                                    try
                                    {
                                        menuiteam(arg1);

                                        dr["RESULT"] = "Sucess";
                                        dr["Message"] = "Completed";
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in menuiteam" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                    }
                                    break;
                                }
                            #endregion
                            #region standardpump
                            case "standardpump":
                                {
                                    try
                                    {
                                        standardpump(arg1, arg2, arg3);

                                        dr["RESULT"] = "Sucess";
                                        dr["Message"] = "Completed";
                                    }
                                    catch (Exception ex)
                                    {
                                        returndriveFromExcel = false;
                                        logg.CreateCustomLog(_epath, "Error in standardpumpsave" + ex.Message.ToString());
                                        dr["RESULT"] = "Failed";
                                        dr["MESSAGE"] = "Terminated" + ex.Message; dr["TIMETAKEN"] = (string)stopwatch3.Elapsed.Seconds.ToString();
                                    }
                                    break;
                                }
                            #endregion
                            #endregion

                            default:
                                returndriveFromExcel = false;
                                throw new Exception("Not a valid Keyword");
                        }
                        dtResultSummary.Rows.Add(dr);
                        sNo = sNo + 1;
                    }
                    else
                    {
                        // logg.CreateCustomLog(_epath, "[Wraper]:Not Executing the Keyword: -> " + keyWord);
                    }

                    if (returndriveFromExcel == false)
                    {
                        logg.CreateCustomLog(_epath, "*********************ScriptTermination ********************************");
                        logg.CreateCustomLog(_epath, "Script: " + excelfilePath + "was terminited due to above errors and host or application was also terminated");
                        break;
                    }

                }
                GenerateReport(_resultsSummaryFile, dtResultSummary);

                if (File.Exists(_resultsFile) == false)
                {
                    Console.WriteLine("Result File does not exist:  " + _resultsFile);
                }
                else
                {
                    DataTable dtResults = GetResultsData(_resultsFile, "");
                    DataRow[] success = dtResults.Select("Result= 'Pass'");
                    DataRow[] failed = dtResults.Select("Result='Fail'");


                    Console.WriteLine("Finished Test (Errors:" + failed.Length.ToString() + "," + " Warnings:" + failcount + ")");
                }
                return returndriveFromExcel;
            }
            catch (Exception ex)

            {
                Console.WriteLine("Generic Error has occurred " + ex.Message);
                return false;
            }

        }
                        #endregion
        #endregion
        #region GetResultsData
        public static DataTable GetResultsData(string testDataFile, string testCase)
        {


            Helper.TestDataManagement testData = new Helper.TestDataManagement();
            Console.WriteLine("Trying to get testat from" + testDataFile);
            testData.GetTestData(testDataFile, "");
            DataTable dt = testData.Data;
            //Console.WriteLine(dt.Rows.Count.ToString());
            return dt;
        }
        #endregion

        #region worddocverification
        private static void worddocverification(string filename, string title)
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            if (File.Exists(filename))
            {
                Console.WriteLine("File exists");
                p.StartInfo.FileName = "WINWORD.EXE";
                p.StartInfo.Arguments = filename;
                p.Start();
                while (!p.MainWindowTitle.Contains(title))
                {
                    Console.WriteLine("Word Title---->" + p.MainWindowTitle.ToString());
                    Thread.Sleep(10);
                    p.Refresh();

                }
                Console.WriteLine("Word title: " + p.MainWindowTitle.ToString());
                p.Kill();

                FileInfo f = new FileInfo(filename);
                long s1 = f.Length;
                Console.WriteLine("File Size: " + s1);
            }
            else
                Console.WriteLine("File does not exist");

        }
        #endregion

        #region specialApply
        private static void specialApply(string pWindow, string childWindow, string btnname)
        {
            WellFloUI.MSUIAutomation wellflocomui1 = new WellFloUI.MSUIAutomation();
            wellflocomui1.GetAppWindow(pWindow);
            wellflocomui1.GetChildWindow(childWindow);
            for (var i = 0; i < 4; i++)
            {
                wellflocomui1.GetChildPane(1);
            }
            wellflocomui1.GetChildPane(2);
            wellflocomui1.ClickChildButton(btnname);
        }
        #endregion

        #region fn_BalloonToolTip
        public static void fn_BalloonToolTip(string str_tooltiptitle, string str_tooltiptext)
        {
            Helper.LogManagement logg1 = new Helper.LogManagement();
            string icon_Path = ConfigurationManager.AppSettings["iconfile"];
            string _epath = ConfigurationManager.AppSettings["logfile"];
            System.Drawing.Icon oIcon = new System.Drawing.Icon(icon_Path);
            NotifyIcon oNtfy = new NotifyIcon();
            oNtfy.Icon = oIcon;
            oNtfy.Visible = true;
            logg1.CreateCustomLog(_epath, " in function ballontooltip");
            oNtfy.Text = "UIAutomationInfo";
            oNtfy.BalloonTipTitle = str_tooltiptitle;
            oNtfy.BalloonTipText = str_tooltiptext;
            oNtfy.ShowBalloonTip(10000);


        }
        #endregion

        #region verify_plotdata
        static void verify_plotdata(string expectedFile, string testcaseid)
        {

            Helper.TestDataManagement testlocal = new Helper.TestDataManagement();
            testlocal.GetTestData(expectedFile, testcaseid);
            string _logicalName = (string)testlocal.Structure.Rows[0]["FieldName"].ToString();
            string _controlValue = (string)testlocal.Data.Rows[0][_logicalName].ToString();
            char[] celldellim = new char[] { ';' };
            string[] arr = _controlValue.Split(celldellim);
            string _testDataPath = ConfigurationManager.AppSettings["testinputdata"];  //arr[0];
            string paramfiletcase = arr[1];
            string expectedfileName = arr[2];
            string tcase = arr[3];
            string otptfile = arr[4];

            testlocal.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcase, "expectedFile", _testDataPath + expectedfileName);
            testlocal.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcase, "testcaseID", tcase);
            testlocal.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcase, "tempFilePath", @"C:\created.txt");
            testlocal.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcase, "resultspath", otptfile);

            //put information of testdata dynamicaly in a vbs file to read 

            RunVBS(_testDataPath + @"\Verifyplotdata.vbs");
            System.IO.File.Delete(@"C:\created.txt");
        }
        #endregion

        #region selectspecfedrows
        static void selectspecfedrows(int rowfrom, int rowto)
        {

            System.Windows.Forms.SendKeys.Flush();
            System.Windows.Forms.SendKeys.SendWait("^{Home}");

            System.Windows.Forms.SendKeys.Flush();
            System.Windows.Forms.SendKeys.SendWait("{Down}");

            System.Windows.Forms.SendKeys.Flush();
            System.Windows.Forms.SendKeys.SendWait("{up}");

            if (rowfrom == rowto)
            {
                for (int ir = 0; ir < rowfrom; ir++)
                {
                    System.Windows.Forms.SendKeys.SendWait("{Down}");
                }
            }
            else
            {
                //for (int ir = 0; ir < rowfrom; ir++)
                //{
                //    System.Windows.Forms.SendKeys.SendWait("{Down}");
                //}

                for (int ir = 1; ir < (rowto - rowfrom); ir++)
                {
                    System.Windows.Forms.SendKeys.SendWait("+{Down}");
                }

            }

        }
        #endregion

        #region RunVBS
        static void RunVBS(string vbsFilepath)
        {
            var proc = System.Diagnostics.Process.Start(vbsFilepath);
            proc.WaitForExit();
        }
        #endregion

        #region GetOSArchitecture
        private static string GetOSArchitecture()
        {
            SelectQuery query = new SelectQuery(@"Select * from Win32_Processor");
            string osbit = "";
            //initialize the searcher with the query it is supposed to execute
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query))
            {
                //execute the query
                foreach (ManagementObject process in searcher.Get())
                {

                    //print process properties
                    Console.WriteLine("/*********Processor Information ***************/");
                    Console.WriteLine("{0}{1}", "Addres Bit 32Bt Or 64 bit :", process["AddressWidth"]);
                    osbit = process["AddressWidth"].ToString();
                }

            }
            return osbit;
        }
        #endregion

        /// <summary>
        /// This function is to update the cell in structure sheet with value passed
        /// </summary>
        /// <param name="expectedDataFile">path of Excel file which needs to be updated</param>
        /// <param name="columnName">Name of the column that needs to be updated </param>
        /// <param name="data">Value for the columnn</param>
        /// <param name="condcolumnName">column name to be used for where condition</param>
        /// <param name="condcolumnValue">value of cell for condcolumnName</param>
        /// 

        #region UpdateStructure
        private static void UpdateStructure(string expectedDataFile, string columnName, string data, string condcolumnName, string condcolumnValue)
        {
            string _excelConnectionPrefix = @"Driver={Microsoft Excel Driver (*.xls)};DriverId=790;ReadOnly=0;Dbq=";
            string _conString = "";

            try
            {
                string testDataPath = ConfigurationManager.AppSettings["testinputdata"];
                if (!testDataPath.EndsWith(@"\"))
                    testDataPath = testDataPath + @"\";
                Console.WriteLine(testDataPath);
                Console.WriteLine("Inside GetExcel Connection");

                _conString = _excelConnectionPrefix + testDataPath + expectedDataFile;

                var _con = new OdbcConnection(_conString);
                _con.Open();
                Console.WriteLine(_con.State.ToString());

                var _command = new OdbcCommand();


                _command.Connection = _con;
                string space = " ";
                if (columnName.Contains(space))
                {
                    columnName = "[" + columnName + "]";
                }

                _command.CommandText = "Update [Structure$] set " + columnName + "='" + data + "' where " + condcolumnName + "='" + condcolumnValue + "'";


                _command.ExecuteNonQuery();

                _con.Close();
                _con.Dispose();
            }
            catch
            {
                throw;
            }

        }
        #endregion

        #region TerminateProcessByForce
        private static void TerminateProcessByForce(string strprocess)
        {
            try
            {
                //Assign the name of the process you want to kill on the remote machine
                string processName = strprocess;

                //Assign the user name and password of the account to ConnectionOptions object
                //which have administrative privilege on the remote machine.
                ConnectionOptions connectoptions = new ConnectionOptions();
                //   connectoptions.Username = @"YourDomainName\UserName";
                //  connectoptions.Password = "User Password";

                //IP Address of the remote machine
                string ipAddress = "127.0.0.1";
                ManagementScope scope = new ManagementScope(@"\\" + ipAddress + @"\root\cimv2", connectoptions);

                //Define the WMI query to be executed on the remote machine
                SelectQuery query = new SelectQuery("select * from Win32_process where name = '" + processName + "'");

                using (ManagementObjectSearcher searcher = new
                            ManagementObjectSearcher(scope, query))
                {
                    foreach (ManagementObject process in searcher.Get())
                    {

                        process.InvokeMethod("Terminate", null);

                    }
                }

            }
            catch (Exception ex)
            {
                //Log exception in exception log.
                //Logger.WriteEntry(ex.StackTrace);
                Console.WriteLine(ex.StackTrace);

            }
        }
        #endregion

        #region writeGridContent
        static void writeGridContent(string testDataFile, string testcase)
        {
            string TableType = "";
            string SearchBy = "";
            string SearchValue = "";
            int tindex = -1;
            var _logicalName = "";//coulmnname/fieldname  alias 
            Helper.TestDataManagement testData = new Helper.TestDataManagement();
            UIAutomation_App.UIAutomationAction uiautomation1 = new UIAutomation_App.UIAutomationAction();

            string _epath = ConfigurationManager.AppSettings["logfile"];
            uiautomation1._eLogPtah = _epath;


            try
            {

                testData.GetVerificationData(testDataFile, testcase);
                TableType = testData.Template.Rows[0]["TableType"].ToString();
                SearchBy = testData.Template.Rows[0]["SearchBy"].ToString().ToLower();
                SearchValue = testData.Template.Rows[0]["SearchValue"].ToString();
                if (SearchBy.ToLower() == "index")
                {
                    tindex = Int32.Parse(SearchValue);
                }

                uiautomation1.uiAutomationWindow = uiautomation1.GetUIAutomationWindow("title", "K2");

                uiautomation1.uiAutomationCurrentParent = uiautomation1.uiAutomationWindow;

                AutomationElement datagridtable = uiautomation1.GetUIAutomationDataGrid(SearchBy, SearchValue, tindex);
                if (datagridtable != null)
                {
                    uiautomation1.logTofile(uiautomation1._eLogPtah, "[writeGridContent]:Datagrid Object has been detected!");
                }
                AutomationElementCollection datarows = datagridtable.FindAll(TreeScope.Children,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem));
                uiautomation1.logTofile(uiautomation1._eLogPtah, "[writeGridContent]: DataGrid Rows count is =" + datarows.Count.ToString());
                for (int ip = 0; ip < datarows.Count; ip++)
                {
                    AutomationElementCollection cells = datarows[ip].FindAll(TreeScope.Children,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom));
                    uiautomation1.logTofile(uiautomation1._eLogPtah, "[writeGridContent]: Cells or columns count " + cells.Count);
                    try
                    {
                        AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(cells[0]);
                        uiautomation1.logTofile(uiautomation1._eLogPtah, "after element node");
                        ValuePattern valpat = (ValuePattern)elementNode.GetCurrentPattern(ValuePattern.Pattern);
                        _logicalName = (string)testData.Template.Rows[0]["FieldName"];
                        var valueofcell = valpat.Current.Value;
                        valpat.SetValue((string)testData.ExpectedData.Rows[ip][_logicalName]);
                    }
                    catch
                    {
                    }
                    for (int i = 0; i < cells.Count; i++)
                    {
                        uiautomation1.logTofile(uiautomation1._eLogPtah, "[writeGridContent]: Trying Value patern " + i);
                        uiautomation1.logTofile(uiautomation1._eLogPtah, "before element node");
                        AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(cells[i]);
                        uiautomation1.logTofile(uiautomation1._eLogPtah, "after element node");
                        ValuePattern valpat = (ValuePattern)elementNode.GetCurrentPattern(ValuePattern.Pattern);
                        _logicalName = (string)testData.Template.Rows[i]["FieldName"];
                        var valueofcell = valpat.Current.Value;
                        valpat.SetValue((string)testData.ExpectedData.Rows[ip][_logicalName]);
                    }

                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[writeGridContent]:Execption Message : " + ex.Message.ToString());
            }
        }
        #endregion

        #region GenerateReport
        static void GenerateReport(string reportFile, DataTable logTable)
        {



            using (StreamWriter writer = new StreamWriter(reportFile, true))
            {
                if (writer.BaseStream.Length == 0)
                {
                    foreach (DataColumn column in logTable.Columns)
                    {


                        writer.Write('\u0022' + column.ColumnName + '\u0022' + ",");

                    }
                    writer.WriteLine();
                }
                for (int i = 0; i < logTable.Rows.Count; i++)
                {


                    foreach (DataColumn column in logTable.Columns)
                    {
                        writer.Write('\u0022' + (string)logTable.Rows[i][column.ColumnName] + '\u0022' + ",");

                    }
                    writer.WriteLine();
                }

            }




        }
        #endregion

        #region CompareData
        static DataTable CompareData(DataTable ExpectedData, DataTable ActualData)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("FieldName");
                dt.Columns.Add("ExpectedValue");
                dt.Columns.Add("ActualValue");
                dt.Columns.Add("Result");
                dt.Columns.Add("DateTime");
                for (int j = 0; j < ActualData.Rows.Count; j++)
                {
                    DataRow dr = dt.NewRow();
                    string actualValue = "";
                    actualValue = ActualData.Rows[j]["COLUMN2"].ToString();
                    actualValue = actualValue.Replace("\t\r\a", "");
                    string expectedValue = "NoData";
                    expectedValue = (string)ExpectedData.Rows[j]["COLUMN2"].ToString();
                    if (Convert.IsDBNull(ExpectedData.Rows[j]["COLUMN1"]) == false)
                    {
                        dr["FieldName"] = (string)ExpectedData.Rows[j]["COLUMN1"].ToString();
                    }
                    else
                    {
                        dr["FieldName"] = "";
                    }
                    dr["ExpectedValue"] = expectedValue;
                    dr["ActualValue"] = actualValue;
                    dr["DateTime"] = DateTime.Now.ToLocalTime().ToString();
                    string result = "Fail";
                    if (actualValue.ToLower().Trim() == expectedValue.ToLower().Trim())
                    {
                        result = "Pass";
                    }
                    dr["Result"] = result;
                    dt.Rows.Add(dr);
                }
                return dt;
            }

            catch (Exception ex)
            {
                //todo add logging comments
                throw new Exception(" Error in function CompareData: " + System.Environment.NewLine + ex.Message);
            }
        }
        #endregion

        #region verifyFileExistence
        static DataTable verifyFileExistence(string expectedData)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("FieldName");
            dt.Columns.Add("ExpectedValue");
            dt.Columns.Add("ActualValue");
            dt.Columns.Add("Result");
            dt.Columns.Add("DateTime");
            DataRow dr = dt.NewRow();
            dr["FieldName"] = expectedData;
            dr["ExpectedValue"] = "Exists";
            if (File.Exists(expectedData) == false)
            {
                dr["ActualValue"] = "NonExists";
                dr["Result"] = "Fail";
            }
            else
            {
                dr["ActualValue"] = "Exists";
                dr["Result"] = "Pass";
            }
            dr["DateTime"] = DateTime.Now.ToLocalTime().ToString();
            dt.Rows.Add(dr);
            return dt;
        }
        #endregion

        #region createOutputImage
        private static void createOutputImage(string output)
        {
            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);

            Graphics graphics = Graphics.FromImage(printscreen as System.Drawing.Image);

            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);

            printscreen.Save(output, ImageFormat.Png);

        }
        #endregion

        #region Create Execution Report
        private static DataTable createExecutionReport(DataTable output)
        {
            DataTable edt = new DataTable();
            edt.Columns.Add("StructureSheetName");
            edt.Columns.Add("Function");
            edt.Columns.Add("TestCase");
            edt.Columns.Add("Action");
            edt.Columns.Add("Parent");
            edt.Columns.Add("ControlType");
            edt.Columns.Add("ControlName");
            edt.Columns.Add("ControlStatus");
            edt.Columns.Add("ControlActionStatus");
            for (int i = 0; i < output.Rows.Count; i++)
            {
                if (output.Rows[i]["Action"].ToString() == "_")
                {
                    if (output.Rows[i]["ControlType"].ToString() == "_")
                    {
                    }
                    else
                    {
                        DataRow dr = edt.NewRow();
                        dr["StructureSheetName"] = output.Rows[i]["StructureSheetName"].ToString();
                        dr["Function"] = output.Rows[i]["FunctionName"].ToString();
                        dr["TestCase"] = output.Rows[i]["TestCaseID"].ToString();
                        dr["Parent"] = output.Rows[i]["ParentSearchValue"].ToString();
                        dr["ControlType"] = output.Rows[i]["ControlType"].ToString();
                        dr["ControlName"] = output.Rows[i]["ControlName"].ToString();
                        dr["ControlStatus"] = output.Rows[i]["Control Detected"].ToString();
                        dr["ControlActionStatus"] = output.Rows[i]["Action Performed on Control"].ToString();
                        edt.Rows.Add(dr);
                    }
                }
                else
                {
                    DataRow dr = edt.NewRow();
                    dr["StructureSheetName"] = output.Rows[i]["StructureSheetName"].ToString();
                    dr["Function"] = output.Rows[i]["FunctionName"].ToString();
                    dr["TestCase"] = output.Rows[i]["TestCaseID"].ToString();
                    string action = output.Rows[i]["Action"].ToString();
                    if (action.ToLower().Trim() == "wait")
                    {
                        string wait = output.Rows[i]["ControlName"].ToString();
                        dr["Action"] = "Waiting for " + wait + " seconds";
                    }
                    else if (action.ToLower().Trim() == "clearwindow")
                    {
                        dr["Action"] = "Clearing parent window";
                    }
                    else if (action.ToLower().Trim() == "keyboard")
                    {
                        string keys = output.Rows[i]["ControlName"].ToString();
                        dr["Action"] = "Performing key board action for " + keys;
                    }
                    edt.Rows.Add(dr);

                }
            }
            return edt;

        }
        #endregion
        #region Create Execution HTML Report
        private static string createExecutionHtmlReport()
        {
            string path = _executionPath;
            var fileNames = Directory.GetFiles(path);
            string scriptCount = fileNames.Count().ToString();
            int Pass = 0;
            int fail = 0;
            for (int j = 0; j < fileNames.Count(); j++)
            {
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileNames[j]);
                if (fileNameWithoutExtension.Contains("Pass"))
                {
                    Pass = Pass + 1;
                }
                else
                {
                    fail = fail + 1;
                }
            }
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html >");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("<title >");
            strHTMLBuilder.Append("</title>");
            strHTMLBuilder.Append("<meta name=" + "\"viewport\"" + "content=" + "\"width=device-width, initial-scale=1\"" + ">");
            strHTMLBuilder.Append("<script  src=" + @"""http://code.jquery.com/jquery-1.11.1.min.js""" + ">");
            strHTMLBuilder.Append("</script>");
            strHTMLBuilder.Append("<script  src=" + @"""http://code.jquery.com/mobile/1.4.5/jquery.mobile-1.4.5.min.js""" + ">");
            strHTMLBuilder.Append("</script>");
            strHTMLBuilder.Append("<link rel=" + "\"stylesheet\"" + "href=" + @"""http://code.jquery.com/mobile/1.4.5/jquery.mobile-1.4.5.min.css""" + ">");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body>");

            strHTMLBuilder.Append("<div data-role=" + @"""page""" + ">");
            strHTMLBuilder.Append("<div data-role=" + @"""main""" + "class=" + @"""ui-content""" + ">");
            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='lightyellow' style='font-family:Garamond; font-size:medium'>");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("Count");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr >");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("Scripts");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append(scriptCount);
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr >");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("Completed");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append(Pass.ToString());
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr >");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("Terminated");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append(fail.ToString());
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr >");
            strHTMLBuilder.Append("</table>");
            for (int i = 0; i < fileNames.Count(); i++)
            {
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileNames[i]);
                DataTable edt = GetResultsData(fileNames[i], "");
                DataTable dt = createExecutionReport(edt);
                dt.TableName = fileNameWithoutExtension;
                strHTMLBuilder.Append("<div data-role=" + "\"collapsibleset\"" + ">");
                strHTMLBuilder.Append("<div data-role=" + "\"collapsible\"" + ">");
                if (fileNameWithoutExtension.Contains("Pass"))
                {
                    strHTMLBuilder.Append("<h1 style=" + "\"color:green\"" + ">");
                    strHTMLBuilder.Append("Script-");
                    strHTMLBuilder.Append(fileNames[i].Replace(".csv", ".xls") + "--------- Completed");
                }
                else
                {
                    strHTMLBuilder.Append("<h1 style=" + "\"color:red\"" + ">");
                    strHTMLBuilder.Append("Script-");
                    strHTMLBuilder.Append(fileNames[i].Replace(".csv", ".xls") + "--------- Terminated");
                }
                strHTMLBuilder.Append("</h1>");
                strHTMLBuilder.Append("<p>");
                strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='lightyellow' style='font-family:Garamond; font-size:smaller'>");

                strHTMLBuilder.Append("<tr >");
                foreach (DataColumn myColumn in dt.Columns)
                {
                    strHTMLBuilder.Append("<td >");
                    strHTMLBuilder.Append(myColumn.ColumnName);
                    strHTMLBuilder.Append("</td>");

                }
                strHTMLBuilder.Append("</tr>");


                foreach (DataRow myRow in dt.Rows)
                {

                    strHTMLBuilder.Append("<tr >");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        strHTMLBuilder.Append("<td >");
                        strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                        strHTMLBuilder.Append("</td>");

                    }
                    strHTMLBuilder.Append("</tr>");
                }

                //Close tags.  
                strHTMLBuilder.Append("</table>");
                strHTMLBuilder.Append("</p>");
                strHTMLBuilder.Append("</div>");
                strHTMLBuilder.Append("</div>");
            }
            strHTMLBuilder.Append("</div>");
            strHTMLBuilder.Append("</div>");
            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");

            string Htmltext = strHTMLBuilder.ToString();

            return Htmltext;

        }
        #endregion
        #region Send Mail
        //protected static void SendMail()
        //{
        //    Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
        //    Microsoft.Office.Interop.Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
        //    mailItem.Subject = "This is the subject";
        //    mailItem.To = "ashok.krishna@me.weatherford.com";
        //    mailItem.Body = "This is the message.";
        //    mailItem.Attachments.Add(_htmlReportsPath + @"\ finalReport.html");
        //    mailItem.Attachments.Add(_executionPath + "ExecutionReport.html");
        //    mailItem.Display(false);
        //    mailItem.Send();
        //    app.Quit();
        //}



        #endregion

        #region dbtohtml
        protected static string ExportDatatableToHtml()
        {
            var fileNames = Directory.GetFiles(_htmlReportsPath);
            string scriptCount = fileNames.Count().ToString();
            int Pass = 0;
            int fail = 0;
            for (int j = 0; j < fileNames.Count(); j++)
            {
                DataTable dt = GetResultsData(fileNames[j], "");
                DataRow[] success = dt.Select("Result= 'Pass'");
                DataRow[] failed = dt.Select("Result='Fail'");
                if (failed.Length == 0)
                {
                    Pass = Pass + 1;
                }
                else
                {
                    fail = fail + 1;
                }
            }
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html >");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("<title >");
            strHTMLBuilder.Append("</title>");
            strHTMLBuilder.Append("<meta name=" + "\"viewport\"" + "content=" + "\"width=device-width, initial-scale=1\"" + ">");
            strHTMLBuilder.Append("<script  src=" + @"""http://code.jquery.com/jquery-1.11.1.min.js""" + ">");
            strHTMLBuilder.Append("</script>");
            strHTMLBuilder.Append("<script  src=" + @"""http://code.jquery.com/mobile/1.4.5/jquery.mobile-1.4.5.min.js""" + ">");
            strHTMLBuilder.Append("</script>");
            strHTMLBuilder.Append("<link rel=" + "\"stylesheet\"" + "href=" + @"""http://code.jquery.com/mobile/1.4.5/jquery.mobile-1.4.5.min.css""" + ">");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body>");
            strHTMLBuilder.Append("<div data-role=" + @"""page""" + ">");
            strHTMLBuilder.Append("<div data-role=" + @"""main""" + "class=" + @"""ui-content""" + ">");
            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='lightyellow' style='font-family:Garamond; font-size:medium'>");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("Count");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr >");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("Scripts");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append(scriptCount);
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr >");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("Pass");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append(Pass.ToString());
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr >");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append("Fail");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("<td >");
            strHTMLBuilder.Append(fail.ToString());
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr >");
            strHTMLBuilder.Append("</table>");
            for (int i = 0; i < fileNames.Count(); i++)
            {
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileNames[i]);
                DataTable dt = GetResultsData(fileNames[i], "");
                DataRow[] success = dt.Select("Result= 'Pass'");
                DataRow[] failed = dt.Select("Result='Fail'");
                dt.TableName = fileNameWithoutExtension;
                strHTMLBuilder.Append("<div data-role=" + "\"collapsibleset\"" + ">");
                strHTMLBuilder.Append("<div data-role=" + "\"collapsible\"" + ">");
                if (failed.Length == 0)
                {
                    strHTMLBuilder.Append("<h1 style=" + "\"color:green\"" + ">");
                    strHTMLBuilder.Append("Script-");
                    strHTMLBuilder.Append( ReturnScriptName( fileNames[i].Replace(".csv", ".xls")) + "--------- PASS");
                }
                else
                {
                    strHTMLBuilder.Append("<h1 style=" + "\"color:red\"" + ">");
                    strHTMLBuilder.Append("Script-");
                    strHTMLBuilder.Append(ReturnScriptName(fileNames[i].Replace(".csv", ".xls")) + "--------- FAIL");
                }
                strHTMLBuilder.Append("</h1>");
                strHTMLBuilder.Append("<p>");
                strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='lightyellow' style='font-family:Garamond; font-size:smaller'>");

                strHTMLBuilder.Append("<tr >");
                foreach (DataColumn myColumn in dt.Columns)
                {
                    strHTMLBuilder.Append("<td >");
                    strHTMLBuilder.Append(myColumn.ColumnName);
                    strHTMLBuilder.Append("</td>");

                }
                strHTMLBuilder.Append("</tr>");


                foreach (DataRow myRow in dt.Rows)
                {

                    strHTMLBuilder.Append("<tr >");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        strHTMLBuilder.Append("<td >");
                        strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                        strHTMLBuilder.Append("</td>");

                    }
                    strHTMLBuilder.Append("</tr>");
                }

                //Close tags.  
                strHTMLBuilder.Append("</table>");
                strHTMLBuilder.Append("</p>");
                strHTMLBuilder.Append("</div>");
                strHTMLBuilder.Append("</div>");
            }
            strHTMLBuilder.Append("</div>");
            strHTMLBuilder.Append("</div>");
            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");

            string Htmltext = strHTMLBuilder.ToString();

            return Htmltext;

        }



        #endregion
        #region compareXML
        private static void CompareXML(string expectedXmlFilePath, string actualXmlPath, string reportTemplate, string reportPath, LogManagement logg, string logpath)
        {
            try
            {
                logg.CreateCustomLog(logpath, "Inside CompareXML");
                TestDataManagement management = new TestDataManagement();
                ReportsManagement management2 = new ReportsManagement();
                management.ActualData.Clear();
                management.ExpectedData.Clear();
                logg.CreateCustomLog(logpath, "Cleared old data");
                DataSet set = new DataSet();
                
                DataSet set2 = new DataSet();
                set.ReadXml(expectedXmlFilePath);
                logg.CreateCustomLog(logpath, "Loaded ExpectedXML");
                set2.ReadXml(actualXmlPath);
                logg.CreateCustomLog(logpath, "Loaded actualXML");
                management.ExpectedData = set.Tables[0];
                management.ActualData = set2.Tables[0];
                logg.CreateCustomLog(logpath, "Created Actual and Expected Data Sets");
                management.CompareData();
                logg.CreateCustomLog(logpath, "DataBinder Compared");
                management2.ResultTable = management.ResultTable;
                Console.WriteLine(management.ResultTable.Rows.Count);
                management2.ReportPath = reportPath;
                management2.GenerateReport(reportTemplate);
                logg.CreateCustomLog(logpath, "Report Generated");
            }
            catch (Exception exception)
            {
                throw new Exception("Error in function CompareXML" + exception.Message);
            }
        }

        #endregion
        #region compareExcel
        private static void CompareExcel(string expectedExcelFilePath, string actualExcelPath)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("FieldName");
                dt.Columns.Add("ExpectedValue");
                dt.Columns.Add("ActualValue");
                dt.Columns.Add("Result");
                dt.Columns.Add("DateTime");
                DataRow drMisMatch = dt.NewRow();
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook expectedWorkBook = excelApp.Workbooks.Open(expectedExcelFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Sheets expectedWorkSheet = expectedWorkBook.Worksheets;
                Excel.Workbook actualWorkBook = excelApp.Workbooks.Open(actualExcelPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Sheets actualWorkSheet = expectedWorkBook.Worksheets;
                drMisMatch["FieldName"] = "ExcelCOmparision";
                drMisMatch["ExpectedValue"] = expectedExcelFilePath;
                drMisMatch["ActualValue"] = actualExcelPath;
                drMisMatch["DateTime"] = DateTime.Now.ToLocalTime().ToString();
                if (expectedWorkSheet.Count == actualWorkSheet.Count)
                {
                    for (int i = 1; i <= expectedWorkSheet.Count; i++)
                    {
                        Excel.Worksheet expected = (Excel.Worksheet)expectedWorkSheet[i];
                        Excel.Worksheet actual = (Excel.Worksheet)actualWorkSheet[i];
                        Excel.Range expectedRange = expected.UsedRange;
                        Excel.Range actualRange = actual.UsedRange;
                        if (expectedRange.Rows.Count == actualRange.Rows.Count)
                        {
                            if (expectedRange.Columns.Count == actualRange.Columns.Count)
                            {
                                for (int j = 1; j <= expectedRange.Rows.Count; j++)
                                {
                                    for (int k = 1; k < expectedRange.Columns.Count; k++)
                                    {
                                        if (expected.Cells[j, k] != actual.Cells[j, k])
                                        {
                                            drMisMatch["Result"] = "Fail";
                                            break;
                                        }
                                        else
                                        {

                                        }

                                    }
                                }
                            }
                            else
                            {
                                drMisMatch["Result"] = "Fail";
                                break;
                            }
                        }
                        else
                        {
                            drMisMatch["Result"] = "Fail";
                            break;
                        }

                    }

                }
                else
                {
                    drMisMatch["Result"] = "Fail";
                }
                excelApp.Quit();

            }
            catch (Exception exception)
            {
                throw new Exception("Error in function CompareXML" + exception.Message);
            }
        }
        #endregion


        #endregion
        #region Lowis related functions
        #region cleanupcstore
        public static void cleanupcstore(string cleanup_path)
        {
            if (Directory.Exists(cleanup_path)) // if folder exists
            {
                Directory.Delete(cleanup_path, true); //recursive delete (all subdirs, files)
            }
        }
        #endregion

        #region menuiteam
        public static void menuiteam(string menuiteam_name)
        {
            AutomationElement ae = AutomationElement.RootElement;
            AutomationElement mn1 = ae.FindFirst(TreeScope.Descendants,
                  new System.Windows.Automation.AndCondition(
                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem),
                 new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, menuiteam_name)));

            ClickControl(mn1);
        }
        #endregion

        #region standardpump
        public static void standardpump(string arg1, string arg2, string arg3)
        {
            AutomationElement ae = AutomationElement.RootElement;
            AutomationElement appwindow1 = ae.FindFirst(TreeScope.Descendants,
                new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                new PropertyCondition(AutomationElement.NameProperty, arg1)));
            appwindow1.SetFocus();
            WindowPattern wn1 = (WindowPattern)appwindow1.GetCurrentPattern(WindowPattern.Pattern);
            wn1.SetWindowVisualState(WindowVisualState.Maximized);

            AutomationElement mn2 = appwindow1.FindFirst(TreeScope.Children,
                 new AndCondition(
                new PropertyCondition(AutomationElement.ClassNameProperty, arg2),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)));

            AutomationElement bn = mn2.FindFirst(TreeScope.Children,
                new AndCondition(
               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
               new PropertyCondition(AutomationElement.HelpTextProperty, arg3)));

            ClickControl(bn);
        }
        #endregion

        #region DeleteFile
        private static void DeleteFile(string fileName, Helper.LogManagement logg, string logpath)
        {
            try
            {

                logg.CreateCustomLog(logpath, "Inside DeleteFile");
                if (System.IO.File.Exists(fileName))
                {
                    System.IO.File.Delete(fileName);
                    logg.CreateCustomLog(logpath, "File found deleted");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error in function DeleteFile" + ex.Message);
            }

        }
        #endregion

        #region ReturnScriptName
        private static string ReturnScriptName(string inputFilePath)
        {
            string ReturnScriptName = null;
            int lastbackslashpos = inputFilePath.LastIndexOf("\\") + 1;
            System.Console.WriteLine("lastbackslashpos: " + lastbackslashpos);
            int arg1Len = inputFilePath.Length;
            System.Console.WriteLine("arg1 Length: " + arg1Len);
            int diff = inputFilePath.Length - lastbackslashpos;
            System.Console.WriteLine("diff Length: " + diff);
            ReturnScriptName = inputFilePath.Substring(lastbackslashpos, diff);
            if (ReturnScriptName.Contains(".xls") == true)
            {
                ReturnScriptName = ReturnScriptName.Replace(".xls", "");
            }
            else if (ReturnScriptName.Contains(".exe") == true)
            {
                ReturnScriptName = ReturnScriptName.Replace(".exe", "");
            }



            return ReturnScriptName;
        }
        #endregion

        #region SearchGridRow
        private static void SearchGridRow(string windowTitle, string gridName, string dataItemCustomvalue, int processID)
        {


            UIAutomation_App.UIAutomationAction uAutomation = new UIAutomation_App.UIAutomationAction();
            uAutomation._processId = processID;
            Console.WriteLine("ProcessId:" + uAutomation._processId);

            try
            {

                System.Threading.Thread.Sleep(1000);

                uAutomation.uiAutomationCurrentParent = null;
                Console.WriteLine("Set current Parent to null ");
                AutomationElement winElement = GetCustomWindowByName(windowTitle, processID);
                uAutomation.uiAutomationCurrentParent = winElement;

                Console.WriteLine("Window Name:" + winElement.Current.Name);

                AutomationElementCollection gridCollection = winElement.FindAll(TreeScope.Descendants,
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataGrid));



                #region Searching Grid in Collection
                AutomationElement grid = null;
                for (int i = 0; i < gridCollection.Count; i++)
                {
                    if (gridCollection[i].Current.Name.ToLower() == gridName.ToLower() || gridCollection[i].Current.AutomationId.ToLower() == gridName.ToLower())
                    {
                        grid = gridCollection[i];
                        break;
                    }

                }
                #endregion


                Console.WriteLine("Grid found:" + grid.Current.Name + " searching for Grid Rows");

                AutomationElementCollection gridRows = grid.FindAll(TreeScope.Children,
                                         new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem));


                Console.WriteLine("Griid Row count" + gridRows.Count);

                Boolean found = false;
                for (int i = 0; i < gridRows.Count; i++)
                {
                    AutomationElementCollection customTags = gridRows[i].FindAll(TreeScope.Children,
                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom));

                    foreach (AutomationElement item in customTags)
                    {
                        if (item.Current.Name.Trim().ToLower() == dataItemCustomvalue.Trim().ToLower())
                        {
                            //SelectionItemPattern selectPat = (SelectionItemPattern)gridRows[i].GetCurrentPattern(SelectionItemPattern.Pattern);
                            SelectionItemPattern selectPat = (SelectionItemPattern)item.GetCurrentPattern(SelectionItemPattern.Pattern);
                            Console.WriteLine("Row found : " + dataItemCustomvalue);
                            selectPat.Select();

                            found = true;
                            break;
                        }
                    }
                    if (found == true)
                        break;
                }
            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message);
            }

        }
        #endregion

        #region AddRowToDataGrid
        private static void AddRowToDataGrid(string verificationFile, string testCase, int processID)
        {

            try
            {

                Console.WriteLine("Key word AddRowToDataGrid");
                UIAutomation_App.UIAutomationAction uAutomation = new UIAutomation_App.UIAutomationAction();

                Helper.TestDataManagement helper = new TestDataManagement();

                DataTable template = helper.GetVerificationData(verificationFile, testCase);
                DataTable data = helper.ExpectedData;

                if (template.Rows.Count <= 0)
                    throw new Exception("No data found in template sheet");
                if (data.Rows.Count <= 0)
                    throw new Exception("No data found in data sheet");

                Console.WriteLine("Read data from Template and Data Sheet");
                string parentSearchBy = (string)template.Rows[0]["ParentSearchBy"].ToString();
                string parentSearchValue = (string)template.Rows[0]["ParentSearchValue"].ToString();

                string gridSearchBy = (string)template.Rows[0]["gridsearchby"].ToString();
                string gridSearchValue = (string)template.Rows[0]["gridsearchvalue"].ToString();

                int targetRow = 0;


                uAutomation._processId = processID;


                AutomationElement winElement = GetCustomWindowByName(parentSearchValue, processID);
                uAutomation.uiAutomationCurrentParent = winElement;


                Console.WriteLine("Read window : " + winElement.Current.Name);
                AutomationElementCollection gridCollection = winElement.FindAll(TreeScope.Descendants,
                                                                                     new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataGrid));

                #region Searching Grid in Collection
                AutomationElement grid = null;
                for (int i = 0; i < gridCollection.Count; i++)
                {
                    if (gridCollection[i].Current.Name.ToLower() == gridSearchValue.ToLower() || gridCollection[i].Current.AutomationId.ToLower() == gridSearchValue.ToLower())
                    {
                        grid = gridCollection[i];
                        break;
                    }

                }
                #endregion


                Console.WriteLine("Read Grid : " + grid.Current.Name);

                AutomationElementCollection gridRows = grid.FindAll(TreeScope.Children,
                                     new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem));

                Console.WriteLine("Read Grid rows : " + gridRows.Count);

                #region Loop from Template Sheet and for each row in Template seek data for test cases from data sheet
                for (int templateRow = 0; templateRow < template.Rows.Count; templateRow++)
                {
                    int columnNo = int.Parse(template.Rows[templateRow]["columnno"].ToString());
                    string fieldName = (string)template.Rows[templateRow]["FieldName"].ToString();
                    for (int dataRow = 0; dataRow < data.Rows.Count; dataRow++)
                    {

                        targetRow = int.Parse(data.Rows[dataRow]["RowNo"].ToString());
                        AutomationElementCollection customControls = gridRows[targetRow].FindAll(TreeScope.Children,
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom));

                        string datavalue = (string)data.Rows[dataRow][fieldName].ToString();
                        Console.WriteLine("Current Control :" + (string)template.Rows[templateRow]["controltype"].ToString());

                        switch ((string)template.Rows[templateRow]["ControlType"].ToString())
                        {
                            case "ucombobox":
                                AutomationElement comboBox = customControls[columnNo].FindFirst(TreeScope.Descendants,
                                         new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox));
                                comboBox.SetFocus();
                                ValuePattern value = (ValuePattern)comboBox.GetCurrentPattern(ValuePattern.Pattern);
                                value.SetValue(datavalue);
                                break;
                            case "uedit":
                                AutomationElement edit = customControls[columnNo].FindFirst(TreeScope.Descendants,
                                     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                                edit.SetFocus();

                                ValuePattern valPat = (ValuePattern)edit.GetCurrentPattern(ValuePattern.Pattern);
                                valPat.SetValue(datavalue);
                                break;
                            default:
                                break;
                        }


                    }
                }
                #endregion


            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message);
            }



        }
        #endregion

        #region ClickMenus
        private static void ClickMenus(string windowTitle, string menus, int processID)
        {
            try
            {
                UIAutomation_App.UIAutomationAction uAutomation = new UIAutomation_App.UIAutomationAction();

                uAutomation._processId = processID;

                AutomationElement winElement = uAutomation.GetUIAutomationWindow("title", windowTitle);
                uAutomation.uiAutomationCurrentParent = winElement;
                Console.WriteLine(winElement.Current.Name);

                string[] menuNames = menus.Split('|');

                #region search for the start in LOWIS Client

                Console.WriteLine("Searching Start");

                AutomationElementCollection textCollection = winElement.FindAll(TreeScope.Descendants,
                                                                                     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));

                AutomationElement start = null;
                for (int i = 0; i < textCollection.Count; i++)
                {
                    if (textCollection[i].Current.Name.ToLower() == "Start".ToLower() || textCollection[i].Current.AutomationId.ToLower() == "Start".ToLower())
                    {
                        start = textCollection[i];
                        break;
                    }

                }

                ClickControl(start);
                Console.WriteLine("Start Button:" + start.Current.Name);
                #endregion




                if (menuNames.Length != 2)
                {
                    throw new Exception("Need to pass menu names separared by pipe e.g. Menu1|Menu2");
                }

                AutomationElementCollection menuCollection = AutomationElement.RootElement.FindAll(TreeScope.Descendants,
                                                                                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));

                AutomationElement firstLevel = null;
                Console.WriteLine("Menu Count" + menuCollection.Count);

                for (int i = 0; i < menuCollection.Count; i++)
                {
                    Console.WriteLine(menuCollection[i].Current.Name);
                    if (menuCollection[i].Current.Name.ToLower() == menuNames[0].ToLower())
                    {
                        firstLevel = menuCollection[i];
                        ExpandCollapsePattern value = (ExpandCollapsePattern)menuCollection[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                        value.Expand();
                        break;
                    }

                }

                AutomationElementCollection secondCollection = firstLevel.FindAll(TreeScope.Descendants,
                                                                                  new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));

                for (int i = 0; i < secondCollection.Count; i++)
                {
                    Console.WriteLine(secondCollection[i].Current.Name);
                    if (secondCollection[i].Current.Name.ToLower() == menuNames[1].ToLower())
                    {

                        InvokePattern value = (InvokePattern)secondCollection[i].GetCurrentPattern(InvokePattern.Pattern);
                        value.Invoke();
                        break;
                    }

                }

            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message);
            }
        }
        #endregion

        #region ClickBeamDesign
        private static void ClickBeamDesign(string windowtitle, string tabName, string buttonAutomationID, int processID)
        {

            try
            {


                UIAutomation_App.UIAutomationAction uAutomation = new UIAutomation_App.UIAutomationAction();
                uAutomation.logTofile(uAutomation._eLogPtah, "Inside ClickBeamDesign");

                uAutomation._processId = processID;

                AutomationElement winElement = uAutomation.GetUIAutomationWindow("title", windowtitle);
                uAutomation.uiAutomationCurrentParent = winElement;

                uAutomation.logTofile(uAutomation._eLogPtah, "Window Name: " + winElement.Current.Name);

                AutomationElement beamDesignTab = uAutomation.GetUIAutomationUltratab("name", tabName, -1);

                uAutomation.logTofile(uAutomation._eLogPtah, "Tab Name: " + beamDesignTab.Current.Name);

                SelectionItemPattern select = (SelectionItemPattern)beamDesignTab.GetCurrentPattern(SelectionItemPattern.Pattern);

                select.Select();
                if (beamDesignTab.Current.IsOffscreen.ToString().ToLower() == "true")
                {

                    uAutomation.logTofile(uAutomation._eLogPtah, "BeamDesign tab not visible");

                    AutomationElement btn = winElement.FindFirst(TreeScope.Descendants,
                                                            new System.Windows.Automation.AndCondition(
                                                                     new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                                                                      new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, buttonAutomationID)));

                    uAutomation.logTofile(uAutomation._eLogPtah, "Button:" + btn.Current.Name);

                    InvokePattern value = (InvokePattern)btn.GetCurrentPattern(InvokePattern.Pattern);
                    value.Invoke();
                    uAutomation.logTofile(uAutomation._eLogPtah, "Clicked Button");
                }


            }

            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }
        #endregion

        #region ClickControl
        private static void ClickControl(AutomationElement control)
        {

            try
            {
                AutoItX3Lib.AutoItX3 at = new AutoItX3Lib.AutoItX3();

                System.Windows.Point clickpoint1 = control.GetClickablePoint();
                Console.WriteLine("Got clickable Points ");
                double x = clickpoint1.X;
                double y = clickpoint1.Y;
                int x1 = Convert.ToInt32(x);
                int y1 = Convert.ToInt32(y);

                at.MouseMove(x1, y1, -1);
                try
                {
                    at.MouseClick("LEFT", x1, y1, 1);

                }
                catch (Exception e)
                {
                    throw new Exception(e.Message);
                }

            }
            catch (Exception ex)
            {

                throw new Exception("error on getclickablepoints :" + ex.Message);

            }
        }
        #endregion

        #region GetCustomWindowByName
        private static AutomationElement GetCustomWindowByName(string searchValue, int processid)
        {
            AutomationElementCollection GetWindowByName = null;
            AutomationElement windowName = null;

            try
            {
                Console.WriteLine("Inside GetCustomWindowByName of Wrapper: searchby " + searchValue);
                Console.WriteLine("Process Id " + processid);

                for (int i = 0; i < 5; i++)
                {
                    GetWindowByName = AutomationElement.RootElement.FindAll(TreeScope.Children,
                              new System.Windows.Automation.AndCondition(
                                           new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                                             new System.Windows.Automation.PropertyCondition(AutomationElement.ProcessIdProperty, processid)));

                    for (int j = 0; j < GetWindowByName.Count; j++)
                    {

                        if (GetWindowByName[j].Current.Name.Trim().ToLower() == searchValue.Trim().ToLower())
                        {
                            windowName = GetWindowByName[j];
                            break;
                        }
                    }
                    if (windowName != null)
                        break;
                }
                Console.WriteLine("Window name: " + windowName.Current.Name);


                return windowName;

            }

            catch (Exception ex)
            {


                throw new Exception(ex.Message);
            }

        }
        #endregion
        #endregion
    }

}




