
#region comments
/*  Version 1.0.0.15
  created on : 21 Dec 2011
  author :   Prasanna Hegde
 * This Solution uses already developed Controls for WPFAction with White and also incorporates UIauomation controls where White 
 * does not work
 * Author : Prasanna Hegde
 *
 * 
 */
/*  Version 2.0.0.1
 * Author : Prasanna Hegde
 * 6-Mar-2012 : Added controltype using full panes hiearchy 
 * 21-Mar-2012: Modified certain functions  to improve Perfomance
 * 23-Mar-2012: Added Aditonal logging for debugging.
 * 24-May-2012: Added controltype infratable and selectinfrarowbynumber method (control type : uiainfratablerowflatnumber ) (click on table object)
 * 5-jun-2012:  Added VErfication of wpfcheckbox for getActualDataForm whihc was wrongly using add data method ; 
 * 15-jun-2012: 1.) Added new parentType uiautomationchildWindow under both AddData & VerifyData
 *              2.) Added method closeGlobalWindow.
 * 21-Jun-2012 :1.) added verification method for wpflistview control    
 * 27-Jun-2012 Changed WebTable Verification Data from "ColumnName" to "FieldName"
 * 
 */
/* Version 2.0.0.2
 * 28-Jun-2012 Added Identifying button using helptext of button : added new control type uiautomationcombobox,uiautomationmenu,
 * uiautomationmenuitem,uiautomationtext,uiautomationimage
 * 3-July-2012: added a New action "clearwindow" only for "adddata":autoamtion developer to decide when to use it in strucuture sheet
 * 11-July-2012 : added Calendar Control specific to application MatBal "uicustomcalendarmatbal"
 and searchby using  automationid facilty for button 
 * *13-July-2012 :added verfication for datagrid Object 
 */
/* Version 2.0.0.3
 * Deepankar:12-Sep-2012 Added function for UIAutomationCheckBox
  */
/* Version 2.0.0.4
 * Author:Sneha
 * 14-Sep-2012 Added UIAutomationCheckbox case in GetActualData method
 * 14-Sep-2012 Added index case in GetUIAutomationcomboox method
 * 17-Sep-2012 Added uiautomationultratabitem case in GetActualData method
 * 17-Sep-2012 Added helptext case in GETUIAutomationmenuitem method
 */
/* Version 2.0.0.5
 * Author:Sneha
 * 26-Sep-2012 Added property UseWhite
  */
/* Version 2.0.0.6
 * Author:Sneha
 * 28-Sep-2012  Added Iskeyboardfocusable condition to use Setfoucs method for ultratabitem
 * 28-Sep-2012  Changed while condition in getuiautomationwindow function
 * 04-Oct-2012  Added Control Type --System.Windows.Automation.ControlType .Header-- as new control for .NET Table row selection
 */
/* Version 2.0.0.7
 * author: Sneha
 * 22-Oct-2012 Added new keyword 'condkeyboard' to send keys only when needed
 * 22-Oct-2012 Added else condition to select combobox item when the name property is same for all items
 */
/* Version 2.0.0.8
 * author: Sneha
 * 1-Nov-2012 Added Control type 'uiautomationradiobutton' in AddData and Verfiy Data
 * 1-Nov-2012 Added index case in GetUIAutomationedit method
 * 2-Nov-2012 Added automationid case in GetUIAutomationribbonbutton
 * 2-Nov-2012 Added Control type 'uiautomationdataitem' in AddData
 * 6-Nov-2012 Added text case in GetUIAutomationsyncfusionpane method
 * 14-Dec-2012 Added control type 'uiautomationspinner' in AddData
 * 14-Dec-2012 Added if condition in condkeyboard case to send key multiple times
 * 31-Dec-2012 Added control type 'uiautomationlistitem' in AddData
 * 2-Jan-2013 Added French value condition in uiautomationsyncfusionpane case to use tab key or right key
 * 4-Jan-2013 Added  try-catch for reading French value in variable _controlName1
 * 4-Mar-2013 Added try catch for controltype.setfocus method
 * 21-Aug-2013 Added Logic for System.Windows.Automation.ControlType .Cusotm with pane  InternetExplorer class to work in both WIn7 and WinXP
 * */
/* Version 2.0.0.9
 * author: Sneha
 * 21-Aug-2013: Added Control type 'rightclicktreeitem' in AddData
 * 21-Aug-2013: Added function 'RightClickControl'
 * 21-Aug-2013: Added 'uiautomationtreeitem' as a parent type
 * 22-Aug-2013: Updated control type 'uiautomationtreeitemclick' in AddData to handle single click, double click and right click. Also it will take care of when we pass control name using test data
 * 22-Aug-2013: Removed Control type 'rightclicktreeitem' in AddData
 * 23-Aug-2013: Added control type 'uiautomationtreeitem'
 * 23-Aug-2013: Chanegd control type ''uiautomationtreeitemclick' to the original one.
 * 29-Aug-2013: Added control type 'uiautomationbuttonverify' in 'GetActualDataForm'. It is to verify the properties of a button
 * 29-Aug-2013: Added property '_eLogPtah' to take logfile path from configuration file.
 * 06-Sep-2013: Added function 'VerifyDataGrid2Content' and 'GetActualDataGrid2Content' to verify the grid in which cell does not have value but its immediate child has . Also child element's control type is not uniform.
 * 25-Sep-2013: Added control types 'uiautomationtreeitemexpandk2' and 'uiautomationtreeitemcollapsek2' to expand and collapse the tree items respectively. It is used in K2 application.
 * 25-Sep-2013: Updated control type 'uiautomationtext' in AddData to add the cases for left click, right click and double click.
 * 30-Sep-2013: Added function 'GetUIAutomationTreeContent' and control type 'uiautomationtreeitems' in VerifyData'.
 * 30-Sep-2013: Added function 'IsColumnPresent' which is used to check whether 'Property' column exists or not in GetActualDataForm method.
 * 01-Oct-2013: updated getuiautomationwindow method
 * 01-Oct-2013: updated uiautomationcheckbox control type to check and uncheck it as per the control value
 * 01-Oct-2013: Added control type 'uiautomationselectlistitem' and 'uiautomationtogglebutton'
 * 15-Oct-2013: Updated function GetUIAutomationImage- added text case and updated automation id case
 * 15-Oct-2013: Updated uiautomationimage control type by adding switch case for single click, double click and left click
 * 16-Oct-2013: Updated getuiautomationedit function
 * 16-Oct-2013: Updated control type uiautomationcheckbox - added case for mouse click
 * 23-Oct-2013: Updated function GetUIAutomationDialogTextControl- Added searchby criteria, it was missing in existing function
 * 23-Oct-2013: Added control type uiautomationwindow and uiautomationtreeitem in function GetActualDataForm - to verify the window title and to verify the text or automation id of treeitem
 * 30-Oct-2013: Added controltype "uiautomationclicklistitem"
 * 30-Oct-2013: Updated GetUIAutomationbutton, GetUIAutomationribbonbuttob, GetUIAutomationcombobox, GetUIAutomationcheckbox methods
 * 30-Oct-2013: Updated GetUIAutomationTreeContent
 * 30-Oct-2013: Updated control type 'uiautomationbutton_verify' to verify text of a button
 * 31-Oct-2013: Added text case in getuiautomationdataitem
 * 31-Oct-2013: Added uiautomationtext control type in GetActualdatform method
 * 07-Nov-2013: Added uiautomationtabverify control type in GetActualdataform method
 * 12-Nov-2013: [Ashok] Added a select case "label" to identify text boxes using label when textbox didn't have any identification properties(name, Automation ID are blank) in GetUIAutomationEdit method
 * 12-Nov-2013: Added 'uiautomationcustomclickcontrol' control type
 * 13-Nov-2013: Added Aditonal logging in GetVerifyDataForm.
 * 15-Nov-2013: 
 * 1. [Deepankar] Added additional check in AddData For UIAutomationEdit where the control may not have valuepattern and hence input data using sendkeys
 * 2. [Deepankar] Changed logic for UIAutomationComboBox and UiAutomationEdit search by Label. Initially the logic was to search for label and 
 *      then check the next immediate sibling. Now added additional check to loop till we get the sibling of proper type i.e. Edit or ComboBox
 * 3. Added Console.Writeline in the Logging functions so that the same get captured and we need not write both console,writeline and log commands
 * 4. Added a select case "label" to identify Buttons using label when textbox didn't have any identification properties(name, Automation ID are blank) in GetUIAutomationEdit method
 * 5. Corrected the variable name in GetComboBox with Index parameter
 * * 18-Nov-2013: 
 * 1. [Deepankar] Moved Search by Label code to a function GetControlByLabel
 * 2. Added function GetControlByIndex
 * 4. Calling GetUIAutomationRadiobutton,GetUIAutomationButton and GetUiAutomationEdit index option using the GetControlByIndex function 
 * 5. Adding trim to searchvalue parameter in GetUiAutomationWindow
 * 6. Adding index in the log for GetControlByIndex
 * 18th -Nov
 *  1.In Fuction GetUiAutomationEdit for index case returning the value for method GetControlByIndex
 *  2. In Fuction GetCOntrol by index corrected the issue where it was Checking only for edit, it should be a based on the controltype passed as parameter
 * 20th Nov- Sneha - Updated GetUIAutomationwindow method
 * 21st Nov- Prasanna - Updated GetActualDataIEPaneTable method
 * 21st  Nov -Deepankar 
 *      1 Adding Start and end for AddDAta
        2. Adding count for number of times button is clicked
        3. Added return value for GetUIAutomationRAdioButton with Index parameter
 * 21st Nov - Sneha - Updated function GetActualDataGrid2Content
 * */
/*Version 2.0.0.10
 * 27 nov 2013 -- Added Common methods GetControlByName,GetControlByAutomationID,GetControlByHelpText
 *                ,GetControlByNameFromCollectionAndIndex,GetControlByAutomationIdFromCollectionAndIndex
 *                for controltypes uiautomationbutton,uiautomationedit,uiautomationtextarea,uiautomationtreeitem,uiautomationmenuitem,uiautomationtabitem
 *                and uiautomationtext
 * 29 Nov 2013 -- Added Doubleclick ,right click method for Controltype 'uiautomationcustomclickcontrol'
 * 02 Dec 2013 - Ashok Krishna K
 *      1. Added coded ui identifying cases for control types button, edit, list, listitem, tab, tabitem, radio button and checkbox.
 * 04 Dec 2013 - Sneha - updated function GetActualDataGrid2Content and GetUIAutomationRibbonButton
 * 04 Dec 2013 - Prasanna - Added 'isenabled' property in uiautomationtabverify controltype
 * 10 Dec 2013 - Prasanna - Added Invoke pattern for uiautomationmenuitem where clickable points are null
 *             - Removed reference to ribbonbutton.Current.Name in logging that was throwing unnecessary exceptions:
 * 12 DEC 2013 Added try catch in  method  verifyDataForm while clearing data from Template,Actual and Expected Data 
 * The catch is not thrown as an exception
 * 17 Feb 2014 - Ashok Krishna K
 *      1. Added drag keyword in action section to perform drag action. Only coordinates are supported now. Please provide start and end coordinates seperated by ';' in data sheet. For eg:- 13,12;15,16 
 * 28 Feb 2014 - Sneha - Added control type uiautomationheaderitem
 * 01 March 2014 - Ashok Krishna K 
 *      1. Added process Id condition in windows identification process
 *      2. Added search by name case in control identification process
 *      3. Added condition to check _controlValue before creating parent.
  *  05 March 2014 - Prasanna 
	  -  Added titlewildcard  for searchby 
	  -  Added Throw execptions in all catch blocks of all functions
 * 19-March-2014 - Sneha
 * - Updated method GetUIAutomationText: case automation id
 * 24-March 2014 - Sneha
 * - Updated GetUIAutomationWindow method
 * - Changed if condition whcih checks for control value in AddData and GetActualDataForm method
 * - Updated GetWindowByPartialName method: changed Descendants to Children
 * 25-March-2014 - Sneha
 * - Updated if condition to construct parent
 * 26-March-2014 - Sneha
 * - Changed property _globaltimeout to _Attempts.
 * 27-March-2014 - Sneha
 * - Added enum GetControlMethod
 * - Added method GetControlType
 * 28-March-2014 - Deepankar
 * - Throw exception in uiautomationultratabitem. Added select pattern in catch if clickcontrolfails
 * * - Throw exception in uiautomationbutton and uiautomationultratabitem
  * */
#region 2.1.0.0
//Deepankar: 01-Apr-2014 Added keyword uiautomationselectultratabitem. This is to handle ultratabs with selection pattern
//Deepankar:  01-Apr-2014: Added back Selection Pattern code in uiautomationultratabitem to avoid breaking of existing code
//Prasanna: 04-Arp-2014  Added Class for generating CSV log files for Add Data Function and made general logging optional with true and false flag
//Prasanna : 08-Apr-2014 Added Attempts for All generic Widnows and ControlTypes to wait for specifed attemps in construcitng controls.
//Deepankar: 10- Apr-2014 Added uiautomationselectultratabitem controltypee in function GetActualDataForm
//Sneha: 10-Apr-2014 Added condition in AddTexttoColumn method
//Sneha: 23-Apr-2014 Added conditionalwait case in actions to wait only when needed
//Sneha: 06-May-2014 Added automationid case in GetUIAutomationUltratab method
//Sneha: 07-May-2014 Updated uiautomaationcombobox control in AddData to select listitem using text
//Sneha: 07-May-2014 Shortened controltype names by using prefix"u"
//Sneha: 09-May-2014 updated uiautomationcombobox control type
//Sneha: 03-June-2014 Added GetControlByHelpTextFromCollectionAndIndex method
//Sneha: 03-June-2014 Added umenuiteminvoke control type
//Sneha: 03-June-2014 Added helptext case in GetUIAutomationRibbonButton method
//Sneha: 18-June-2014 Added name case in GetUIAutomationListcontent method
//Sneha: 04-July-2014 Added index case in GetUIAutomationPane method
//Sneha: 11-July-2014 Added GetUIAutomationListCollection method and ulistitemcollection control type
//Sneha: 11-July-2014 Added 'togglestate' property in ubuttonverify controltype.
// Deepankar: 18-July Modified code in ucombobox 
//            AddData function  1. Added try catch for Clickcontrol while clicking the text value in Listitem. If the clic fails then the same is tried using selection pattern of listitem    
//           GetActualDataForm function  2. Added try catch for Clickcontrol and Setfocus. Also removed if condition  where listitem name was matched with expected value  

#endregion
#region 2.2 22-July
//22-July 
//1. : Searching the window collection in children instead of descendents in function GetwindowbyName
//2.  Removed line in function GetUiAutomationButton where window.automationid was being logged. This was causing issue where window did not have automationid
//3.  Added keyword custominvoke 
#endregion
#region 25-July
//1. : Searching the window collection in Descendents instead of children as the change affects Matbal
//2:  Added code for ucombox in getActualformdata to handle situation where the listitem value is found in the text node and not the listitem name property
#endregion
# region 4th September 2014
//Added keywords for controls ulistitemselect,ulistitemclick
//Added code to handle ExpandCollapsePattern for comboboxes
//Searching in comboboxes using TreeScope.Descendants instead of children
#endregion
#region 5th September 2014
//Added action autoitdrag
#endregion
#region 9JUN2015
// Added Condition to indentify Window even if process id attached gets chnged ; Happend for Lowis
// Added Control Types pvmaskedit and pvcombobox for Lowis speific controls  9_JUN_2015
#endregion
#endregion

using System;
using System.Threading;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Configuration;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Data;
using System.IO;
using System.Globalization;
using System.Linq;
//using White.Core;
//using White.Core.UIItems.WindowItems;
//using White.Core.UIItems.MenuItems;
//using White.Core.UIItems.Finders;
//using White.Core.UIItems;
//using White.Core.UIItems.WindowStripControls;
//using White.Core.UIItems.ListBoxItems;
//using White.Core.UIItems.TreeItems;
//using White.Core.UIItems.TabItems;
using Helper;
using System.Data.Odbc;
using System.Windows.Automation;
using AutoItX3Lib;
using System.Management.Instrumentation;
using System.Management;
using Microsoft.Win32;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITest.Framework;
using Microsoft.VisualStudio.TestTools.UITest.Playback;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
using CUIT_app;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace UIAutomation_App
{
    public class UIAutomationAction
    {
        TestDataManagement testData = new TestDataManagement();
        TestDataManagement testDataHieararchy = new TestDataManagement();
        ReportsManagement report = new ReportsManagement();
        //WPF_App.WpfAction wpfapp = new WPF_App.WpfAction();
        AutoItX3Lib.AutoItX3 at = new AutoItX3Lib.AutoItX3();
        Helper.LogManagement action4 = new Helper.LogManagement();
        Helper.TestDataManagement action3 = new Helper.TestDataManagement();
        UIAutomationLog uilog = new UIAutomationLog();
        private string _error = "Error in Function ";
        private string _testcase = null;
        public string hrchyfile { get; set; }
        //this is used in generic exception inside catch block
        //  private string _searchTxtAuto = "The valid search criteria is Text & AutomationID.";
        public string uiAutoamtionreportPath
        {
            get;
            set;
        }
        public string uiAfileName
        {
            get;
            set;
        }
        //this is passed as a parameter in default case of switch where searchBy is both by text and automationin

        // private string _searchText = "The valid search criteria is Text.";
        //this is passed as a parameter in default case of switch where searchBy condition is only text

        //  private string _searchAutoID = "The valid search criteria is AutomationID.";
        //this is passed as a parameter in default case of switch where searchBy condition is only automationid

        //  private string _txtAutoIndex = "The valid Search criteria is Text, Automation Id & Indexed.";
        //this is passed as a parameter in default case of switch where searchBy condition is text, automationid and also index.

        //  private string _immediateParent = "";
        //this is used in AddData(int rowPosition) method
        //   private Menu _globalMenu = null;
        public string _testDataPath { get; set; }
        //this is used in AddData(int rowPosition) method
        //    public Application _application { get; set; }
        //    public Window _globalWindow { get; set; }
        //    public GroupBox _globalGroup { get; set; }
        public DataTable ActualData { get; set; }
        public Boolean UseWhite { get; set; }
        public Boolean UseDetaillog { get; set; }
        public String _eLogPtah { get; set; }
        public int _processId { get; set; }
        private int _attempts = 100;
        public int _Attempts
        {
            get
            {
                return _attempts;
            }
            set
            {
                _attempts = value;
            }
        }

        public bool actresult_K2;
        public String _reportsPath { get; set; }
        public String _reportsSectionPath { get; set; }
        public String ptestDataPath { get; set; }
        public String ptestCase { get; set; }
        public String pkeyword { get; set; }
        public enum GetControlMethod
        {
            Name,
            NameandIndex,
            AutomationID,
            AutomationIDandIndex
        };
        /// <summary>
        /// Used For Getting UIAutomation Application
        /// </summary>
        AutomationElement uiAutomationapp = AutomationElement.RootElement;
        public AutomationElement GetControlType(GetControlMethod method, System.Windows.Automation.ControlType controlType, string searchvalue, int _index = -1)
        {
            AutomationElement ctrl = null;
            switch (method.ToString().ToLower())
            {
                case "name":
                    try
                    {
                        ctrl = GetControlByName(controlType, searchvalue);
                    }
                    catch (Exception ex)
                    {
                        logTofile(_eLogPtah, "Execption from [Getcontroltype:name]" + ex.Message.ToString());
                        logTofile(_eLogPtah, "Line number is Source File where Error is encountered is:" + GetStacktrace(ex));
                        throw new Exception("Execption from [Getcontroltype:name]" + ex.Message.ToString());
                    }
                    break;
                case "nameandindex":
                    try
                    {
                        ctrl = GetControlByNameFromCollectionAndIndex(controlType, searchvalue, _index);
                    }
                    catch (Exception ex)
                    {
                        logTofile(_eLogPtah, "Execption from [Getcontroltype:nameandindex]" + ex.Message.ToString());
                        logTofile(_eLogPtah, "Line number is Source File where Error is encountered is:" + GetStacktrace(ex));
                        throw new Exception("Execption from [Getcontroltype:nameandindex]" + ex.Message.ToString());
                    }
                    break;
                case "automationid":
                    try
                    {
                        ctrl = GetControlByAutomationId(controlType, searchvalue);
                    }
                    catch (Exception ex)
                    {
                        logTofile(_eLogPtah, "Execption from [Getcontroltype:automationid]" + ex.Message.ToString());
                        logTofile(_eLogPtah, "Line number is Source File where Error is encountered is:" + GetStacktrace(ex));
                        throw new Exception("Execption from [Getcontroltype:automationid]" + ex.Message.ToString());
                    }
                    break;
                case "automationidandindex":
                    try
                    {
                        ctrl = GetControlByAutomationIdFromCollectionAndIndex(controlType, searchvalue, _index);
                    }
                    catch (Exception ex)
                    {
                        logTofile(_eLogPtah, "Execption from [Getcontroltype:automationidandindex]" + ex.Message.ToString());
                        logTofile(_eLogPtah, "Line number is Source File where Error is encountered is:" + GetStacktrace(ex));
                        throw new Exception("Execption from [Getcontroltype:automationidandindex]" + ex.Message.ToString());
                    }
                    break;
            }
            return ctrl;
        }
        /// <summary>
        /// Used for referring to the last parent window under which rest of the automation elements (controls) are found. This is controlType.Window when checked in UI spy.
        public AutomationElement uiAutomationWindow = null;
        /// <summary>
        /// Used for referring to the last parent object under which rest of the automation elements (controls) are found. This could window, pane or any such objects
        /// </summary>
        public AutomationElement uiAutomationCurrentParent = null;

        // *****************Object Identification Functions for UIautomation *************************************************
        #region UIAObjectLibrary
        /// <summary>
        /// GetUIAutomationTextarea : ContorlType.Document control.
        /// use "uiautomationtextarea"  under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext  </param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationTextarea(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "[GetUIAutomationTextarea]: Outside  Try for CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
            AutomationElement textarea = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();
            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationTextarea]: Inside Try for  CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                switch (searchBy.Trim().ToLower())
                {

                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationTextare]: Current Parent Name: " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");
                                textarea = GetControlByName(System.Windows.Automation.ControlType.Document, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by name & index from collection:");
                                textarea = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.Document, searchValue, index);

                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect textea = " + duration2);


                            break;
                        }
                    case "name":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationTextare]: Current Parent Name: " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");
                                textarea = GetControlByName(System.Windows.Automation.ControlType.Document, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by name & index from collection:");
                                textarea = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.Document, searchValue, index);

                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect textea = " + duration2);


                            break;
                        }
                    case "automationid":
                        {
                            #region CommentedCode
                            //logTofile(_eLogPtah, "[GetUIAutomationTextarea]: Current Parent Automationid " + uiAutomationCurrentParent.Current.AutomationId.ToString());
                            //AutomationElementCollection textareacol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                            //     new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Document));
                            //int j = 0;
                            //for (int i = 0; i < textareacol.Count; i++)
                            //{

                            //    if (textareacol[i].Current.AutomationId == searchValue)
                            //    {
                            //        logTofile(_eLogPtah, "[GetUIAutomationTextarea]: searching for:  " + searchValue + "   obtained:   " + textareacol[i].Current.AutomationId);
                            //        if (index <= 0)
                            //        {
                            //            textarea = textareacol[i];
                            //            break;
                            //        }
                            //        else
                            //        {
                            //            if (j == index)
                            //                textarea = textareacol[i];
                            //        }
                            //        j++;
                            //    }
                            //}

                            #endregion
                            logTofile(_eLogPtah, "[GetUIAutomationTextarea]: Current Parent Name: " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by Automation ID:");
                                textarea = GetControlByAutomationId(System.Windows.Automation.ControlType.Document, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by name & index from collection:");
                                textarea = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.Document, searchValue, index);

                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect textea = " + duration2);
                            break;

                        }

                }
                return textarea;


            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationTextarea ]: Generic exception encoutered");
                Console.WriteLine("Exception: " + ex.Message);
                throw new Exception(ex.Message);
            }

        } //fucntion end 
        /// <summary>
        /// This function is used to identify automation element with controltype.custom that supports Invoke Pattern.
        /// use "uiautomationcustominvokecontrol"  under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationCustominvokecontrol(string searchBy, string searchValue, int index)
        {
            AutomationElement customcontrol = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                    case "name":
                        {
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");

                                customcontrol = GetControlByName(System.Windows.Automation.ControlType.Custom, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                customcontrol = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.Custom, searchValue, index);

                            }

                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect customcontrol = " + duration2);

                            break;
                        }
                    case "automationid":
                        {
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by AutoID");

                                customcontrol = GetControlByAutomationId(System.Windows.Automation.ControlType.Custom, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                customcontrol = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.Custom, searchValue, index);
                            }

                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect customcontrol = " + duration2);

                            break;
                        }

                }
                return customcontrol;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationCustominvokecontrol]: exeption encoutered" + ex.Message.ToString());
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        }
        /// <summary>
        /// This function is used to identify automation element with controltype.custom that supports ValuePattern.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationCustomvaluecontrol(string searchBy, string searchValue, int index)
        {
            AutomationElement customvaluecontrol = null;
            try
            {
                AutomationElementCollection customvaluecontrolcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                     new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom
                         ));
                if (customvaluecontrolcol == null || customvaluecontrolcol.Count <= 0)
                {
                    logTofile(_eLogPtah, "[GetUIAutomationCustomvaluecontrol]: Could not Find Object with Given Search conditions in application.");
                }
                int j = 0;
                for (int i = 0; i < customvaluecontrolcol.Count; i++)
                {
                    logTofile(_eLogPtah, "[GetUIAutomationCustomvaluecontrol]: Inside Collection of this ciontrol typer.");
                    if (customvaluecontrolcol[i].Current.Name == searchValue || customvaluecontrolcol[i].Current.AutomationId == searchValue)
                    {

                        if (index <= 0)
                        {
                            customvaluecontrol = customvaluecontrolcol[i];
                            break;
                        }
                        else
                        {
                            if (j == index)
                                customvaluecontrol = customvaluecontrolcol[i];
                        }

                    }
                    else
                    {
                        logTofile(_eLogPtah, "[GetUIAutomationCustomvaluecontrol]: looking purely for index as no automation id nor Name :) .");
                        if (index <= 0)
                        {
                            customvaluecontrol = customvaluecontrolcol[0];
                            logTofile(_eLogPtah, "[GetUIAutomationCustomvaluecontrol]: Custom control was returned.");
                        }
                        else
                        {
                            if (j == index)
                                customvaluecontrol = customvaluecontrolcol[j];
                            logTofile(_eLogPtah, "[GetUIAutomationCustomvaluecontrol]: Custom control was returned.");
                        }
                    }
                    j++;
                }
                return customvaluecontrol;





            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationCustomvaluecontrol]: exeption encoutered" + ex.Message.ToString());
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        }
        /// <summary>
        /// This function is used to identify tabitem i,e controltype.tabitem 
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationUltratab(string searchBy, string searchValue, int index)
        {
            AutomationElement ultratabitem = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                    case "name":
                        {
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");

                                ultratabitem = GetControlByName(System.Windows.Automation.ControlType.TabItem, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index & name ");

                                ultratabitem = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.TabItem, searchValue, index);

                            }

                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect tabitem = " + duration2);

                            break;
                        }
                    case "automationid":
                        {
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automationid");

                                ultratabitem = GetControlByAutomationId(System.Windows.Automation.ControlType.TabItem, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index & name ");

                                ultratabitem = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.TabItem, searchValue, index);

                            }

                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect tabitem = " + duration2);

                            break;
                        }

                }
                return ultratabitem;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationUltratab]:Genreic  exeption encoutered");
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        } //fucntion end 

        /// <summary>
        /// This function is used to identify System.Windows.Automation.ControlType .Custom This Fucntion is to be used when we come across SyncFustion grid Controls.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationsyncfusionpane(string searchBy, string searchValue, int index)
        {
            AutomationElement syncfusionpane = null;


            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "automationid":
                        {
                            if (index == -1)
                            {
                                syncfusionpane = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue));
                                return syncfusionpane;
                            }
                            else
                            {
                                //break;
                                AutomationElementCollection syncfusionpanecol = uiAutomationWindow.FindAll(TreeScope.Descendants,
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue));
                                syncfusionpane = syncfusionpanecol[index];
                                return syncfusionpane;
                            }

                        }
                    case "name":
                    case "text":
                        {
                            if (index == -1)
                            {
                                syncfusionpane = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue));
                                return syncfusionpane;
                            }
                            else
                            {
                                //break;
                                AutomationElementCollection syncfusionpanecol = uiAutomationWindow.FindAll(TreeScope.Descendants,
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue));
                                syncfusionpane = syncfusionpanecol[index];
                                return syncfusionpane;
                            }

                        }


                }
                return syncfusionpane;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationsyncfusionpane]:  exception encoutered");
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        } //fucntion end 
        /// <summary>
        /// This function is used to identify buttons(System.Windows.Automation.ControlType .Button) This function is to be used when invoke patern is present 
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationbutton(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "[GetUIAutomationbutton]: Outside  Try for CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
            AutomationElement button = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();

            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationbutton]: Inside Try for  CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                switch (searchBy.Trim().ToLower())
                {

                    case "label":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationbutton]: Current Parent Name " + uiAutomationCurrentParent.Current.AutomationId.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            logTofile(_eLogPtah, "Find by label '" + searchValue + "'");
                            logTofile(_eLogPtah, "Finding current root");
                            button = GetControlByLabel(System.Windows.Automation.ControlType.Button, searchValue);
                            #region Commented moved to function
                            //AutomationElement currentRoot = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants
                            //    , new System.Windows.Automation.AndCondition(new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Text)
                            //    , new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));

                            //var foundButtonControl = false;
                            //AutomationElement buttonControl;
                            //while (foundButtonControl == false)
                            //{

                            //    buttonControl = TreeWalker.ControlViewWalker.GetNextSibling(currentRoot);

                            //    logTofile(_eLogPtah, "Current root is " + buttonControl.Current.System.Windows.Automation.ControlType .ProgrammaticName);
                            //    if (buttonControl.Current.System.Windows.Automation.ControlType  == System.Windows.Automation.ControlType .Button)
                            //    {
                            //        logTofile(_eLogPtah, "Found Button: " + buttonControl.Current.System.Windows.Automation.ControlType .ProgrammaticName);
                            //        logTofile(_eLogPtah, "Process Id: " + buttonControl.Current.ProcessId.ToString());
                            //        button = buttonControl;
                            //        foundButtonControl = true;
                            //    }
                            //    else
                            //    {
                            //        currentRoot = buttonControl;
                            //    }
                            //} 
                            #endregion
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect Button = " + duration2);
                            break;

                        }


                    case "text":
                    case "name":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationbutton]: Current Parent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");
                                button = GetControlByName(System.Windows.Automation.ControlType.Button, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                logTofile(_eLogPtah, "Function : GetUIAutomationbutton ------> before buttons search : ");
                                button = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.Button, searchValue, index);
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect button = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationbutton]: ------> Button name :was found  =" + button.Current.Name.ToString());
                            break;
                        }

                    case "helptext":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationbutton]:------>  " + uiAutomationCurrentParent.Current.HelpText + "Helptext" + uiAutomationCurrentParent.Current.Name);
                            button = GetControlByHelpText(System.Windows.Automation.ControlType.Button, searchValue);
                            logTofile(_eLogPtah, "[GetUIAutomationbutton]: ------> Button helptext :was found  =" + button.Current.HelpText.ToString());
                            break;
                        }

                    case "automationid":
                        {

                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automation id");
                                button = GetControlByAutomationId(System.Windows.Automation.ControlType.Button, searchValue);
                                logTofile(_eLogPtah, "[GetUIAutomationbutton]: ------> Button automation id :was found  =" + button.Current.AutomationId.ToString());
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                logTofile(_eLogPtah, "Function : GetUIAutomationbutton ------> before buttons search : ");
                                button = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.Button, searchValue, index);
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect button = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationbutton]: ------> Button name :was found  =" + button.Current.AutomationId.ToString());
                            break;
                        }

                    case "index":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationbutton]: ------> call by Index");
                            button = GetControlByIndex(System.Windows.Automation.ControlType.Button, index);
                            #region Code Commented Moved to Common Function
                            //AutomationElementCollection buttoncol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                            //    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Button));

                            //int ind = Convert.ToInt32(searchValue);
                            //for (int k = 0; k < buttoncol.Count; k++)
                            //{
                            //    if (ind == k)
                            //    {
                            //        button = buttoncol[k];
                            //        logTofile(_eLogPtah, " Index found");
                            //        break;
                            //    }
                            //    else
                            //    {
                            //        logTofile(_eLogPtah, " Index not found");
                            //    }

                            //} 
                            #endregion
                            break;

                        }


                }//switdh
                return button;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationbutton]:  exception encoutered" + ex.Message);
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        }
        /// <summary>
        /// This function is used to identify buttons(System.Windows.Automation.ControlType .Button) This function is to be used when invoke patern is not present
        /// or Invoke pattern actions do not work as intened ("AutoIt" clicks will be used for such buttons .)
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>

        public AutomationElement GetUIAutomationRibbonButton(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: Outside  Try for CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
            AutomationElement ribbonbutton = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();

            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: Inside Try for  CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                switch (searchBy.Trim().ToLower())
                {

                    case "label":
                        {

                            logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: Current Parent Name " + uiAutomationCurrentParent.Current.AutomationId.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));


                            logTofile(_eLogPtah, "Find by label '" + searchValue + "'");

                            logTofile(_eLogPtah, "Finding current root");

                            ribbonbutton = GetControlByLabel(System.Windows.Automation.ControlType.Button, searchValue);
                            #region Commented moved to function
                            //AutomationElement currentRoot = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants
                            //    , new System.Windows.Automation.AndCondition(new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Text)
                            //    , new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));

                            //var foundButtonControl = false;
                            //AutomationElement buttonControl;
                            //while (foundButtonControl == false)
                            //{

                            //    buttonControl = TreeWalker.ControlViewWalker.GetNextSibling(currentRoot);

                            //    logTofile(_eLogPtah, "Current root is " + buttonControl.Current.System.Windows.Automation.ControlType .ProgrammaticName);
                            //    if (buttonControl.Current.System.Windows.Automation.ControlType  == System.Windows.Automation.ControlType .Button)
                            //    {
                            //        logTofile(_eLogPtah, "Found Button: " + buttonControl.Current.System.Windows.Automation.ControlType .ProgrammaticName);
                            //        logTofile(_eLogPtah, "Process Id: " + buttonControl.Current.ProcessId.ToString());
                            //        button = buttonControl;
                            //        foundButtonControl = true;
                            //    }
                            //    else
                            //    {
                            //        currentRoot = buttonControl;
                            //    }
                            //} 
                            #endregion


                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect RibbonButton = " + duration2);
                            break;

                        }


                    case "text":
                    case "name":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: Current Parent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");
                                ribbonbutton = GetControlByName(System.Windows.Automation.ControlType.Button, searchValue);
                                #region Commented moved to function
                                //ribbonbutton = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                //   new System.Windows.Automation.AndCondition(
                                //         new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Button),
                                //     new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));
                                #endregion
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                ribbonbutton = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.Button, searchValue, index);
                                #region Commented moved to function
                                //logTofile(_eLogPtah, "Function : GetUIAutomationRibbonButton ------> before buttons search : ");
                                //AutomationElementCollection buttoncol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                //    new System.Windows.Automation.AndCondition(
                                //    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Button),
                                //    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));
                                //ribbonbutton = buttoncol[index];
                                #endregion
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect button = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: ------> Button name :was found  =" + ribbonbutton.Current.Name.ToString());
                            break;
                        }

                    case "helptext":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]:------>  " + uiAutomationCurrentParent.Current.HelpText + "Helptext" + uiAutomationCurrentParent.Current.Name);
                            if ((index == -1) == true)
                            {
                                ribbonbutton = GetControlByHelpText(System.Windows.Automation.ControlType.Button, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                ribbonbutton = GetControlByHelpTextFromCollectionAndIndex(System.Windows.Automation.ControlType.Button, searchValue, index);
                            }
                            #region Commented moved to function
                            //ribbonbutton = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                            //    new System.Windows.Automation.AndCondition(
                            //    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Button),
                            //    new System.Windows.Automation.PropertyCondition(AutomationElement.HelpTextProperty, searchValue)));

                            //logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: ------> Button helptext :was found  =" + ribbonbutton.Current.HelpText.ToString());
                            #endregion
                            break;
                        }

                    case "automationid":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: Current Parent Name " + uiAutomationCurrentParent.Current.AutomationId.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automation id");
                                ribbonbutton = GetControlByAutomationId(System.Windows.Automation.ControlType.Button, searchValue);
                                #region commented moved to function
                                //ribbonbutton = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                //   new System.Windows.Automation.AndCondition(
                                //   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Button),
                                //   new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue)));

                                //logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: ------> Button automation id :was found  =" + ribbonbutton.Current.AutomationId.ToString());
                                #endregion
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                ribbonbutton = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.Button, searchValue, index);
                                #region commented moved to function
                                //logTofile(_eLogPtah, "Function : GetUIAutomationRibbonButton ------> before buttons search : ");

                                //AutomationElementCollection buttoncol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                //    new System.Windows.Automation.AndCondition(
                                //    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Button),
                                //    new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue)));

                                //ribbonbutton = buttoncol[index];
                                #endregion
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect button = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: ------> Button name :was found  =" + ribbonbutton.Current.AutomationId.ToString());
                            break;
                        }

                    case "index":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationRibbonButton]: ------> call by Index");
                            ribbonbutton = GetControlByIndex(System.Windows.Automation.ControlType.Button, index);

                            #region Code Commented Moved to Common Function
                            //AutomationElementCollection buttoncol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                            //    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Button));

                            //int ind = Convert.ToInt32(searchValue);
                            //for (int k = 0; k < buttoncol.Count; k++)
                            //{
                            //    if (ind == k)
                            //    {
                            //        button = buttoncol[k];
                            //        logTofile(_eLogPtah, " Index found");
                            //        break;
                            //    }
                            //    else
                            //    {
                            //        logTofile(_eLogPtah, " Index not found");
                            //    }

                            //} 
                            #endregion
                            break;

                        }


                }//switdh
                return ribbonbutton;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationbutton]:  exception encoutered" + ex.Message);
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        }

        public AutomationElement GetUIAutomationSpinner(string searchBy, string searchValue, int index)
        {
            AutomationElement UISpinner = null;

            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationSpinner]: Inside ");
                            AutomationElementCollection spinnercol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Spinner
                                     ));
                            int j = 0;
                            for (int i = 0; i < spinnercol.Count; i++)
                            {
                                logTofile(_eLogPtah, "[GetUIAutomationSpinner]: Inside Collection  ");
                                if (spinnercol[i].Current.Name == searchValue)
                                {

                                    if (index <= 0)
                                    {
                                        UISpinner = spinnercol[i];
                                        break;
                                    }
                                    else
                                    {
                                        if (j == index)
                                            UISpinner = spinnercol[i];
                                    }
                                    j++;
                                }
                            }
                            return UISpinner;
                            //break;
                        }

                    case "automationid":
                        {
                            logTofile(_eLogPtah, "Function : GetUIAutomationSpinner ------> before search : ");
                            logTofile(_eLogPtah, "[GetUIAutomationSpinner]:------>  " + uiAutomationCurrentParent.Current.AutomationId + " text " + uiAutomationCurrentParent.Current.Name);
                            AutomationElementCollection spinnercol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Spinner));

                            int j = 0;
                            logTofile(_eLogPtah, "[GetUIAutomationSpinner]: ------> Spinners present on window(CurrentParent) : " + uiAutomationCurrentParent.Current.Name + " = " + spinnercol.Count);
                            for (int i = 0; i < spinnercol.Count; i++)
                            {

                                if (spinnercol[i].Current.AutomationId.ToString() == searchValue)
                                {

                                    if (index <= 0)
                                    {
                                        UISpinner = spinnercol[i];
                                        break;
                                    }
                                    else
                                    {
                                        if (j == index)
                                            UISpinner = spinnercol[i];
                                    }
                                    j++;
                                }
                            }
                            logTofile(_eLogPtah, "[GetUIAutomationSpinner]: ------> Spinner automation id :was found  =" + UISpinner.Current.AutomationId.ToString());
                            //  return button;
                            break;
                        }


                }
                return UISpinner;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Function : GetUIAutomationSpinner  exeption encoutered");
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        }
        /// <summary>
        /// This function is used to identify treeitems (System.Windows.Automation.ControlType .treeitems) This function is to be used when invoke patern is not present
        /// or Invoke pattern actions do not work as intened ("AutoIt" clicks will be used for such treeitems .)
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationtreeitem(string searchBy, string searchValue, int index)
        {
            AutomationElement treeitem = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");

                                treeitem = GetControlByName(System.Windows.Automation.ControlType.TreeItem, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                treeitem = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.TreeItem, searchValue, index);

                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect treeitem = " + duration2);

                        }
                        break;

                }
                return treeitem;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationtreeitem]: exeption encoutered");
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        }
        /// <summary>
        /// This function is used get target window controltype.window from all available windows in desktop.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationWindow(string searchBy, string searchValue)
        {
            bool windowsearch = true;
            if (uiAutomationWindow != null && uiAutomationCurrentParent != null)
            {
                uiAutomationCurrentParent = uiAutomationWindow;
                if ((uiAutomationWindow.Current.Name.ToLower().Contains(searchValue.ToLower())) || (uiAutomationWindow.Current.AutomationId == searchValue.ToLower()))
                {
                    logTofile(_eLogPtah, "[GetUIAutomationWindow]:---> seaching same MasterParentWindow was Avaoided to save time !!  " + uiAutomationWindow.Current.Name.ToLower() + "--" + searchValue.ToLower());
                    windowsearch = false;
                }
            }

            if (windowsearch == true)
            { // start of First IF 
                logTofile(_eLogPtah, "[GetUIAutomationWindow]:---> Section :---> searching for Parent Window :  ");
                try
                { //Start of Try 
                    logTofile(_eLogPtah, "[GetUIAutomationWindow]:---> Inside Try   ");
                    //    int numwait = 0;
                    switch (searchBy.ToLower())
                    { //Start of Switch 
                        case "name":
                        case "text":
                        case "title":
                            { //start of case title 
                                logTofile(_eLogPtah, "using  title criteria:  " + searchValue);
                                uiAutomationWindow = GetWindowByName(searchValue);
                                if (uiAutomationWindow == null)
                                {
                                    logTofile(_eLogPtah, "Searching window using collection and contains");
                                    for (var i = 0; i < _Attempts; i++)
                                    {
                                        System.Threading.Thread.Sleep(1);
                                        logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");

                                        uiAutomationWindow = GetWindowByPartialName(searchValue);

                                        if (uiAutomationWindow != null)
                                        {
                                            break;
                                        }
                                    }

                                }
                                if (uiAutomationWindow != null)
                                {
                                    logTofile(_eLogPtah, "window : " + uiAutomationWindow.Current.Name.ToString());
                                }
                                else
                                {
                                    logTofile(_eLogPtah, "Window with Title: " + searchValue + "was not found");
                                }
                                if (UseWhite == true)
                                {
                                    //wpfapp._application = _application;
                                    //try
                                    //{
                                    //    _globalWindow = wpfapp.GetWPFWindow(searchValue.Trim());
                                    //}
                                    //catch (Exception e)
                                    //{
                                    //    logTofile(_eLogPtah, "exception in getwpfwidow was :" + e.Message);
                                    //}
                                }
                                else
                                {
                                    logTofile(_eLogPtah, "White window not initialized");
                                }
                                logTofile(_eLogPtah, "after");

                                //    return uiAutomationWindow;
                                if (uiAutomationWindow != null)
                                {
                                    WindowPattern winpat = (WindowPattern)uiAutomationWindow.GetCurrentPattern(WindowPattern.Pattern);

                                    try   // This is to handle situations when the window cannot be setfocus
                                    {
                                        winpat.SetWindowVisualState(WindowVisualState.Maximized);
                                        uiAutomationWindow.SetFocus();
                                    }
                                    catch (Exception ex)
                                    {
                                        logTofile(_eLogPtah, "[GetUIAutomationWindow]: Focus could not be set adue to  " + ex.Message.ToString());
                                    }
                                }
                                break;
                            } // end of title case 
                        case "titlewildcard":
                            { //start of case titlewildcard 
                                logTofile(_eLogPtah, "using  title wildcard  criteria:  " + searchValue);
                                for (var i = 0; i < _Attempts; i++)
                                {
                                    System.Threading.Thread.Sleep(1);
                                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");

                                    uiAutomationWindow = GetWindowByPartialName(searchValue);

                                    if (uiAutomationWindow != null)
                                    {
                                        break;
                                    }
                                }
                                if (uiAutomationWindow != null)
                                {
                                    logTofile(_eLogPtah, "window : " + uiAutomationWindow.Current.Name.ToString());
                                }
                                else
                                {
                                    logTofile(_eLogPtah, "Window with Title wildcard : " + searchValue + "was not found");
                                }
                                if (UseWhite == true)
                                {
                                    //wpfapp._application = _application;
                                    //try
                                    //{
                                    //    _globalWindow = wpfapp.GetWPFWindow(searchValue.Trim());
                                    //}
                                    //catch (Exception e)
                                    //{
                                    //    logTofile(_eLogPtah, "exception in getwpfwidow was :" + e.Message);
                                    //}
                                }
                                else
                                {
                                    //logTofile(_eLogPtah, "White window not initialized");
                                }
                                logTofile(_eLogPtah, "after");

                                //    return uiAutomationWindow;
                                if (uiAutomationWindow != null)
                                {
                                    WindowPattern winpat = (WindowPattern)uiAutomationWindow.GetCurrentPattern(WindowPattern.Pattern);

                                    try   // This is to handle situations when the window cannot be setfocus
                                    {
                                        winpat.SetWindowVisualState(WindowVisualState.Maximized);
                                        uiAutomationWindow.SetFocus();
                                    }
                                    catch (Exception ex)
                                    {
                                        logTofile(_eLogPtah, "[GetUIAutomationWindow]: Focus could not be set adue to  " + ex.Message.ToString());
                                    }
                                }
                                break;
                            } // end of titlewildcard case 
                        case "titleexact":
                            {

                                logTofile(_eLogPtah, "using  title exact  criteria:  " + searchValue);

                                for (var i = 0; i < _Attempts; i++)
                                {
                                    System.Threading.Thread.Sleep(1);
                                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");

                                    uiAutomationWindow = GetWindowByName(searchValue);

                                    if (uiAutomationWindow != null)
                                    {

                                        break;
                                    }
                                }
                                if (uiAutomationWindow != null)
                                {
                                    logTofile(_eLogPtah, "window : " + uiAutomationWindow.Current.Name.ToString());
                                }
                                else
                                {
                                    logTofile(_eLogPtah, "Window with Exact : " + searchValue + "was not found");
                                }
                                if (UseWhite == true)
                                {
                                    /*  wpfapp._application = _application;
                                      try
                                      {
                                          _globalWindow = wpfapp.GetWPFWindow(searchValue.Trim());
                                      }
                                      catch (Exception e)
                                      {
                                          logTofile(_eLogPtah, "exception in getwpfwidow was :" + e.Message);
                                      } */
                                }
                                else
                                {
                                    logTofile(_eLogPtah, "White window not initialized as option is disabled");
                                }


                                if (uiAutomationWindow != null)
                                {
                                    WindowPattern winpat = (WindowPattern)uiAutomationWindow.GetCurrentPattern(WindowPattern.Pattern);

                                    try   // This is to handle situations when the window cannot be setfocus
                                    {
                                        winpat.SetWindowVisualState(WindowVisualState.Maximized);
                                        uiAutomationWindow.SetFocus();
                                    }
                                    catch (Exception ex)
                                    {
                                        logTofile(_eLogPtah, "[GetUIAutomationWindow]: Focus could not be set adue to  " + ex.Message.ToString());
                                    }
                                }
                                break;
                            }

                        case "automationid":
                            {
                                uiAutomationWindow = GetWindowByAutomationId(searchValue);
                                if (UseWhite == true)
                                {
                                    /* wpfapp._application = _application;
                                     try
                                     {
                                         _globalWindow = wpfapp.GetWPFWindow(searchValue.Trim());
                                     }
                                     catch (Exception e)
                                     {
                                         logTofile(_eLogPtah, "exception in getwpfwidow was :" + e.Message);
                                     } */
                                }
                                else
                                {
                                    // logTofile(_eLogPtah, "White window not initialized");
                                }
                                WindowPattern winpat = (WindowPattern)uiAutomationWindow.GetCurrentPattern(WindowPattern.Pattern);
                                try   // This is to handle situations when the window cannot be setfocus
                                {
                                    winpat.SetWindowVisualState(WindowVisualState.Maximized);
                                    uiAutomationWindow.SetFocus();
                                }
                                catch (Exception ex)
                                {
                                    logTofile(_eLogPtah, "[GetUIAutomationWindow]: Focus could not be set adue to  " + ex.Message.ToString());
                                }
                            } // end of automation id case 
                            break;
                    } //Switch end
                    uiAutomationCurrentParent = uiAutomationWindow;
                    return uiAutomationWindow;

                } //end of try 


                            //   logTofile(_eLogPtah, System.DateTime.Now.ToString() + "Waiting for window" + Environment.NewLine);
                catch (Exception ex)
                {
                    logTofile(_eLogPtah, "Fucntion GetUIAutomationWindow: exeption encoutered");
                    throw new Exception(_error + "UIAutoamtionWindow:" + System.Environment.NewLine + ex.Message);

                }
            } //end of try catch 
            else
            {
                uiAutomationCurrentParent = uiAutomationWindow;
                return uiAutomationWindow;
            }
        }
        /// <summary>
        /// This function is used get target pane  controltype.pane under  a given window
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>

        public AutomationElement GetUIAutomationPane(string searchBy, string searchValue)
        {
            try
            {
                AutomationElement paneObj = null;
                switch (searchBy.ToLower())
                {
                    case "automationid":
                        paneObj = uiAutomationCurrentParent.FindFirst(TreeScope.Children,
                        new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue, PropertyConditionFlags.IgnoreCase));
                        uiAutomationCurrentParent = paneObj;
                        return uiAutomationCurrentParent;
                    case "name":
                    case "text":
                        paneObj = uiAutomationCurrentParent.FindFirst(TreeScope.Children,
                        new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue, PropertyConditionFlags.IgnoreCase));
                        uiAutomationCurrentParent = paneObj;
                        return uiAutomationCurrentParent;
                    case "index":
                        paneObj = GetControlByIndex(System.Windows.Automation.ControlType.Pane, Int32.Parse(searchValue));
                        uiAutomationCurrentParent = paneObj;
                        return uiAutomationCurrentParent;
                    default:
                        return null;

                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// This function is used identify a row from Grouped Table ( Infragistics Table Controls..that have been grouped by column names,
        /// wherein we would like to select a row based on text (typically expand the row by clicking on it)
        /// use "uiainfratablerow" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationGroupInfraTableRow(string searchBy, string searchValue, int index)
        {
            try
            {
                AutomationElement tablearow = null;
                AutomationElementCollection table = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                        new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom));
                for (int k = 0; k < table.Count; k++)
                {

                    ValuePattern valpat = SupportsValuePattern(table[k]);
                    logTofile(_eLogPtah, "[GetUIAutomationGroupInfraTableRow] : -->Passing row number -- " + k);
                    if (valpat != null)
                    {
                        logTofile(_eLogPtah, "[GetUIAutomationGroupInfraTableRow] :  --> table  " + k + "foond some paatern supported");
                        logTofile(_eLogPtah, "[GetUIAutomationGroupInfraTableRow] : --> detected " + valpat.Current.Value);
                        if (valpat.Current.Value.Contains(searchValue))
                        {


                            System.Threading.Thread.Sleep(100);
                            tablearow = table[k];
                            return tablearow;


                        }


                    }

                }
                return tablearow;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationGroupInfraTableRow] :  exeption encoutered" + ex.Message);
                throw new Exception(ex.Message);
            }

        }
        /// <summary>
        /// This function seems to be obslolete is not be used
        /// </summary>
        /// <param name="searchBy"></param>
        /// <param name="searchValue"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationFlatInfraTableRow(string searchBy, string searchValue, int index)
        {
            try
            {
                AutomationElement tablearow = null;
                AutomationElementCollection table = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                        new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom));
                for (int k = 0; k < table.Count; k++)
                {

                    ValuePattern valpat = SupportsValuePattern(table[k]);
                    logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRow] -->: Passing row number -- " + k);
                    if (valpat != null)
                    {
                        logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRow]  --> table  " + k + "foond some paatern supported");
                        logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRow] --> detected " + valpat.Current.Value);
                        if (valpat.Current.Value.Contains(searchValue))
                        {


                            System.Threading.Thread.Sleep(100);
                            tablearow = table[k];
                            return tablearow;


                        }


                    }

                }
                return tablearow;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRow] :  exeption encoutered" + ex.Message);
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// This fucntion is used to identify a row from Flat Table of Infragistics Inc.here we would just select a row from the table
        /// use "uiainfratablerowflatnumber under controltype column in structure sheets.
        /// 
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationFlatInfraTableRowByNumber(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] --> My Current Parent is having name as " + uiAutomationCurrentParent.Current.Name + "auto id as " + uiAutomationCurrentParent.Current.AutomationId.ToString());
            try
            {
                AutomationElement tablearow = null;
                AutomationElementCollection table = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                        new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom));
                logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] -->: row count with controltype.custom within table object count = " + table.Count);
                for (int k = 0; k < Convert.ToInt32(searchValue) + 3; k++)
                {
                    if (k == 0)
                    {
                        logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] -->: k=0");
                        uiAutomationCurrentParent = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                            new System.Windows.Automation.AndCondition(
                                     new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom),
                                     new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, "row 1")
                                     )
                                     );
                        logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] :-> After Assign should be row 1 and is: " + uiAutomationCurrentParent.Current.Name);
                        tablearow = uiAutomationCurrentParent;
                    }
                    else
                    {
                        logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] :-> Current Parent  Before going to  Next Sibling is =" + uiAutomationCurrentParent.Current.Name);
                        logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] :-> Getting sibling k=" + k);
                        uiAutomationCurrentParent = TreeWalker.ControlViewWalker.GetNextSibling(uiAutomationCurrentParent);
                        logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] :-> Current Parent  via Next Sibling is =" + uiAutomationCurrentParent.Current.Name);
                        if (uiAutomationCurrentParent.Current.Name == "row " + Convert.ToInt32(searchValue))
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] :-> Match found with row " + uiAutomationCurrentParent.Current.Name);
                            tablearow = uiAutomationCurrentParent;
                            logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] :-> After Assign should be row " + k + " and is: " + uiAutomationCurrentParent.Current.Name);
                        }

                    }
                }
                return tablearow;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationFlatInfraTableRowByNumber] :  exeption encoutered" + ex.Message);
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// This fucntion is used to identify text block controltype.text
        /// use "uiautomationdialogtextcontrol" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationDialogTextControl(string searchBy, string searchValue, int index)
        {
            AutomationElement dlgtextcontrol = null;

            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "name":
                    case "text":
                        {
                            dlgtextcontrol = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                new System.Windows.Automation.AndCondition(
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Text),
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)
                                     ));
                            return dlgtextcontrol;

                        }
                    case "index":
                        {
                            dlgtextcontrol = GetControlByIndex(System.Windows.Automation.ControlType.Text, index);
                            return dlgtextcontrol;

                        }

                }
                return dlgtextcontrol;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationDialogTextControl]: exception encoutered" + ex.Message);
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        }
        /// <summary>
        /// This fucntion is used to identify combobox control controltype.combobox
        /// use "uiautomationdialogtextcontrol" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationComboBox(string searchBy, string searchValue, int index)
        {
            AutomationElement combocontrol = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();
            logTofile(_eLogPtah, " before try");
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "label":
                        {

                            logTofile(_eLogPtah, "[GetUIAutomationCombobox]: Current Parent Name " + uiAutomationCurrentParent.Current.AutomationId.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));


                            logTofile(_eLogPtah, "Find by label '" + searchValue + "'");

                            logTofile(_eLogPtah, "Finding current root");

                            combocontrol = GetControlByLabel(System.Windows.Automation.ControlType.ComboBox, searchValue);
                            #region Commented moved to function
                            //AutomationElement currentRoot = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants
                            //    , new System.Windows.Automation.AndCondition(new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Text)
                            //    , new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));

                            //var foundComboControl = false;
                            //while (foundComboControl == false)
                            //{

                            //    combocontrol = TreeWalker.ControlViewWalker.GetNextSibling(currentRoot);

                            //    logTofile(_eLogPtah, "Current root is " + combocontrol.Current.System.Windows.Automation.ControlType .ProgrammaticName);
                            //    if (combocontrol.Current.System.Windows.Automation.ControlType  == System.Windows.Automation.ControlType .ComboBox)
                            //    {
                            //        foundComboControl = true;
                            //    }
                            //    else
                            //    {
                            //        currentRoot = combocontrol;
                            //    }
                            //}
                            //logTofile(_eLogPtah, "[GetUIAutomationcombobox]: ------> combobox automation id :was found  =" + combocontrol.Current.AutomationId.ToString()); 
                            #endregion

                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect combobox = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationcombobox]: ------> combobox name :was found  =" + combocontrol.Current.AutomationId.ToString());
                            break;

                        }
                    case "name":
                    case "text":
                        {

                            logTofile(_eLogPtah, "[GetUIAutomationCombobox]: Current Parent Name " + uiAutomationCurrentParent.Current.AutomationId.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));


                            logTofile(_eLogPtah, "Find by label '" + searchValue + "'");

                            logTofile(_eLogPtah, "Finding current root");

                            combocontrol = GetControlByName(System.Windows.Automation.ControlType.ComboBox, searchValue);
                            #region Commented moved to function
                            //AutomationElement currentRoot = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants
                            //    , new System.Windows.Automation.AndCondition(new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Text)
                            //    , new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));

                            //var foundComboControl = false;
                            //while (foundComboControl == false)
                            //{

                            //    combocontrol = TreeWalker.ControlViewWalker.GetNextSibling(currentRoot);

                            //    logTofile(_eLogPtah, "Current root is " + combocontrol.Current.System.Windows.Automation.ControlType .ProgrammaticName);
                            //    if (combocontrol.Current.System.Windows.Automation.ControlType  == System.Windows.Automation.ControlType .ComboBox)
                            //    {
                            //        foundComboControl = true;
                            //    }
                            //    else
                            //    {
                            //        currentRoot = combocontrol;
                            //    }
                            //}
                            //logTofile(_eLogPtah, "[GetUIAutomationcombobox]: ------> combobox automation id :was found  =" + combocontrol.Current.AutomationId.ToString()); 
                            #endregion

                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect combobox = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationcombobox]: ------> combobox name :was found  =" + combocontrol.Current.AutomationId.ToString());
                            break;

                        }

                    case "automationid":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationCombobox]: Current Parent Name " + uiAutomationCurrentParent.Current.AutomationId.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automation id");
                                combocontrol = GetControlByAutomationId(System.Windows.Automation.ControlType.ComboBox, searchValue);
                                logTofile(_eLogPtah, "[GetUIAutomationcombobox]: ------> combobox automation id :was found  =" + combocontrol.Current.AutomationId.ToString());
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                logTofile(_eLogPtah, "Function : GetUIAutomationcombob ------> before combobox search : ");
                                combocontrol = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.ComboBox, searchValue, index);
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect combobox = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationcombobox]: ------> combobox name :was found  =" + combocontrol.Current.AutomationId.ToString());
                            break;

                        }

                    case "index":
                        {
                            logTofile(_eLogPtah, "Find by index");
                            combocontrol = GetControlByIndex(System.Windows.Automation.ControlType.ComboBox, index);

                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationComboBox]:not able to get combobox  ");
                throw new Exception(ex.Message);
            }
            return combocontrol;
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .Pane which is Internet Explorer that has been embedded in a Window
        /// use "uiautomationdialogtextcontrol" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        private AutomationElement GetUIAutomationIEPane(string searchBy, string searchValue, int index)
        {
            try
            {
                AutomationElement uiautomationpane = null;
                if (uiAutomationCurrentParent == null)
                {
                    uiAutomationCurrentParent = uiAutomationWindow;
                }
                else
                {
                    logTofile(_eLogPtah, "[GetUIAutomationIEPane]: No Parent Set parnt before proceeding");
                }
                uiautomationpane = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                            new System.Windows.Automation.AndCondition(new System.Windows.Automation.PropertyCondition(AutomationElement.ClassNameProperty, "Internet Explorer_Server"),
                                                              new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Pane)));


                return uiautomationpane;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationIEPane]:  exeption encoutered" + ex.Message);
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .Table  which is under Internet Explorer Pane that has been embedded in a Window
        /// use "uiautomationdialogtextcontrol" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        private AutomationElement GetUIAutomationIEPaneTable(string searchBy, string searchValue, int index)
        {
            try
            {
                AutomationElement uiautomationiepanetable = null;
                AutomationElement uiautomationpanetbl = GetUIAutomationIEPane("", "", -1);
                if (uiautomationpanetbl != null)
                {
                    logTofile(_eLogPtah, "[GetUIAutomationIEPaneTable]: The Ie pane table was found!");
                }
                AutomationElementCollection tables = uiautomationpanetbl.FindAll(TreeScope.Children,
                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Table));
                uiautomationiepanetable = tables[index];
                return uiautomationiepanetable;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Function -> GetUIAutomationIEPaneTable:  exeption encoutered" + ex.Message);
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .Table  which is Infragistics grid or Table in a Window
        /// use "uiautomationinfratbaleflat" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationInfraTableFlat(string searchBy, string searchValue, int index)
        {
            try
            {
                AutomationElement uiautomationinfratableflat = null;

                uiautomationinfratableflat = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                   new System.Windows.Automation.AndCondition(
                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Table),
                       new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue)));
                uiAutomationCurrentParent = uiautomationinfratableflat;
                return uiautomationinfratableflat;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Function -> GetUIAutomationInfraTableFlat:  exeption encoutered" + ex.Message);
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .SplitButton  which can be a button of a drodown arrow
        /// use "uiautomationsplitter" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationSpillter(string searchBy, string searchValue, int index)
        {
            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationSpillter]: Inside Try.");
                AutomationElement uiautomationsplitter = null;
                AutomationElementCollection splittercol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                        new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.SplitButton
                                            ));
                logTofile(_eLogPtah, "[GetUIAutomationSpillter]: Spliiter button count ." + splittercol.Count.ToString());

                if (splittercol == null)
                {
                    logTofile(_eLogPtah, "[GetUIAutomationSpillter]: Colletion itself is null .Could not Find Object with Given Search conditions in application.");
                }


                if (index <= 0)
                {
                    uiautomationsplitter = splittercol[0];
                }
                else
                {
                    uiautomationsplitter = splittercol[index];
                }


                if (uiautomationsplitter == null)
                {
                    logTofile(_eLogPtah, "[GetUIAutomationSpillter]: control null -->Could not Find Object with Given Search conditions in application.");
                }
                return uiautomationsplitter;
                //break;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationSpillter]: Error." + ex.Message.ToString());
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .Edit  which can be a button of a drodown arrow
        /// use "uiautomationedit" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement
            GetUIAutomationEdit(string searchBy, string searchValue, int index)
        {

            logTofile(_eLogPtah, "[GetUIAutomationEdit]: Outside  Try for CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
            AutomationElement uiautomationedit = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();

            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationEdit]: Inside Try.");

                switch (searchBy.ToLower())
                {
                    case "label":
                        {

                            logTofile(_eLogPtah, "[GetUIAutomationEdit]: Current Parent :" + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            logTofile(_eLogPtah, "Find by label '" + searchValue + "'");
                            logTofile(_eLogPtah, "Finding current root");
                            uiautomationedit = GetControlByLabel(System.Windows.Automation.ControlType.Edit, searchValue);
                            #region Code Commented moved to Function


                            //var foundEditControl = false;
                            //AutomationElement currentRoot;

                            //currentRoot = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants
                            //    , new System.Windows.Automation.AndCondition(new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Text)
                            //    , new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));

                            //logTofile(_eLogPtah, "Current root is " + currentRoot.Current.System.Windows.Automation.ControlType .ToString());

                            //while (foundEditControl == false)
                            //{

                            //    uiautomationedit = TreeWalker.ControlViewWalker.GetNextSibling(currentRoot);

                            //    logTofile(_eLogPtah, "Current root is " + uiautomationedit.Current.System.Windows.Automation.ControlType .ProgrammaticName);
                            //    if (uiautomationedit.Current.System.Windows.Automation.ControlType  == System.Windows.Automation.ControlType .Edit)
                            //    {
                            //        foundEditControl = true;
                            //    }
                            //    else
                            //    {
                            //        currentRoot = uiautomationedit;
                            //    }
                            //} 
                            #endregion
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect edit = " + duration2);

                            break;
                        }
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationEdit]: Current Parent :" + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");
                                uiautomationedit = GetControlByName(System.Windows.Automation.ControlType.Edit, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");

                                AutomationElementCollection editcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                    new System.Windows.Automation.AndCondition(
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Edit),
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));
                                uiautomationedit = editcol[index];
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect edit = " + duration2);

                            break;
                        }
                    case "automationid":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationEdit]: Current Parent :" + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automation id");

                                uiautomationedit = GetControlByAutomationId(System.Windows.Automation.ControlType.Edit, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");


                                uiautomationedit = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.Edit, searchValue, index);
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect edit = " + duration2);

                            break;
                        }

                    case "index":
                        {
                            uiautomationedit = GetControlByIndex(System.Windows.Automation.ControlType.Edit, index);

                            #region Commented code moved to function
                            //AutomationElementCollection editcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                            //    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .Edit));
                            //int ind = Convert.ToInt32(searchValue);
                            //for (int k = 0; k < editcol.Count; k++)
                            //{
                            //    if (ind == k)
                            //    {
                            //        uiautomationedit = editcol[k];
                            //        logTofile(_eLogPtah, " Index found");
                            //        break;
                            //    }
                            //    else
                            //    {
                            //        logTofile(_eLogPtah, " Index not found");
                            //    }

                            //} 
                            #endregion
                        }
                        break;

                }
                return uiautomationedit;
            }


            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationEdit]:Execption in uiautomation edit :" + ex.Message.ToString());
                throw new Exception(ex.Message);
                // return null;
            }

        }
        /// <summary>
        /// This fucntion is used to get all items values from a System.Windows.Automation.ControlType .List control used for Verification of A combobox items 
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public string GetUIAutomationListContent(string searchBy, string searchValue, int index)
        {
            string uilist = null;
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "automationid":
                        {
                            AutomationElement listcontrol = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                new System.Windows.Automation.AndCondition(
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue),
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.List
                                     )));
                            AutomationElementCollection listcontrolcol = listcontrol.FindAll(TreeScope.Children,
                                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.ListItem));
                            if (listcontrolcol == null)
                            {
                                logTofile(_eLogPtah, "[GetUIAutomationListContent]: Could not Find Object with Given Search conditions in application.");

                            }
                            else
                            {

                                for (int i = 0; i < listcontrolcol.Count; i++)
                                {
                                    uilist = uilist + listcontrolcol[i].Current.Name + ";";

                                }

                            }

                        }
                        break;
                    case "name":
                    case "text":
                        {
                            AutomationElement listcontrol = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                new System.Windows.Automation.AndCondition(
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue),
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.List
                                     )));
                            AutomationElementCollection listcontrolcol = listcontrol.FindAll(TreeScope.Children,
                                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.ListItem));
                            if (listcontrolcol == null)
                            {
                                logTofile(_eLogPtah, "[GetUIAutomationListContent]: Could not Find Object with Given Search conditions in application.");

                            }
                            else
                            {

                                for (int i = 0; i < listcontrolcol.Count; i++)
                                {
                                    uilist = uilist + listcontrolcol[i].Current.Name + ";";

                                }

                            }
                        }
                        break;

                }
                return uilist;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationListContent]:Genreic  exeption encoutered");
                Console.WriteLine("Exception " + ex.Message);
                return null;
            }
        }
        public string GetUIAutomationListCollection(string searchBy, string searchValue, int index)
        {
            string uilist = null;
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "automationid":
                        {
                            AutomationElementCollection listcontrolcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                               new System.Windows.Automation.AndCondition(
                                   new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue),
                                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.ListItem)));
                            if (listcontrolcol == null)
                            {
                                logTofile(_eLogPtah, "[GetUIAutomationListCollection]: Could not Find Object with Given Search conditions in application.");
                            }
                            else
                            {
                                for (int i = 0; i < listcontrolcol.Count; i++)
                                {
                                    if (TreeWalker.ControlViewWalker.GetFirstChild(listcontrolcol[i]) != null)
                                    {
                                        AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(listcontrolcol[i]);
                                        string _controltype = elementNode.Current.LocalizedControlType.ToString();
                                        if (_controltype.ToLower() == "text")
                                        {
                                            uilist = uilist + elementNode.Current.Name + ";";
                                        }
                                        else
                                            uilist = uilist + " " + ";";
                                    }
                                    else
                                    {
                                        logTofile(_eLogPtah, "[GetUIAutomationListCollection]:no child ");
                                    }
                                }
                            }
                        }
                        break;
                    case "name":
                    case "text":
                        {
                            AutomationElementCollection listcontrolcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                new System.Windows.Automation.AndCondition(
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue),
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.ListItem)));
                            if (listcontrolcol == null)
                            {
                                logTofile(_eLogPtah, "[GetUIAutomationListCollection]: Could not Find Object with Given Search conditions in application.");
                            }
                            else
                            {
                                for (int i = 0; i < listcontrolcol.Count; i++)
                                {
                                    if (TreeWalker.ControlViewWalker.GetFirstChild(listcontrolcol[i]) != null)
                                    {
                                        AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(listcontrolcol[i]);
                                        string _controltype = elementNode.Current.LocalizedControlType.ToString();
                                        if (_controltype.ToLower() == "text")
                                        {
                                            uilist = uilist + elementNode.Current.Name + ";";
                                        }
                                        else
                                            uilist = uilist + " " + ";";
                                    }
                                    else
                                    {
                                        logTofile(_eLogPtah, "[GetUIAutomationListCollection]:no child ");
                                    }
                                }
                            }
                        }
                        break;

                }
                return uilist;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationListCollection]:Genreic  exeption encoutered");
                Console.WriteLine("Exception " + ex.Message);
                return null;
            }
        }
        public string GetUIAutomationTreeContent(string searchBy, string searchValue, int index)
        {
            string uitrees = null;
            try
            {
                AutomationElement treecontrol = null;
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                    case "name":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationTreeContent: Current Parent :" + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automation id");

                                treecontrol = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                new System.Windows.Automation.AndCondition(
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue),
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.TreeItem
                                     )));
                                AutomationElementCollection treecontrolcol = treecontrol.FindAll(TreeScope.Children,
                                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.TreeItem));
                                if (treecontrol == null)
                                {
                                    logTofile(_eLogPtah, "[GetUIAutomationTreeContent]: Could not Find Object with Given Search conditions in application.");
                                }
                                else
                                {

                                    for (int i = 0; i < treecontrolcol.Count; i++)
                                    {
                                        uitrees = uitrees + treecontrolcol[i].Current.Name + ";";

                                    }

                                }
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                AutomationElementCollection treeitemcontrolcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                new System.Windows.Automation.AndCondition(
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue),
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.TreeItem
                                     )));
                                treecontrol = treeitemcontrolcol[index];
                                AutomationElementCollection treecontrolcol = treecontrol.FindAll(TreeScope.Descendants,
                                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.TreeItem));
                                if (treecontrol == null)
                                {
                                    logTofile(_eLogPtah, "[GetUIAutomationTreeContent]: Could not Find Object with Given Search conditions in application.");
                                }
                                else
                                {

                                    for (int i = 0; i < treecontrolcol.Count; i++)
                                    {
                                        uitrees = uitrees + treecontrolcol[i].Current.Name + ";";

                                    }

                                }
                            }

                        }
                        break;

                }
                return uitrees;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationTreecontent]:Genreic  exeption encoutered");
                Console.WriteLine("Exception " + ex.Message);
                return null;
            }
        }
        /// <summary>
        /// This function is used to retrieve controltype.pane using AutomationID
        /// </summary>
        /// <param name="name">Automation Id of the pane</param>
        public void GetDescenDentPaneWithName(string name)
        {

            if (uiAutomationCurrentParent == null)
            {
                logTofile(_eLogPtah, "Uiautomation curentparent was  null how ?");

            }
            try
            {
                logTofile(_eLogPtah, "[GetDescenDentPaneWithName]: Inside GetPane: ");
                logTofile(_eLogPtah, "[GetDescenDentPaneWithName]: Curret parent name or autoid:  " + uiAutomationCurrentParent.Current.Name);
                AutomationElement paneObj = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
             new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, name, PropertyConditionFlags.IgnoreCase));
                uiAutomationCurrentParent = paneObj;
                logTofile(_eLogPtah, "[GetDescenDentPaneWithName]: Found Pane" + paneObj.Current.AutomationId.ToString() + Environment.NewLine);
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Curretn Parent name  " + uiAutomationCurrentParent.Current.Name + "Currne tAPrent autoid " + uiAutomationCurrentParent.Current.AutomationId);
                logTofile(_eLogPtah, "Traget panes autoamtion id was " + name);
                logTofile(_eLogPtah, "Execption ====> [GetDescenDentPaneWithName]: " + ex.Message + " " + System.DateTime.Now.ToString() + Environment.NewLine);
            }



        }
        /// <summary>
        /// This function is used to retrieve controltype.pane using position or index
        /// </summary>
        /// <param name="position">integer based position </param>
        public void GetChildPane(int position)
        {
            try
            {
                logTofile(_eLogPtah, "[GetChildPane]:Inside GetChildPane: ");
                for (int k = 1; k <= position; k++)
                {
                    logTofile(_eLogPtah, "Count " + k.ToString() + Environment.NewLine);
                    if (k == 1)
                        uiAutomationCurrentParent = TreeWalker.ControlViewWalker.GetFirstChild(uiAutomationCurrentParent);
                    else
                        uiAutomationCurrentParent = TreeWalker.ControlViewWalker.GetNextSibling(uiAutomationCurrentParent);
                }
                logTofile(_eLogPtah, "[GetChildPane]:Found Pane" + uiAutomationCurrentParent.Current.Name.ToString() + Environment.NewLine);
                logTofile(_eLogPtah, "[GetChildPane]:Found Pane" + uiAutomationCurrentParent.Current.AutomationId.ToString() + Environment.NewLine);
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetChildPane]: Exception in  function :GetChildPane " + ex.Message + " " + System.DateTime.Now.ToString() + Environment.NewLine);
            }
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .Menu  
        /// use "uiautomationmenu" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationMenu(string searchBy, string searchValue, int index)
        {
            AutomationElement menucontrol = null;
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "automationid":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationMenu]:trying to to get Menu using autoid");
                            menucontrol = GetControlByAutomationId(System.Windows.Automation.ControlType.Menu, searchValue);
                            break;
                        }
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationMenu]::trying to to get Menu using text ");
                            logTofile(_eLogPtah, "Passing search value as text " + searchValue);
                            logTofile(_eLogPtah, "parent window---" + uiAutomationCurrentParent.Current.Name);

                            menucontrol = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.Menu, searchValue, index);
                            break;
                        }

                }
            }
            catch
            {
                logTofile(_eLogPtah, "[GetUIAutomationMenu]:not able to get Menu ");
            }
            return menucontrol;



        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .MenuItem 
        /// use "uiautomationmenuitem" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationMenuItem(string searchBy, string searchValue, int index)
        {
            AutomationElement menuitemcontrol = null;
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "automationid":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationMenuItem]:trying to to get Menu using autoid");
                            menuitemcontrol = GetControlByAutomationId(System.Windows.Automation.ControlType.MenuItem, searchValue);
                            break;
                        }

                    case "helptext":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationMenuItem]::trying to to get Menu using helptext ");
                            logTofile(_eLogPtah, "[GetUIAutomationMenuItem]:Curent parent is :" + uiAutomationCurrentParent.Current.Name.ToString());
                            menuitemcontrol = GetControlByHelpText(System.Windows.Automation.ControlType.MenuItem, searchValue);
                            break;
                        }

                    case "text":
                    case "name":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationMenuItemu]::trying to to get Menu using text ");
                            logTofile(_eLogPtah, "[GetUIAutomationMenuItemu]:Curent parent is :" + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");

                                menuitemcontrol = GetControlByName(System.Windows.Automation.ControlType.MenuItem, searchValue);

                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");

                                menuitemcontrol = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.MenuItem, searchValue, index);
                            }
                            break;
                        }
                }
                return menuitemcontrol;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationMenuItem]:not able to get Menu " + ex.Message);
                return null;
            }


        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .Thumb
        /// use "uiautomationthumb" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationThumb(string searchBy, string searchValue, int index)
        {
            AutomationElement thumb = null;
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "Function : GetUIAutomationthumb------> before buttons search : ");
                            logTofile(_eLogPtah, "[GetUIAutomationthumb]:------>  " + uiAutomationCurrentParent.Current.AutomationId + "text" + uiAutomationCurrentParent.Current.Name);
                            AutomationElementCollection thumbcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Thumb));
                            int j = 0;
                            logTofile(_eLogPtah, "[GetUIAutomationbutton]: ------> Buttons present on window(CurrentParent) : " + uiAutomationCurrentParent.Current.Name + "= " + thumbcol.Count);
                            for (int i = 0; i < thumbcol.Count; i++)
                            {
                                if (thumbcol[i].Current.Name.ToString() == searchValue)
                                {
                                    if (index <= 0)
                                    {
                                        thumb = thumbcol[i];
                                        break;
                                    }
                                    else
                                    {
                                        if (j == index)
                                            thumb = thumbcol[i];
                                    }
                                    j++;
                                }
                            }
                            logTofile(_eLogPtah, "[GetUIAutomationthumb: ------> thumb name :was found  =" + thumb.Current.Name.ToString());
                            break;
                        }

                    case "helptext":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationthumb]:------>  " + uiAutomationCurrentParent.Current.HelpText + "Helptext" + uiAutomationCurrentParent.Current.Name);
                            AutomationElementCollection thumbcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Thumb));
                            int j = 0;
                            logTofile(_eLogPtah, "[GetUIAutomationthumb]]: ------> Buttons present on window(CurrentParent) : " + uiAutomationCurrentParent.Current.Name + "= " + thumbcol.Count);
                            for (int i = 0; i < thumbcol.Count; i++)
                            {

                                if (thumbcol[i].Current.HelpText.ToString() == searchValue)
                                {
                                    // logTofile(_eLogPtah, "[GetUIAutomationbutton]: ------> Button name : =" + buttoncol[i].Current.Name.ToString());
                                    if (index <= 0)
                                    {
                                        thumb = thumbcol[i];

                                        break;
                                    }
                                    else
                                    {
                                        if (j == index)
                                            thumb = thumbcol[i];
                                    }
                                    j++;
                                }
                            }
                            logTofile(_eLogPtah, "[GetUIAutomationthumb]]: ------> Button name :was found  =" + thumb.Current.HelpText.ToString());
                            break;
                        }
                    case "index":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationthumb]:------> Using index criteria  ");
                            AutomationElementCollection thumbcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Thumb));
                            logTofile(_eLogPtah, "[GetUIAutomationthumb]]: ------> Thumbs present on window(CurrentParent) : " + uiAutomationCurrentParent.Current.Name + "= " + thumbcol.Count);
                            for (int i = 0; i < thumbcol.Count; i++)
                            {
                                if (i == index)
                                {
                                    thumb = thumbcol[i];
                                    if (thumb != null)
                                    {
                                        logTofile(_eLogPtah, "[GetUIAutomationthumb]]Got automation id as : " + thumb.Current.AutomationId.ToString());
                                    }
                                    else
                                    {
                                        logTofile(_eLogPtah, "[GetUIAutomationthumb]] no thumb control obtained ");
                                    }
                                }
                            }
                            logTofile(_eLogPtah, "[GetUIAutomationthumb]]: ------> Thumb name :was found  =" + thumb.Current.AutomationId.ToString());
                            break;
                        }

                    case "appspecificindex":
                        {
                            AutomationElementCollection dashboardlist = uiAutomationCurrentParent.FindAll(TreeScope.Children,
                             new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Thumb));
                            break;
                        }
                }//switch
            }
            catch
            {
                logTofile(_eLogPtah, "Excepetion : In Thumb element fucntion");
            }
            return thumb;
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .Text This is used to click on a text control block
        /// use "uiautomationtext" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationText(string searchBy, string searchValue, int index)
        {
            AutomationElement text = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                    case "name":
                        {
                            logTofile(_eLogPtah, "Function : GetUIAutomationtext ------> before buttons search : ");
                            logTofile(_eLogPtah, "[GetUIAutomationtext]:------>  " + uiAutomationCurrentParent.Current.AutomationId + "text" + uiAutomationCurrentParent.Current.Name);
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");

                                text = GetControlByName(System.Windows.Automation.ControlType.Text, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by name and index :");
                                text = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.Text, searchValue, index);

                            }

                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect text = " + duration2);

                            break;
                        }
                    case "automationid":
                        {
                            logTofile(_eLogPtah, "Function : GetUIAutomationTExt ------> before TExt search : ");
                            logTofile(_eLogPtah, "[GetUIAutomationtext]:------>  " + uiAutomationCurrentParent.Current.AutomationId + "text" + uiAutomationCurrentParent.Current.Name);
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automationid");
                                text = GetControlByAutomationId(System.Windows.Automation.ControlType.Text, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by automation id and index :");
                                text = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.Text, searchValue, index);
                            }

                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect text = " + duration2);

                            break;

                        }

                }//Switch

            }
            catch
            {
                logTofile(_eLogPtah, "Error in GetUIAutomtionText");
            }
            return text;
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .Image This is used to click on a Image
        /// use "uiautomationimage" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationImage(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "[GetUIAutomationimage]: Outside  Try for CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
            AutomationElement image = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "automationid":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationimage]: Current Parent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automation id");

                                image = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                   new System.Windows.Automation.AndCondition(
                                         new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Image),
                                     new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue)));
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                logTofile(_eLogPtah, "Function : GetUIAutomationimage ------> before image search : ");

                                AutomationElementCollection imagecol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                    new System.Windows.Automation.AndCondition(
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Image),
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue)));
                                logTofile(_eLogPtah, "Count: " + imagecol.Count);
                                image = imagecol[index];
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect image = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationimage]: ------> Image automation id :was found  =" + image.Current.AutomationId.ToString());
                            break;


                        }
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationimage]: Current Parent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");

                                image = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                   new System.Windows.Automation.AndCondition(
                                         new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Image),
                                     new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                logTofile(_eLogPtah, "Function : GetUIAutomationimage ------> before image search : ");

                                AutomationElementCollection imagecol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                    new System.Windows.Automation.AndCondition(
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Image),
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));
                                image = imagecol[index];
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect image = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationimage]: ------> Image name :was found  =" + image.Current.Name.ToString());
                            break;
                        }

                    case "helptext":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationimage]:------>  " + uiAutomationCurrentParent.Current.HelpText + "Helptext" + uiAutomationCurrentParent.Current.Name);
                            AutomationElementCollection imagecol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Image));
                            int j = 0;
                            logTofile(_eLogPtah, "[GetUIAutomationimage]: ------> Buttons present on window(CurrentParent) : " + uiAutomationCurrentParent.Current.Name + "= " + imagecol.Count);
                            for (int i = 0; i < imagecol.Count; i++)
                            {

                                if (imagecol[i].Current.HelpText.ToString() == searchValue)
                                {
                                    // logTofile(_eLogPtah, "[GetUIAutomationbutton]: ------> Button name : =" + buttoncol[i].Current.Name.ToString());
                                    if (index <= 0)
                                    {
                                        image = imagecol[i];
                                        break;
                                    }
                                    else
                                    {
                                        if (j == index)
                                            image = imagecol[i];
                                    }
                                    j++;
                                }
                            }
                            logTofile(_eLogPtah, "[GetUIAutomationImage]: ------> Image name :was found  =" + image.Current.HelpText.ToString());
                            break;
                        }
                }//switdh
            }
            catch
            {
                logTofile(_eLogPtah, "[GetUIAutomationImage]:Error in GetUIAutomationImage");
            }
            return image;
        }
        public AutomationElement GetUIAutomationDataItem(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "inside getuiautomationdataitem");
            AutomationElement dtitem = null;
            try
            {
                logTofile(_eLogPtah, "inside try");
                switch (searchBy.ToLower())
                {
                    case "index":
                        {
                            logTofile(_eLogPtah, "inside case index");
                            AutomationElementCollection dataitemcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataItem));
                            logTofile(_eLogPtah, " Count of dataitems: " + dataitemcol.Count);
                            int k = Convert.ToInt32(searchValue);
                            logTofile(_eLogPtah, "index of dataitem: " + k.ToString());
                            //AutomationElement dtitem = null;
                            for (int i = 0; i < dataitemcol.Count; i++)
                            {
                                logTofile(_eLogPtah, "iteration: " + i.ToString());
                                if (i == k)
                                {
                                    logTofile(_eLogPtah, "Match found");
                                    dtitem = dataitemcol[i];
                                    logTofile(_eLogPtah, " Name of dataitem is: " + dtitem.Current.Name);
                                    break;
                                }

                            }
                            break;
                        }
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationdataitem]: Current Parent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");

                                dtitem = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                   new System.Windows.Automation.AndCondition(
                                         new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataItem),
                                     new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                logTofile(_eLogPtah, "Function : GetUIAutomationdataitem ------> before search ");

                                AutomationElementCollection dataitemcol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                    new System.Windows.Automation.AndCondition(
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataItem),
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));
                                dtitem = dataitemcol[index];

                            }

                            logTofile(_eLogPtah, "[GetUIAutomationdataitem]: ------> dataitem name :was found  =" + dtitem.Current.Name.ToString());


                            break;
                        }

                }
                logTofile(_eLogPtah, "Returned dataitem is :" + dtitem.Current.Name);
                return dtitem;
            }
            catch (Exception e)
            {
                logTofile(_eLogPtah, "[GetUIAutomationDataItem]:Exception was encoutered with Dataitem" + e.Message.ToString());
                throw new Exception(e.Message);
                // return dtitem;
            }
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .datagridThis is used to click on a Image
        /// use "uiautomationidatagrid" under controltype column in structure sheets.
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationDataGrid(string searchBy, string searchValue, int index)
        {
            AutomationElement uiautomationdatagrid = null;

            try
            {
                switch (searchBy.ToLower())
                {
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationDataGrid] Saerching by TExt");
                            logTofile(_eLogPtah, "[GetUIAutomationDataGrid] serarchvalue is " + searchValue);
                            logTofile(_eLogPtah, "[GetUIAutomationDataGrid] Saerching by TExt");
                            logTofile(_eLogPtah, "[GetUIAutomationDataGrid] Current paren tis " + uiAutomationCurrentParent.Current.Name);
                            uiautomationdatagrid = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                   new System.Windows.Automation.AndCondition(
                       new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue),
                       new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataGrid)));
                            break;
                        }

                    case "automationid":
                        {
                            uiautomationdatagrid = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                            new System.Windows.Automation.AndCondition(
                                new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, searchValue),
                                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataGrid)));
                            break;

                        }
                    case "index":
                        {
                            uiautomationdatagrid = GetControlByIndex(System.Windows.Automation.ControlType.DataGrid, index);
                            logTofile(_eLogPtah, "Grid found by Index");
                            break;

                        }

                    default:
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationDataGrid]: Invalid SearchBy criteria Valid are only text,Index and automationid");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationDataGrid]:Exception was encoutered with Datagrid" + ex.Message.ToString());
                throw new Exception(ex.Message);
            }

            return uiautomationdatagrid;
        }
        /// <summary>
        /// This fucntion is used to identify System.Windows.Automation.ControlType .CheckBox. 
        /// </summary>
        /// <param name="searchBy">Should be either automationid, text(Name in UI spy),helptext(tooltip over object)  or index</param>
        /// <param name="searchValue">Value for the automationid, text(Name in UI spy) or index passed in searchby parameter.</param>
        /// <param name="index">Ordinal identifier when test object or automation element does not have automationid text, or helptext</param>
        /// <returns></returns>
        public AutomationElement GetUIAutomationCheckBox(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "[GetUIAutomationCheckBox]: Outside  Try for CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
            AutomationElement uiAutomationCheckBox = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();
            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationCheckBox]: Inside Try for  CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                    case "name":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationCheckBox]: Current Parent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");

                                uiAutomationCheckBox = GetControlByName(System.Windows.Automation.ControlType.CheckBox, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by name index");

                                uiAutomationCheckBox = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.CheckBox, searchValue, index);
                            }

                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect checkbox = " + duration2);
                            break;
                        }
                    case "automationid":
                        {
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automation id");

                                logTofile(_eLogPtah, "[GetUIAutomationCheckBox]: CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                                logTofile(_eLogPtah, "[GetUIAutomationCheckBox]: Current automationid " + uiAutomationCurrentParent.Current.Name.ToString());
                                uiAutomationCheckBox = GetControlByAutomationId(System.Windows.Automation.ControlType.CheckBox, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                uiAutomationCheckBox = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.CheckBox, searchValue, index);
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect checkbox = " + duration2);
                            break;

                        }

                    case "index":
                        {
                            uiAutomationCheckBox = GetControlByIndex(System.Windows.Automation.ControlType.CheckBox, index);
                        }
                        break;

                }
                return uiAutomationCheckBox;


            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationCheckBox ]: Generic exeption encoutered");
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
                //  return null;
            }

        } //fucntion end 

        public AutomationElement GetUIAutomationListItem(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "[GetUIAutomationListItem]: Outside  Try for CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
            AutomationElement uiAutomationListItem = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();
            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationListItem]: Inside Try for  CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                switch (searchBy.Trim().ToLower())
                {
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationListItem]: Current Parent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");

                                uiAutomationListItem = GetControlByName(System.Windows.Automation.ControlType.ListItem, searchValue);
                                logTofile(_eLogPtah, "Listitem :" + uiAutomationListItem.Current.Name.ToString());
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                uiAutomationListItem = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.ListItem, searchValue, index);

                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect listitem = " + duration2);
                            break;

                        }
                    case "automationid":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationListItem]: CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "[GetUIAutomationListItem]: Current automationid " + uiAutomationCurrentParent.Current.Name.ToString());
                            if ((index == -1) == true)
                            {
                                uiAutomationListItem = GetControlByAutomationId(System.Windows.Automation.ControlType.ListItem, searchValue);
                            }
                            else
                            {
                                uiAutomationListItem = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.ListItem, searchValue, index);
                            }
                            break;

                        }

                    case "index":
                        {
                            uiAutomationListItem = GetControlByIndex(System.Windows.Automation.ControlType.ListItem, index);
                            break;
                        }
                }
                return uiAutomationListItem;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationListItem ]: Generic exeption encoutered");
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
                // return null;
            }

        } //fucntion end 

        public AutomationElement GetUIAutomationRadioButton(string searchBy, string searchValue, int index)
        {
            AutomationElement radiocontrol = null;
            //    AutomationElementCollection radiocollection = null;
            try
            {
                switch (searchBy.Trim().ToLower())
                {
                    case "automationid":
                        {
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "[GetUIAutomationRadioButton]:passing automation id as  " + searchValue);
                                radiocontrol = GetControlByAutomationId(System.Windows.Automation.ControlType.RadioButton, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "[GetUIAutomationRadioButton]:passing automation id and index   " + searchValue + "index " + index);
                                radiocontrol = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.RadioButton, searchValue, index);
                            }
                            break;
                        }
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationRadioButton]:passing text as  " + searchValue);
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "[GetUIAutomationRadioButton]:passing automation id as  " + searchValue);
                                radiocontrol = GetControlByName(System.Windows.Automation.ControlType.RadioButton, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "[GetUIAutomationRadioButton]:passing automation id and index   " + searchValue + "index " + index);
                                radiocontrol = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.RadioButton, searchValue, index);
                            }
                            break;
                        }

                    case "index":
                        {

                            radiocontrol = GetControlByIndex(System.Windows.Automation.ControlType.RadioButton, index);

                            #region Commented call moved to Common Function
                            //radiocollection = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                            //    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType .RadioButton));

                            //int ind = Convert.ToInt32(searchValue);
                            //for (int k = 0; k < radiocollection.Count; k++)
                            //{
                            //    if (ind == k)
                            //    {
                            //        radiocontrol = radiocollection[k];
                            //        logTofile(_eLogPtah, " Index found");
                            //        break;
                            //    }
                            //    else
                            //    {
                            //        logTofile(_eLogPtah, " Index not found");
                            //    }

                            //} 
                            #endregion

                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationRadio]:not able to get Radio Button  ");
                throw new Exception(ex.Message);
            }
            return radiocontrol;

        }
        public AutomationElement GetUIAutomationHeader(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "[GetUIAutomationHeader]: Outside  Try for CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
            AutomationElement uiAutomationHeader = null;
            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationHeader]: Inside Try for  CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                switch (searchBy.Trim().ToLower())
                {
                    case "name":
                    case "text":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationHeader]: Current Parent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            AutomationElementCollection headerCol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Header));
                            int j = 0;

                            logTofile(_eLogPtah, "[GetUIAutomationHeader]: header count " + headerCol.Count.ToString());
                            for (int i = 0; i < headerCol.Count; i++)
                            {

                                if (headerCol[i].Current.Name == searchValue)
                                {
                                    logTofile(_eLogPtah, "[GetUIAutomationHeader]: searching for:  " + searchValue + "   obtained:   " + headerCol[i].Current.Name);
                                    if (index <= 0)
                                    {
                                        uiAutomationHeader = headerCol[i];
                                        break;
                                    }
                                    else
                                    {
                                        if (j == index)
                                            uiAutomationHeader = headerCol[i];
                                    }
                                    j++;
                                }
                            }

                            return uiAutomationHeader;
                            //break;
                        }
                    case "automationid":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationHeader]: CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "[GetUIAutomationHeader]: Current automationid " + uiAutomationCurrentParent.Current.Name.ToString());
                            AutomationElementCollection headerCol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Header));
                            logTofile(_eLogPtah, "[GetUIAutomationHeader]: Checkbox count " + headerCol.Count.ToString());

                            int j = 0;
                            for (int i = 0; i < headerCol.Count; i++)
                            {

                                if (headerCol[i].Current.AutomationId == searchValue)
                                {
                                    logTofile(_eLogPtah, "[GetUIAutomationHeader: searching for:  " + searchValue + "   obtained:   " + headerCol[i].Current.AutomationId);
                                    if (index <= 0)
                                    {
                                        uiAutomationHeader = headerCol[i];
                                        break;
                                    }
                                    else
                                    {
                                        if (j == index)
                                            uiAutomationHeader = headerCol[i];
                                    }
                                    j++;
                                }
                            }
                            return uiAutomationHeader;

                        }

                    case "index":
                        {
                            AutomationElementCollection headerCol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                 new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Header));

                            int ind = Convert.ToInt32(searchValue);
                            for (int k = 0; k < headerCol.Count; k++)
                            {
                                if (ind == k)
                                {
                                    uiAutomationHeader = headerCol[k];
                                    logTofile(_eLogPtah, " Index found");
                                    break;
                                }
                                else
                                {
                                    logTofile(_eLogPtah, " Index not found");
                                }

                            }

                        }
                        break;

                }
                return uiAutomationHeader;


            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationHeader ]: Generic exeption encoutered");
                logTofile(_eLogPtah, "Erroring Line number in [GetUIAutomationHeader ] " + GetStacktrace(ex).ToString());
                Console.WriteLine("Exception " + ex.Message);


                throw new Exception(ex.Message);

            }


        }

        public AutomationElement GetUIAutomationHeaderitem(string searchBy, string searchValue, int index)
        {
            logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]: Outside  Try for CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
            AutomationElement headeritem = null;
            System.Diagnostics.Stopwatch stopwatch3 = new System.Diagnostics.Stopwatch();
            stopwatch3.Start();

            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]: Inside Try for  CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                    case "name":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]: Current Parent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by name");
                                headeritem = GetControlByName(System.Windows.Automation.ControlType.HeaderItem, searchValue);
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                logTofile(_eLogPtah, "Function : GetUIAutomationHeaderitem ------> before headeritems search : ");
                                headeritem = GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType.HeaderItem, searchValue, index);
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect headeritem = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]: ------> headeritem name :was found  =" + headeritem.Current.Name.ToString());
                            break;
                        }

                    case "helptext":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]:------>  " + uiAutomationCurrentParent.Current.HelpText + "Helptext" + uiAutomationCurrentParent.Current.Name);
                            headeritem = GetControlByHelpText(System.Windows.Automation.ControlType.HeaderItem, searchValue);
                            logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]: ------> headeritem helptext :was found  =" + headeritem.Current.HelpText.ToString());
                            break;
                        }

                    case "automationid":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]: Current Parent Name " + uiAutomationCurrentParent.Current.AutomationId.ToString());
                            logTofile(_eLogPtah, "index  = " + index);
                            logTofile(_eLogPtah, "index bool state indx= -1  " + (index == -1));
                            if ((index == -1) == true)
                            {
                                logTofile(_eLogPtah, "Find by automation id");
                                headeritem = GetControlByAutomationId(System.Windows.Automation.ControlType.HeaderItem, searchValue);
                                logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]: ------> headeritem automation id :was found  =" + headeritem.Current.AutomationId.ToString());
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Find by index");
                                logTofile(_eLogPtah, "Function : GetUIAutomationHeaderitem ------> before headeritem search : ");
                                headeritem = GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType.HeaderItem, searchValue, index);
                            }
                            stopwatch3.Stop();
                            long duration2 = stopwatch3.ElapsedMilliseconds;
                            logTofile(_eLogPtah, "Total time  in mSecs to detect headeritem = " + duration2);
                            logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]: ------> headeritem name :was found  =" + headeritem.Current.AutomationId.ToString());
                            break;
                        }

                    case "index":
                        {
                            logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]: ------> call by Index");
                            headeritem = GetControlByIndex(System.Windows.Automation.ControlType.HeaderItem, index);
                            break;

                        }

                }
                return headeritem;
            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetUIAutomationHeaderitem]:  exception encoutered" + ex.Message);
                Console.WriteLine("Exception " + ex.Message);
                throw new Exception(ex.Message);
            }

        }

        public AutomationElement GetUIAutomationPVMaskEdit(string searchBy, string searchValue, int index)
        {
            AutomationElement maskedit = null;
            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationPvMaskEdit]: Inside Try for  CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                switch (searchBy.Trim().ToLower())
                {
                    case "classname":
                        {
                            maskedit = GetControlByClassNameandIndex(searchValue, index);
                            break;
                        }
                }
            }
            catch
            {

            }
            return maskedit;
        }

        public AutomationElement GetUIAutomationPVComboBox(string searchBy, string searchValue, int index)
        {
            AutomationElement maskedit = null;
            try
            {
                logTofile(_eLogPtah, "[GetUIAutomationPvMaskEdit]: Inside Try for  CurrentParent Name " + uiAutomationCurrentParent.Current.Name.ToString());
                switch (searchBy.Trim().ToLower())
                {
                    case "classname":
                        {
                            maskedit = GetControlByClassNameandIndex(searchValue, index);
                            break;
                        }
                }
            }
            catch
            {

            }
            return maskedit;
        }
        // *****************Object Identification Functions for UIautomation *************************************************
        #endregion




        //this function is called inside GetWPFWindow() and its overloaded method
        #region AddData
        private void AddData(int rowPosition)
        {
            string parentType = "";
            string parentSearchBy = "";
            string parentSearchValue = "";
            string controlaction = "";
            string section = "";
            var _controlType = "";
            var _logicalName = "";
            var _controlName1 = "{Right}";
            try
            {
                #region recordsinexcel
                string _controlValue = null;
                section = testData.Structure.Rows[0]["Section"].ToString();
                // When to Create Data table for first time ?
                //   if (System.IO.File.Exists(_reportsPath + uiAfileName + "Log.csv") == false)
                //   {
                //    uilog.AddHeaders();
                //   }
                for (int i = 0; i < testData.Structure.Rows.Count; i++)
                {
                    uilog.AddHeaders();
                    uilog.createnewrow();
                    parentType = testData.Structure.Rows[i]["ParentType"].ToString();
                    parentSearchBy = testData.Structure.Rows[i]["ParentSearchBy"].ToString().ToLower();
                    parentSearchValue = testData.Structure.Rows[i]["ParentSearchValue"].ToString();
                    controlaction = testData.Structure.Rows[i]["ParentSearchValue"].ToString();
                    uilog.AddTexttoColumn("FunctionName", "AddData");
                    uilog.AddTexttoColumn("StructureSheetName", ptestDataPath);
                    uilog.AddTexttoColumn("TestCaseID", ptestCase);
                    uilog.AddTexttoColumn("ParentType", parentType);
                    uilog.AddTexttoColumn("ParentSearchBy", parentSearchBy);
                    uilog.AddTexttoColumn("ParentSearchValue", parentSearchValue);

                    if (Convert.IsDBNull(testData.Structure.Rows[i]["FieldName"]) == false)
                    {
                        _logicalName = (string)testData.Structure.Rows[i]["FieldName"].ToString();
                        logTofile(_eLogPtah, "Read  FieldName/Logical Name : " + _logicalName);
                        uilog.AddTexttoColumn("FieldName", _logicalName);
                    }
                    if (_logicalName.Length > 0)
                    {
                        _controlValue = (string)testData.Data.Rows[0][_logicalName].ToString();
                        logTofile(_eLogPtah, "Logical Name: " + _logicalName + " Input Value :  " + _controlValue);
                        uilog.AddTexttoColumn("ControlValue", _controlValue);
                    }


                    logTofile(_eLogPtah, "******************* Section (Screen Name ) : " + section + "*************************************");

                    if ((string)testData.Structure.Rows[i]["inputdata"].ToString().ToLower() == "y")
                    {

                        if (Convert.IsDBNull(testData.Structure.Rows[i]["ParentType"]) == false)
                        {
                            if (!String.IsNullOrEmpty(_controlValue) || String.IsNullOrEmpty(_controlType))
                            {

                                #region ConstructParent
                                switch (parentType.Trim().ToLower())
                                {
                                    case "window":
                                        {
                                            try
                                            {
                                                /* wpfapp._application = _application;
                                                 _globalWindow = wpfapp.GetWPFWindow(parentSearchValue);
                                                 _globalWindow.Click();
                                                 * */
                                            }
                                            catch (Exception ex)
                                            {
                                                logTofile(_eLogPtah, "[AddData]->[Window]: Execption was encountered" + ex.Message.ToString());
                                            }
                                            break;
                                        }
                                    /*  case "groupbox":
                                          {
                                              if (_immediateParent.Trim().ToLower() == "window")
                                                  wpfapp.GetWPFGroupBox(_globalWindow, parentSearchBy, parentSearchValue);
                                              else
                                                  wpfapp.GetWPFGroupBox(_globalGroup, parentSearchBy, parentSearchValue);
                                              break; 
                                          }
                                      case "wpfmenu":
                                          {
                                              _globalMenu = wpfapp.GetWPFMenu(_globalWindow, parentSearchBy, parentSearchValue);
                                              _globalMenu.Click();
                                              break; 
                                          } */
                                    case "uiautomationwindow":
                                    case "uwindow":
                                        {

                                            //  if (uiAutomationWindow == null || uiAutomationWindow.Current.Name != parentSearchValue)

                                            if (uiAutomationWindow == null)
                                            {
                                                uiAutomationWindow = GetUIAutomationWindow(parentSearchBy, parentSearchValue);

                                            }
                                            else
                                            {
                                                uiAutomationCurrentParent = uiAutomationWindow;
                                                logTofile(_eLogPtah, "[AddData]->[uiautomationwindow]->UI automationwindow was already set " + uiAutomationWindow.Current.Name.ToString());
                                            }
                                            break;


                                        }

                                    case "uiautomationchildwindow":
                                    case "uchildwindow":
                                        {


                                            uiAutomationCurrentParent = GetUIAutomationWindow(parentSearchBy, parentSearchValue);
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationchildwindow]->UI automationchild window was already set " + uiAutomationCurrentParent.Current.Name.ToString());

                                            break;


                                        }
                                    //  case "uiautomationpane":
                                    //    {
                                    //        uiAutomationCurrentParent = GetUIAutomationPane(parentSearchBy, parentSearchValue);
                                    //        break;
                                    //    }
                                    case "commonhieararchy":
                                        {
                                            string parentH = "";
                                            string parentHvalue = "";
                                            string parentHSearchby = "";
                                            testDataHieararchy.GetTestData(hrchyfile, parentSearchValue);
                                            # region HiearachySheet
                                            for (int ih = 0; ih < testDataHieararchy.Data.Rows.Count; ih++)
                                            {
                                                parentH = testDataHieararchy.Data.Rows[ih]["Parent"].ToString();
                                                parentHvalue = testDataHieararchy.Data.Rows[ih]["Value"].ToString();
                                                parentHSearchby = testDataHieararchy.Data.Rows[ih]["HSearchBy"].ToString();
                                                switch (parentH.ToString().ToLower())
                                                {
                                                    case "uiautomationwindow":
                                                        {
                                                            if (uiAutomationWindow == null)
                                                            {
                                                                uiAutomationWindow = GetUIAutomationWindow(parentHSearchby, parentHvalue);
                                                            }
                                                            else
                                                            {

                                                                uiAutomationCurrentParent = uiAutomationWindow;
                                                                logTofile(_eLogPtah, "[AddData]->[commonhieararchy]->[uiautomationwindow] : loaded with UIautomaiton window Value");
                                                                logTofile(_eLogPtah, "[AddData]->[commonhieararchy]->[uiautomationwindow] :UI automationwindow was already set and hehce Will not be Reloaded unless you force by some means:COMMONH");
                                                            }
                                                            break;
                                                        }
                                                    case "uiautomationpane":
                                                        {
                                                            GetDescenDentPaneWithName(parentHvalue);
                                                            break;
                                                        }
                                                    // this is used to junp to nth pane when no automation id or text was avaialable for panes.
                                                    case "uiautomationchildpane":
                                                        {
                                                            GetChildPane(Int32.Parse(parentHvalue));
                                                            break;
                                                        }
                                                }
                                            }
                                            # endregion HiearachySheet
                                            break;
                                        }
                                    case "uiautomationpane":
                                    case "upane":
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                GetDescenDentPaneWithName(parentSearchValue);
                                            }
                                            break;
                                        }
                                    case "uiautomationchildpane":
                                    case "uchildpane":
                                        {
                                            GetChildPane(Int32.Parse(parentSearchValue));
                                            break;
                                        }

                                    case "uiautomationtreeitem":
                                    case "utreeitem":
                                        {
                                            uiAutomationCurrentParent = GetUIAutomationtreeitem(parentSearchBy, parentSearchValue, -1);
                                            break;
                                        }

                                    default:
                                        throw new Exception("[AddData]:Not a valid parent type.");
                                }
                                #endregion ConstructParent
                            }
                            else
                            {
                                logTofile(_eLogPtah, "Parent is not constructed as control value is null");
                            }
                        }

                        var _action = "";
                        var _searchBy = "";
                        var _index = -1;
                        var _controlName = "";
                        if (Convert.IsDBNull(testData.Structure.Rows[i]["ControlName"]) == false)
                        {
                            _controlName = (string)testData.Structure.Rows[i]["ControlName"];
                            uilog.AddTexttoColumn("ControlName", _controlName);
                        }
                        logTofile(_eLogPtah, "SearchValue of control has been read from datatable " + _controlName);
                        try
                        {
                            if (Convert.IsDBNull(testData.Structure.Rows[i]["FrenchValue"]) == false)
                            {
                                _controlName1 = (string)testData.Structure.Rows[i]["FrenchValue"].ToString();
                                logTofile(_eLogPtah, "Read  FrenchValue : " + _controlName1);
                            }
                        }
                        catch
                        {
                        }

                        if (Convert.IsDBNull(testData.Structure.Rows[i]["Action"]) == false)
                        {
                            _action = (string)testData.Structure.Rows[i]["Action"];
                            logTofile(_eLogPtah, "Action:" + _action);
                            uilog.AddTexttoColumn("Action", _action);
                            #region ActionColumn
                            switch (_action.Trim().ToLower())
                            {
                                case "drag":
                                    logTofile(_eLogPtah, "[Using drag action] ");
                                    string[] coord = _controlValue.Split(';');
                                    string startCoordinate = coord[0];
                                    string endCoordinate = coord[1];
                                    clsCUIT_app app = new clsCUIT_app();
                                    app._eLogPtah = _eLogPtah;
                                    app.Drag(startCoordinate, endCoordinate);
                                    uilog.AddTexttoColumn("Action Performed on Control", "Drag:");
                                    break;

                                case "autoitdrag":
                                    logTofile(_eLogPtah, "[Using autoit drag action]");
                                    string[] coordinates = _controlValue.Split(';');
                                    string startcoord = coordinates[0];
                                    string endcoord = coordinates[1];
                                    string[] start = startcoord.Split(',');
                                    string[] end = endcoord.Split(',');
                                    int x1 = Convert.ToInt32(start[0]);
                                    int y1 = Convert.ToInt32(start[1]);
                                    int x2 = Convert.ToInt32(end[0]);
                                    int y2 = Convert.ToInt32(end[1]);
                                    at.MouseClickDrag("Left", x1, y1, x2, y2, 10);
                                    uilog.AddTexttoColumn("Action Performed on Control", "AutoIt Drag");
                                    break;

                                case "keyboard":
                                    logTofile(_eLogPtah, "[Using Keyboard searching] " + _controlName);


                                    System.Windows.Forms.SendKeys.Flush();
                                    logTofile(_eLogPtah, "Waiting for : " + _controlName);
                                    System.Windows.Forms.SendKeys.SendWait(_controlName);
                                    uilog.AddTexttoColumn("Action Performed on Control", "Sent Keystroke: Unconditional" + _controlName);
                                    break;

                                case "condkeyboard":
                                    logTofile(_eLogPtah, "[Checking if Conditional Keyboard is to be used]");
                                    if (_controlValue.Length > 0)
                                    {
                                        logTofile(_eLogPtah, "[Using Conditional Keyboard] " + _controlName);
                                        if (Convert.IsDBNull(testData.Structure.Rows[i]["Index"]) == false)
                                        {
                                            _index = Convert.ToInt32(testData.Structure.Rows[i]["Index"]);
                                            logTofile(_eLogPtah, " value of index " + _index);
                                            for (int j = 0; j < _index; j++)
                                            {
                                                System.Windows.Forms.SendKeys.Flush();
                                                Console.WriteLine("Waiting for : " + _controlName);
                                                System.Windows.Forms.SendKeys.SendWait(_controlName);
                                                uilog.AddTexttoColumn("Action Performed on Control", "Sent Keystroke: conditional using index value" + _controlName);
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.Forms.SendKeys.Flush();
                                            Console.WriteLine("Waiting for : " + _controlName);
                                            System.Windows.Forms.SendKeys.SendWait(_controlName);
                                            uilog.AddTexttoColumn("Action Performed on Control", "Sent Keystroke: conditional not using index value" + _controlName);
                                        }

                                    }
                                    break;

                                case "wait":
                                    logTofile(_eLogPtah, "[Waiting for ] " + _controlName);
                                    Console.WriteLine("Waiting for : " + _controlName);
                                    Thread.Sleep(int.Parse(_controlName) * 1000);
                                    uilog.AddTexttoColumn("Action Performed on Control", "Wait for " + _controlName + "Seconds");
                                    break;

                                case "conditionalwait":
                                    logTofile(_eLogPtah, "[Checking if Conditional wait is to be used]");
                                    if (_controlValue.Length > 0)
                                    {
                                        logTofile(_eLogPtah, "[Using Conditional wait] " + _controlName);
                                        Console.WriteLine("Waiting for : " + _controlName);
                                        Thread.Sleep(int.Parse(_controlName) * 1000);
                                        uilog.AddTexttoColumn("Action Performed on Control", "Wait for " + _controlName + "Seconds");


                                    }
                                    break;
                                /*  case "pagedown":
                                      Console.WriteLine("pagedown");
                                      _globalWindow.Focus();
                                      Thread.Sleep(1000);
                                      _globalWindow.Keyboard.PressSpecialKey(White.Core.WindowsAPI.KeyboardInput.SpecialKeys.PAGEDOWN);
                                      uilog.AddTexttoColumn("Action Performed on Control", "Pagedown");
                                      break; */

                                /*  case "pageup":
                                      _globalWindow.Focus();
                                      _globalWindow.Keyboard.PressSpecialKey(White.Core.WindowsAPI.KeyboardInput.SpecialKeys.PAGEUP);
                                      break; */
                                case "refresh":
                                    break;
                                case "clearwindow":
                                    // When we switch to a new UI automation window we must use "clear windo" in new structure sheet new window
                                    uiAutomationWindow = null;
                                    uilog.AddTexttoColumn("Action Performed on Control", "clearwindow");
                                    break;
                                default:
                                    logTofile(_eLogPtah, "[AddData]:Other Action than Specified Action: ---> " + _action.Trim().ToLower());
                                    throw new Exception("Valid action types are keyboard, wait, pagedown, pageup");
                            }
                            #endregion ActionColumn
                        }
                        if (Convert.IsDBNull(testData.Structure.Rows[i]["ControlType"]) == false)
                        {
                            _controlType = (string)testData.Structure.Rows[i]["ControlType"].ToString().ToLower();
                            //    _logicalName = (string)testData.Structure.Rows[i]["FieldName"].ToString();
                            _searchBy = (string)testData.Structure.Rows[i]["SearchBy"];
                            Console.WriteLine(_logicalName);
                            logTofile(_eLogPtah, "Field Name : =" + _logicalName);
                            if (Convert.IsDBNull(testData.Structure.Rows[i]["Index"]) == false)
                            {
                                logTofile(_eLogPtah, "Trying to parse index for : " + _logicalName);
                                _index = int.Parse(testData.Structure.Rows[i]["Index"].ToString());
                                logTofile(_eLogPtah, "Index was parsed. : " + _logicalName);
                            }

                            if (_logicalName.Length > 0)
                            {
                                _controlValue = (string)testData.Data.Rows[rowPosition][_logicalName].ToString();
                            }
                            if (_logicalName.Length > 0 && _controlValue.Length == 0)
                            {
                                logTofile(_eLogPtah, "controlValue Length : " + _controlValue.Length);
                                logTofile(_eLogPtah, "both Logical name was of 0 lenght and  Control valuewas of 0 lenght :nothing doing !!!" + _logicalName);
                            }
                            else
                            {
                                #region ControlTypes
                                uilog.AddTexttoColumn("ControlType", _controlType);
                                switch (_controlType.Trim().ToLower())
                                {

                                    /* case "wpfmenuitem":
                                         wpfapp.GetWPFMenuItem(_globalMenu, _controlName).Click();
                                         break;

                                     case "wpftoolstrip":
                                         wpfapp.GetWPFToolStrip(_globalWindow, _controlName).Focus();
                                         break;

                                     case "wpflistbox":
                                         wpfapp.GetWPFListBox(_globalWindow, _searchBy, _controlName).Focus();
                                         wpfapp.GetWPFListBox(_globalWindow, _searchBy, _controlName).Item(_controlValue).Select();
                                         break;

                                     case "wpflabel":
                                         string labelName = wpfapp.GetWPFLabel(_globalWindow, _searchBy, _controlName).Text.ToString();
                                         testData.UpdateTestData(testData.TestDataFile, testData.TestCase, _logicalName, labelName);
                                         break; */

                                    case "uiautomationmenu":
                                    case "umenu":
                                        {
                                            if (_controlValue.Length > 0)
                                            {

                                                AutomationElement umenu = GetUIAutomationMenu(_searchBy, _controlName, _index);
                                                logTofile(_eLogPtah, "returned menu :" + umenu.Current.Name);
                                                ClickControl(umenu);
                                                Thread.Sleep(2000);
                                            }
                                            break;

                                        }
                                    case "uiautomationmenuitem":
                                    case "umenuitem":
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                AutomationElement umenuitem = GetUIAutomationMenuItem(_searchBy, _controlName, _index);
                                                try
                                                {
                                                    ClickControl(umenuitem);
                                                }
                                                catch (Exception ex)
                                                {
                                                    logTofile(_eLogPtah, "Execption from click control was " + ex.Message.ToString());
                                                    logTofile(_eLogPtah, "Using Invoke Pattern for Menuitem alternately ");
                                                    InvokePattern invkptn = (InvokePattern)umenuitem.GetCurrentPattern(InvokePattern.Pattern);
                                                    invkptn.Invoke();
                                                }
                                            }
                                            break;
                                        }
                                    case "umenuiteminvoke":
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                AutomationElement umenuiteminvoke = GetUIAutomationMenuItem(_searchBy, _controlName, _index);
                                                string uname = umenuiteminvoke.Current.Name.ToString();
                                                try
                                                {
                                                    InvokePattern invkbuttonptn = (InvokePattern)umenuiteminvoke.GetCurrentPattern(InvokePattern.Pattern);
                                                    logTofile(_eLogPtah, "Got the Invoke Pattern");
                                                    logTofile(_eLogPtah, "[AddData][uitem]: Menuitem will be clicked : " + _controlValue.ToString() + ":Times");
                                                    invkbuttonptn.Invoke();
                                                    System.Threading.Thread.Sleep(20);
                                                    logTofile(_eLogPtah, "[AddData][uitem]:Clicked menuitem : " + uname);

                                                    uilog.AddTexttoColumn("Action Performed on Control", "Invoke Pattern");
                                                }
                                                catch
                                                {
                                                    logTofile(_eLogPtah, "No  Invoke Pattern hence using Click Control");
                                                    ClickControl(umenuiteminvoke);
                                                    uilog.AddTexttoColumn("Action Performed on Control", "Invoke Pattern");
                                                }
                                            }
                                            break;
                                        }
                                    case "uiautomationthumb":
                                    case "uthumb":
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                AutomationElement uthumb = GetUIAutomationThumb(_searchBy, _controlName, _index);
                                                if (uthumb != null)
                                                {
                                                    logTofile(_eLogPtah, "[Adddata]:[uiautomationthumb]:found control trying to double click it");
                                                    try
                                                    {
                                                        DoubleClickControl(uthumb);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        logTofile(_eLogPtah, "[Adddata]:[uiautomationthumb] Error--.>" + ex.Message.ToString());
                                                    }
                                                }
                                                else
                                                {
                                                    logTofile(_eLogPtah, "[Adddata]:[uiautomationthumb] thumb control Not found ");
                                                }

                                            }
                                            break;
                                        }
                                    case "uiautomationtext":
                                    case "utext":
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                AutomationElement utext = GetUIAutomationText(_searchBy, _controlName, _index);

                                                switch (_controlValue.ToLower())
                                                {
                                                    case "1":
                                                        ClickControl(utext);
                                                        break;
                                                    case "l":
                                                        ClickControl(utext);
                                                        break;
                                                    case "r":
                                                        RightClickControl(utext);
                                                        break;
                                                    case "d":
                                                        DoubleClickControl(utext);
                                                        break;
                                                    default:
                                                        Console.WriteLine("No valid input provided for clicking text");
                                                        break;
                                                }
                                            }
                                            break;
                                        }
                                    case "uiautomationimage":
                                    case "uimage":
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                AutomationElement uimage = GetUIAutomationImage(_searchBy, _controlName, _index);
                                                switch (_controlValue.ToLower())
                                                {
                                                    case "1":
                                                        ClickControl(uimage);
                                                        break;
                                                    case "l":
                                                        ClickControl(uimage);
                                                        break;
                                                    case "r":
                                                        RightClickControl(uimage);
                                                        break;
                                                    case "d":
                                                        DoubleClickControl(uimage);
                                                        break;
                                                    default:
                                                        Console.WriteLine("No valid input provided for clicking image");
                                                        break;
                                                }
                                            }
                                            break;

                                        }

                                    case "uiautomationdataitem":
                                    case "udataitem":
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                logTofile(_eLogPtah, "control value : " + _controlValue);
                                                AutomationElement udtitem = GetUIAutomationDataItem(_searchBy, _controlName, _index);
                                                SelectionItemPattern selpat = (SelectionItemPattern)udtitem.GetCurrentPattern(SelectionItemPattern.Pattern);
                                                selpat.Select();
                                                switch (_controlValue.ToLower())
                                                {
                                                    case "d":
                                                        DoubleClickControl(udtitem);
                                                        break;
                                                    case "l":
                                                        ClickControl(udtitem);
                                                        break;
                                                    case "r":
                                                        RightClickControl(udtitem);
                                                        break;
                                                    case "1":
                                                        ClickControl(udtitem);
                                                        break;
                                                }
                                            }
                                        }
                                        break;

                                    case "uiautomationspinner":
                                    case "uspinner":
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                logTofile(_eLogPtah, "control value:" + _controlValue);
                                                AutomationElement uspinner = GetUIAutomationSpinner(_searchBy, _controlName, _index);
                                                uspinner.SetFocus();

                                            }
                                        }
                                        break;
                                    /*  case "wpflistview":
                                          wpfapp.GetWPFDataGrid(_globalWindow, _searchBy, _controlName).Focus();
                                          break; */

                                    /*  case "wpfcombobox":
                                          if (_controlValue.Length > 0)
                                          {
                                              wpfapp.GetWPFComboBox(_globalWindow, _searchBy, _controlName, _index).Select(_controlValue);
                                              uilog.AddTexttoColumn("Control Detected", "Yes");
                                              uilog.AddTexttoColumn("Action Performed on Control", "select Combo item" + _controlValue);
                                          }
                                          break; */
                                    case "uiautomationcombobox":
                                    case "ucombobox":
                                        bool itemClicked = false;
                                        AutomationElement combo = GetUIAutomationComboBox(_searchBy, _controlName, _index);
                                        try
                                        {
                                            combo.SetFocus();
                                            ExpandCollapsePattern expandPat = (ExpandCollapsePattern)combo.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                            if (expandPat != null)
                                            {
                                                logTofile(_eLogPtah, "Expanding the combobox");
                                                expandPat.Expand();
                                                Thread.Sleep(100);
                                            }
                                        }
                                        catch
                                        {

                                        }
                                        //Click control was giving issue in K2, so added try catch
                                        try
                                        {
                                            ClickControl(combo);
                                        }
                                        catch
                                        {
                                        }

                                        System.Threading.Thread.Sleep(1000);

                                        //Control value is item to select
                                        AutomationElementCollection comboitems = combo.FindAll(TreeScope.Descendants,
                                            new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.ListItem));

                                        logTofile(_eLogPtah, "[Adddata][uiautomationcombobox]:collection count was " + comboitems.Count);

                                        for (int icb = 0; icb <= comboitems.Count - 1; icb++)
                                        {

                                            if (comboitems[icb].Current.Name.ToLower() == _controlValue.ToLower()) //if listitemname matches 
                                            {
                                                logTofile(_eLogPtah, "[Adddata][uiautomationcombobox]: Search By Name " + _controlValue.ToLower());
                                                SelectionItemPattern selpat = (SelectionItemPattern)comboitems[icb].GetCurrentPattern(SelectionItemPattern.Pattern);
                                                logTofile(_eLogPtah, "[Adddata][uiautomationcombobox]:match found for -->" + _controlValue.ToLower());
                                                selpat.Select();
                                                uilog.AddTexttoColumn("Action Performed on Control", "List items matched using selection pattern");
                                                break;
                                            }
                                            else if (_controlValue == icb.ToString())  //Index
                                            {
                                                logTofile(_eLogPtah, "[Adddata][uiautomationcombobox]: Search By Index " + _controlValue.ToLower());
                                                SelectionItemPattern selpat = (SelectionItemPattern)comboitems[icb].GetCurrentPattern(SelectionItemPattern.Pattern);
                                                logTofile(_eLogPtah, "[Adddata][uiautomationcombobox]:match found for -->" + _controlValue.ToLower());
                                                selpat.Select();
                                                uilog.AddTexttoColumn("Action Performed on Control", "Selection done by index");
                                                break;
                                            }
                                            else if (TreeWalker.ControlViewWalker.GetFirstChild(comboitems[icb]) != null) // Find by the text childnode of Listitem
                                            {

                                                AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(comboitems[icb]);
                                                string _controltype = elementNode.Current.LocalizedControlType.ToString();
                                                logTofile(_eLogPtah, "[Adddata][uiautomationcombobox]: Search By Text inside ListItem " + elementNode.Current.Name.ToLower());
                                                if (_controltype.ToLower() == "text" && elementNode.Current.Name.ToLower() == _controlValue.ToLower())
                                                {
                                                    logTofile(_eLogPtah, "[Adddata][uiautomationcombobox]: Match Found " + elementNode.Current.Name.ToLower());
                                                    try
                                                    {
                                                        ClickControl(elementNode);
                                                        itemClicked = true;
                                                        uilog.AddTexttoColumn("Action Performed on Control", "Text node selected using autoit " + _controlName);
                                                    }
                                                    catch
                                                    {

                                                    }
                                                    if (itemClicked == false)
                                                    {
                                                        SelectionItemPattern selpat = (SelectionItemPattern)comboitems[icb].GetCurrentPattern(SelectionItemPattern.Pattern);
                                                        if (selpat != null)
                                                        {
                                                            selpat.Select();
                                                            uilog.AddTexttoColumn("Action Performed on Control", "Text node selected using selection pattern " + _controlName);
                                                        }
                                                    }
                                                    logTofile(_eLogPtah, "[Adddata][uiautomationcombobox]:match found for -->" + _controlValue.ToLower());

                                                    break;
                                                }
                                            }

                                            else
                                            {
                                                logTofile(_eLogPtah, "[Adddata][uiautomationcombobox]:no condtition matched for  -->" + _controlValue.ToLower());
                                            }
                                        }


                                        break;

                                    /*   case "wpfcheckbox":
                                           if (_controlValue.Length > 0)
                                           {
                                               if (_controlValue.ToLower() == "on" || _controlValue.ToLower() == "1")
                                               {
                                                   wpfapp.GetWPFCheckBox(_globalWindow, _searchBy, _controlName, _index).Select();
                                                   uilog.AddTexttoColumn("Control Detected", "Yes");
                                                   uilog.AddTexttoColumn("Action Performed on Control", "Check the checkbox :");
                                               }
                                               else
                                               {
                                                   wpfapp.GetWPFCheckBox(_globalWindow, _searchBy, _controlName, _index).UnSelect();
                                                   uilog.AddTexttoColumn("Control Detected", "Yes");
                                                   uilog.AddTexttoColumn("Action Performed on Control", "UnCheck the checkbox :");
                                               }
                                           }
                                           break; */
                                    case "uiautomationcheckbox":
                                    case "ucheckbox":
                                        logTofile(_eLogPtah, "Inside [uiautomationcheckbox]:");
                                        if (_controlValue.Length > 0)
                                        {
                                            AutomationElement uiCheckBox = GetUIAutomationCheckBox(_searchBy, _controlName, _index);
                                            TogglePattern togPattern = (TogglePattern)uiCheckBox.GetCurrentPattern(TogglePattern.Pattern);
                                            ToggleState togstate = togPattern.Current.ToggleState;
                                            switch (_controlValue)
                                            {
                                                case "1":
                                                    if (togstate == ToggleState.Off)
                                                    {
                                                        togPattern.Toggle();
                                                    }
                                                    else
                                                        logTofile(_eLogPtah, "Checkbox is already checked");
                                                    break;
                                                case "0":
                                                    if (togstate == ToggleState.On)
                                                        togPattern.Toggle();
                                                    else
                                                        logTofile(_eLogPtah, "Checkbox is already unchecked");
                                                    break;
                                                case "c":
                                                    ClickControl(uiCheckBox);
                                                    break;
                                                default:
                                                    logTofile(_eLogPtah, "Provide valid input for control value i.e. either 1 or 0");
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "Checkbox for " + _controlName + " need not be checked");
                                        }
                                        break;

                                    case "uiautomationlistitem":
                                    case "ulistitem":
                                        logTofile(_eLogPtah, "Inside [uiautomationlistitem]:");
                                        logTofile(_eLogPtah, "Control value length: " + _controlValue.Length);
                                        if (_controlValue.Length > 0)
                                        {
                                            AutomationElement uiListItem = GetUIAutomationListItem(_searchBy, _controlName, _index);
                                            TogglePattern togPattern = (TogglePattern)uiListItem.GetCurrentPattern(TogglePattern.Pattern);
                                            logTofile(_eLogPtah, "[Adddata][uiautomationlistitem]:match found for -->" + _controlValue.ToLower());
                                            togPattern.Toggle();
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "ListItem for " + _controlName + " need not be checked");
                                        }
                                        break;
                                    case "codeduibutton":
                                        {

                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinButton winbutton = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (win == null)
                                            {
                                                logTofile(_eLogPtah, "Unable to find the window");
                                            }
                                            if (_index == -1)
                                            {
                                                winbutton = app.GetCUITButton(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                winbutton = app.GetCUITButton(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Inside [uiautomationlistitem]:");
                                            Mouse.Click(winbutton);
                                            Playback.Cleanup();
                                            break;
                                        }
                                    case "codeduidatarowheader":
                                        {

                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinRowHeader winrowheader = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                winrowheader = app.GetCUITDataRowHeader(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                winrowheader = app.GetCUITDataRowHeader(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Inside [uiautomationlistitem]:");
                                            Mouse.Click(winrowheader);
                                            Playback.Cleanup();
                                            break;
                                        }
                                    case "codeduidatacolumnheader":
                                        {

                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinColumnHeader wincolumnheader = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                wincolumnheader = app.GetCUITDataColumnHeader(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                wincolumnheader = app.GetCUITDataColumnHeader(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Inside [uiautomationlistitem]:");
                                            Mouse.Click(wincolumnheader);
                                            Playback.Cleanup();
                                            break;
                                        }
                                    case "codeduiradiobutton":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinRadioButton radioButton = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                radioButton = app.GetCUITRadioButton(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                radioButton = app.GetCUITRadioButton(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            Mouse.Click(radioButton);
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "codeduidatarow":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinRow row = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                row = app.GetCUITDataRow(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                row = app.GetCUITDataRow(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "codeduidatacell":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinCell cell = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                cell = app.GetCUITDataCell(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                cell = app.GetCUITDataCell(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            cell.Value = _controlValue;
                                            Playback.Cleanup();
                                            break;
                                        }
                                    case "codeduimenuitem":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinMenuItem menuItem = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                menuItem = app.GetCUITMenuItem(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                menuItem = app.GetCUITMenuItem(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            switch (_controlValue.ToLower())
                                            {
                                                case "1":
                                                    Mouse.Click(menuItem);
                                                    break;
                                                case "r":
                                                    Mouse.Click(menuItem, System.Windows.Forms.MouseButtons.Right);
                                                    break;
                                                case "d":
                                                    Mouse.DoubleClick(menuItem);
                                                    break;
                                                default:
                                                    Console.WriteLine("No valid input provided for clicking image");
                                                    break;
                                            }
                                            Mouse.DoubleClick(menuItem);
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "codeduilistitem":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinListItem listitem = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                listitem = app.GetCUITListItem(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                listitem = app.GetCUITListItem(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            Mouse.Click(listitem);
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "codeduilist":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinList list = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                list = app.GetCUITList(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                list = app.GetCUITList(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            int item = int.Parse(_controlValue);
                                            int[] selected = new int[1];
                                            selected[0] = item - 1;
                                            list.SelectedIndices = selected;
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "codeduitextcontrol":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinText text = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                text = app.GetCUITTextcontrol(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                text = app.GetCUITTextcontrol(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            Mouse.Click(text);
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "codeduitreeitem":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinTreeItem treeItem = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                treeItem = app.GetCUITTreeItem(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                treeItem = app.GetCUITTreeItem(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            switch (_controlValue.ToLower())
                                            {
                                                case "1":
                                                    Mouse.Click(treeItem);
                                                    break;
                                                case "r":
                                                    Mouse.Click(treeItem, System.Windows.Forms.MouseButtons.Right);
                                                    break;
                                                case "d":
                                                    Mouse.DoubleClick(treeItem);
                                                    break;
                                                default:
                                                    Console.WriteLine("No valid input provided for clicking image");
                                                    break;
                                            }
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "codeduicheckbox":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinCheckBox checkBox = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                checkBox = app.GetCUITCHeckbox(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                checkBox = app.GetCUITCHeckbox(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            if (_controlValue == "0")
                                            {
                                                checkBox.Checked = false;
                                            }
                                            else
                                            {
                                                checkBox.Checked = true;
                                            }
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "codeduitabpage":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinTabPage tab = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);
                                            if (_index == -1)
                                            {
                                                tab = app.GetCUITTabpage(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                tab = app.GetCUITTabpage(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            Mouse.Click(tab);
                                            logTofile(_eLogPtah, "Clicked Tab Page");
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "codeduiedit":
                                        {
                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinEdit edit = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);


                                            if (_index == -1)
                                            {
                                                edit = app.GetCUITEdit(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                edit = app.GetCUITEdit(win, _searchBy, _controlName, _index);
                                            }


                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            if (edit.Enabled)
                                            {
                                                edit.SetFocus();
                                                edit.Text = _controlValue;
                                            }
                                            Playback.Cleanup();
                                            break;


                                        }
                                    case "clickcordinates":
                                        {
                                            string[] coordinates;
                                            coordinates = _controlName.Split(',');
                                            AutoItX3Lib.AutoItX3 at1 = new AutoItX3Lib.AutoItX3();
                                            int x1 = Convert.ToInt32(coordinates[0]);
                                            int y1 = Convert.ToInt32(coordinates[1]);
                                            at1.MouseClick("LEFT", x1, y1, 1);
                                            break;
                                        }
                                    case "codeduicombobox":
                                        {


                                            clsCUIT_app app = new clsCUIT_app();
                                            app._eLogPtah = _eLogPtah;
                                            WinComboBox comboBox = null;
                                            WinWindow win = app.GetCUITWindow(parentSearchBy, parentSearchValue);

                                            if (_index == -1)
                                            {
                                                comboBox = app.GetCUITComboBox(win, _searchBy, _controlName, -1);
                                            }
                                            else
                                            {
                                                comboBox = app.GetCUITComboBox(win, _searchBy, _controlName, _index);
                                            }
                                            Playback.Initialize();
                                            logTofile(_eLogPtah, "Playback Initialized");
                                            int selectedIndex = int.Parse(_controlValue);
                                            comboBox.SelectedIndex = selectedIndex;
                                            Playback.Cleanup();
                                            break;


                                        }

                                    case "uiautomationselectlistitem":
                                    case "uselectlistitem":
                                    case "ulistitemselect":
                                        logTofile(_eLogPtah, "Inside [uiautomationselectlistitem]:");
                                        logTofile(_eLogPtah, "Control value length: " + _controlValue.Length);
                                        if (_controlValue.Length > 0)
                                        {
                                            AutomationElement uiListItem = GetUIAutomationListItem(_searchBy, _controlName, _index);
                                            SelectionItemPattern selpat = (SelectionItemPattern)uiListItem.GetCurrentPattern(SelectionItemPattern.Pattern);
                                            selpat.Select();

                                            logTofile(_eLogPtah, "[Adddata][uiautomationselectlistitem]:match found for -->" + _controlValue.ToLower());

                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "ListItem for " + _controlName + " need not be selected");
                                        }
                                        break;

                                    case "uiautomationclicklistitem":
                                    case "uclicklistitem":
                                    case "ulistitemclick":
                                        logTofile(_eLogPtah, "Inside [uiautomationclicklistitem]:");
                                        logTofile(_eLogPtah, "Control value length: " + _controlValue.Length);
                                        if (_controlValue.Length > 0)
                                        {
                                            AutomationElement uiListItem = GetUIAutomationListItem(_searchBy, _controlName, _index);
                                            switch (_controlValue.ToLower())
                                            {
                                                case "1":
                                                    ClickControl(uiListItem);
                                                    break;
                                                case "l":
                                                    ClickControl(uiListItem);
                                                    break;
                                                case "r":
                                                    RightClickControl(uiListItem);
                                                    break;
                                                case "d":
                                                    DoubleClickControl(uiListItem);
                                                    break;
                                                default:
                                                    Console.WriteLine("No valid input provided for clicking image");
                                                    break;
                                            }

                                            logTofile(_eLogPtah, "[Adddata][uiautomationselectlistitem]:match found for -->" + _controlValue.ToLower());

                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "ListItem for " + _controlName + " need not be selected");
                                        }
                                        break;

                                    case "uiautomationradiobutton":
                                    case "uradiobutton":
                                        if (_controlValue.ToLower().Equals("1"))
                                        {
                                            AutomationElement uiRadio = GetUIAutomationRadioButton(_searchBy, _controlName, _index);
                                            SelectionItemPattern selpat = (SelectionItemPattern)uiRadio.GetCurrentPattern(SelectionItemPattern.Pattern);
                                            logTofile(_eLogPtah, "[Adddata][uiautomationradiobutton]: match found for  -->" + _controlValue.ToLower());
                                            selpat.Select();
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "Radio Button for " + _controlName + " need not be selected");
                                        }
                                        break;
                                    /*   case "wpfbutton":
                                           if (_controlValue.Length > 0)
                                               logTofile(_eLogPtah, "[AddData]:[wpfbutton]:Trying to Click Button:== " + _logicalName);
                                           {
                                               try
                                               {
                                                   if (wpfapp.GetWPFButton(_globalWindow, _searchBy, _controlName, _index) != null)
                                                   {
                                                       uilog.AddTexttoColumn("Control Detected", "Yes");
                                                   }
                                                   else
                                                   {
                                                       uilog.AddTexttoColumn("Control Detected", "No");
                                                   }
                                                   wpfapp.GetWPFButton(_globalWindow, _searchBy, _controlName, _index).Click();

                                                   uilog.AddTexttoColumn("Action Performed on Control", "Click  button [White method]: " + _controlName);
                                                   logTofile(_eLogPtah, "[AddData]:[wpfbutton]:Trying to Clicked Button:== " + _logicalName);
                                               }
                                               catch (Exception ex)
                                               {
                                                   logTofile(_eLogPtah, "Error in wpfbutton:" + ex.ToString());
                                               }
                                           }
                                           break; */


                                    case "splitterdropdown":
                                        logTofile(_eLogPtah, "Got Control type Splieeter");
                                        logTofile(_eLogPtah, "SearchBy " + _searchBy + "Search value" + _controlName + "index" + _index);



                                        break;
                                    case "pvmaskedit":
                                        if (_controlValue.Length > 0)
                                        {
                                            AutomationElement uiPvMaskedit = GetUIAutomationPVMaskEdit(_searchBy, _controlName, _index);

                                            ClickControl(uiPvMaskedit);
                                            Thread.Sleep(2000);
                                            KeyBoardEnter(_controlValue);

                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "Radio Button for " + _controlName + " need not be selected");
                                        }
                                        break;

                                    case "pvcombobox":
                                        if (_controlValue.Length > 0)
                                        {
                                            AutomationElement uiPvCombo = GetUIAutomationbutton(_searchBy, _controlName, _index);

                                            ClickControl(uiPvCombo);
                                            Thread.Sleep(2000);
                                            Playback.Initialize();
                                            WinWindow cmbwin = new WinWindow();
                                            cmbwin.SearchProperties.Add(WinWindow.PropertyNames.Name, "ComboBox");

                                            while (cmbwin.Exists == false)
                                            {
                                                ClickControl(uiPvCombo);
                                                Playback.Wait(2000);

                                            }
                                            Console.WriteLine("Combowin exists confirmed");
                                            WinWindow listwin = new WinWindow(cmbwin);
                                            listwin.WindowTitles.Add("ComboBox");
                                            listwin.SetFocus();
                                            WinList flist = new WinList(listwin);
                                            flist.WindowTitles.Add("ComboBox");
                                            Playback.Wait(2000);
                                            int ik = 0;
                                            while (flist.Exists == false)
                                            {
                                                Console.WriteLine("Clcik was not performed with enough strength:  " + ik);
                                                ClickControl(uiPvCombo);
                                                Playback.Wait(2000);
                                                ik++;
                                            }
                                            Console.WriteLine("List obtained Confirmd");



                                            AutomationElement ae = AutomationElement.RootElement;
                                            Condition cond = new System.Windows.Automation.AndCondition(
                                                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                                new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, "ComboBox", PropertyConditionFlags.IgnoreCase)
                                                );
                                            AutomationElement cmbowin = ae.FindFirst(TreeScope.Descendants, cond);

                                            Condition cond2 =
                                               new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataItem);

                                            AutomationElementCollection alldataitems = cmbowin.FindAll(TreeScope.Descendants, cond2);
                                            Console.WriteLine("Got collection count =" + alldataitems.Count);
                                            foreach (AutomationElement inditem in alldataitems)
                                            {
                                                if (inditem.Current.Name == _controlValue)
                                                {
                                                    InvokePattern invk = (InvokePattern)inditem.GetCurrentPattern(InvokePattern.Pattern);
                                                    invk.Invoke();
                                                    break;
                                                }
                                            }


                                            //  flist.SelectedItemsAsString = selvalue;
                                            Playback.Cleanup();
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "Radio Button for " + _controlName + " need not be selected");
                                        }
                                        break;
                                    /*  case "wpfradiobutton":
                                          try
                                          {
                                              if (_controlValue.Length > 0)
                                              {
                                                  if (wpfapp.GetWPFRadioButton(_globalWindow, _searchBy, _controlName, _index) != null)
                                                  {
                                                      uilog.AddTexttoColumn("Control Detected", "Yes");
                                                  }
                                                  else
                                                  {
                                                      uilog.AddTexttoColumn("Control Detected", "No");
                                                  }
                                                  wpfapp.GetWPFRadioButton(_globalWindow, _searchBy, _controlName, _index).Select();
                                                  uilog.AddTexttoColumn("Control Detected", "Yes");
                                                  uilog.AddTexttoColumn("Action Performed on Control", "selected Radio button:" + _controlName);
                                                  Thread.Sleep(1000);
                                              }

                                          }
                                          catch (Exception ex)
                                          {
                                              uilog.AddTexttoColumn("Control Detected", "No");
                                              logTofile(_eLogPtah, "[AddData][wpfradiobutton]: Could not Find Object with Given Search conditions in application." + ex.Message.ToString());
                                          }
                                          break;
                                      case "wpftextbox":
                                          if (_controlValue != DBNull.Value.ToString())
                                          {
                                              if (_globalWindow == null)
                                              {
                                                  logTofile(_eLogPtah, "Add Data : Ooops Global window was got as Null !!!!");
                                              }
                                              if (wpfapp.GetWPFTextBox(_globalWindow, _searchBy, _controlName, _index) != null)
                                              {
                                                  uilog.AddTexttoColumn("Control Detected", "Yes");
                                              }
                                              else
                                              {
                                                  uilog.AddTexttoColumn("Control Detected", "No");
                                              }
                                              wpfapp.GetWPFTextBox(_globalWindow, _searchBy, _controlName, _index).SetValue(_controlValue);
                                              uilog.AddTexttoColumn("Action Performed on Control", "Enter Value in TextBox :" + _controlValue);
                                              System.Windows.Forms.SendKeys.SendWait("{TAB}");
                                          }
                                          break;

                                   /*   case "wpfmultilinetextbox":
                                          logTofile(_eLogPtah, "  [Add Data]->[wpfmultilinetextbox]-->looking for " + _logicalName);
                                          logTofile(_eLogPtah, "  [Add Data]->[wpfmultilinetextbox]-->Global Parent " + _globalWindow.Name.ToString());
                                          if (uiAutomationCurrentParent != null)
                                          {
                                              logTofile(_eLogPtah, " [AddData]->wpfmultileinetextbox uicurrent parent " + uiAutomationCurrentParent.Current.Name + " Automation Id" + uiAutomationCurrentParent.Current.AutomationId + "Control Type: " + uiAutomationCurrentParent.Current.ControlType.ToString()
                                                  );
                                          }
                                          else
                                          {
                                              logTofile(_eLogPtah, " [AddData]->wpfmultileinetextbox uicurrent parnet  was still null ");
                                          }
                                          if (_controlValue != DBNull.Value.ToString())
                                          {
                                              SearchCriteria _tbsearchcriteria = SearchCriteria.ByText(_controlName);
                                              {
                                                  logTofile(_eLogPtah, " Function --> Add Data- >global window was not set ..trying to set it  first ");
                                                  //   GetUIAutomationWindow(parentSearchBy, parentSearchValue);
                                              }
                                              try
                                              {
                                                  var tbeditor = _globalWindow.Get<MultilineTextBox>(_tbsearchcriteria);

                                                  if (tbeditor != null)
                                                  {
                                                      uilog.AddTexttoColumn("Control Detected", "Yes");
                                                  }
                                                  else
                                                  {
                                                      uilog.AddTexttoColumn("Control Detected", "No");
                                                  }
                                                  logTofile(_eLogPtah, "[AddData]->[wpfmultileinetextbox] : chekcing if this editor box is enabled or no " + tbeditor.Enabled);
                                                  // if (tbeditor.Enabled==true)
                                                  // {
                                                  tbeditor.Click();
                                                  System.Windows.Forms.SendKeys.Flush();
                                                  System.Windows.Forms.SendKeys.SendWait("{HOME}");
                                                  System.Windows.Forms.SendKeys.Flush();
                                                  System.Windows.Forms.SendKeys.SendWait("+{END}");
                                                  System.Windows.Forms.SendKeys.Flush();
                                                  System.Windows.Forms.SendKeys.SendWait("{DEL}");
                                                  // }
                                                  // *************Clear all vlaues ========================

                                                  // *************Clear all vlaues ========================
                                                  if (tbeditor != null)
                                                  {
                                                      logTofile(_eLogPtah, "[AddData]->[wpfmultileinetextbox] : multileline textbox  ");
                                                  }
                                                  else
                                                  {
                                                      logTofile(_eLogPtah, "[AddData]->[wpfmultileinetextbox] : Error in detecting  multilline textbox:  ");
                                                  }
                                                  tbeditor.Text = _controlValue;
                                                  logTofile(_eLogPtah, "[AddData]->[wpfmultileinetextbox] : Trying to  input on data editor ");
                                                  if (tbeditor.Text != _controlValue)
                                                  {
                                                      logTofile(_eLogPtah, "[AddData]->[wpfmultileinetextbox]  Error in inputing Data data was Not input properly by white  ");
                                                  }

                                                  uilog.AddTexttoColumn("Action Performed on Control", "Enter Value in TextBoxmulti (WPF) :" + _controlValue);
                                              }
                                              catch (Exception ex)
                                              {
                                                  logTofile(_eLogPtah, " [AddData]->[wpfmultileinetextbox] Error retriving WPfmultilene textbox " + ex.Message.ToString());
                                              }

                                          }
                                          break; */

                                    case "maskedwpftextbox":
                                        logTofile(_eLogPtah, " Inside->[maskedwpftextbox] ");
                                        AutomationElement maskeditbox = GetUIAutomationEdit(_searchBy, _controlName, _index);
                                        logTofile(_eLogPtah, "Obtained the maskedit box");
                                        logTofile(_eLogPtah, "bonding rec value" + maskeditbox.Current.BoundingRectangle.Y.ToString());
                                        ValuePattern editval1 = (ValuePattern)maskeditbox.GetCurrentPattern(ValuePattern.Pattern);
                                        editval1.SetValue("");
                                        editval1.SetValue(_controlValue);
                                        break;
                                    case "uiautomationedit":
                                    case "uedit":
                                        AutomationElement editbox = GetUIAutomationEdit(_searchBy, _controlName, _index);

                                        if (editbox != null)
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationedit]:-> controlname " + _controlName + " was found ");
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationedit]:-> controlname " + _controlName + " was NOT found ");
                                        }
                                        if (_controlValue != null)
                                        {
                                            try
                                            {

                                                editbox.SetFocus();
                                            }
                                            catch (Exception ex)
                                            {
                                                logTofile(_eLogPtah, "[AddData]->[uiautomationedit]:-setfocus issue encoutered " + ex.Message.ToString());
                                            }
                                        }

                                        logTofile(_eLogPtah, "[AddData]->[uiautomationedit]:- Searching for Value Pattern");
                                        ValuePattern editval = null;
                                        try
                                        {
                                            editval = (ValuePattern)editbox.GetCurrentPattern(ValuePattern.Pattern);
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationedit]:- Value Pattern found");
                                            editval.SetValue(_controlValue);
                                            uilog.AddTexttoColumn("Action Performed on Control", "Entered Value: [Value Pattern]" + _controlValue);
                                        }
                                        catch
                                        {
                                            if (editval == null) //no value pattern found 
                                            {
                                                logTofile(_eLogPtah, "[AddData]->[uiautomationedit]:- Value Pattern not found");
                                                ClickControl(editbox);
                                                System.Windows.Forms.SendKeys.Flush();
                                                System.Windows.Forms.SendKeys.SendWait(_controlValue);
                                            }

                                        }
                                        logTofile(_eLogPtah, "[AddData]->[uiautomationedit]:- Value entered successfully");

                                        break;
                                    case "uiautomationtextarea":
                                    case "utextarea":
                                        AutomationElement textarea = GetUIAutomationTextarea(_searchBy, _controlName, _index);
                                        if (textarea != null)
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationtextarea]:-> controlname " + _controlName + " was found ");
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationtextarea]:-> controlname " + _controlName + " was NOT found ");
                                        }
                                        if (_controlValue != null)
                                        {
                                            try
                                            {
                                                if (textarea.Current.IsKeyboardFocusable == true)
                                                {
                                                    textarea.SetFocus();
                                                }
                                                try
                                                {
                                                    ClickControl(textarea);
                                                }
                                                catch (Exception ex)
                                                {
                                                    logTofile(_eLogPtah, "[AddData]->[uiautomationtextarea]:-Clickble epoints issue: " + ex.Message.ToString());
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                logTofile(_eLogPtah, "[AddData]->[uiautomationtextarea]:-setfocus issue encoutered " + ex.Message.ToString());
                                            }
                                        }
                                        if (textarea.Current.IsEnabled == true)
                                        {
                                            TextPattern txtptn = (TextPattern)textarea.GetCurrentPattern(TextPattern.Pattern);
                                            // following loop is just to ensure that values are entered correctly in textbox.
                                            string rtt = txtptn.DocumentRange.GetText(10000);
                                            do
                                            {
                                                System.Windows.Forms.SendKeys.Flush();
                                                System.Windows.Forms.SendKeys.SendWait("{HOME}");
                                                System.Windows.Forms.SendKeys.Flush();
                                                System.Windows.Forms.SendKeys.SendWait("+{END}");
                                                System.Windows.Forms.SendKeys.Flush();
                                                System.Windows.Forms.SendKeys.SendWait("{DEL}");
                                                System.Windows.Forms.SendKeys.SendWait(_controlValue);

                                                txtptn = (TextPattern)textarea.GetCurrentPattern(TextPattern.Pattern);
                                                rtt = txtptn.DocumentRange.GetText(10000);
                                                logTofile(_eLogPtah, "checking for value " + rtt.ToString());
                                            } while (rtt != _controlValue);
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "Target TextBox is not enabled for entering data");
                                        }
                                        uilog.AddTexttoColumn("Action Performed on Control", "Enter Value: using keystrokes " + _controlValue);
                                        break;

                                    case "uiautomationcustominvokecontrol":
                                    case "ucustominvokecontrol":
                                    case "ucustominvoke":
                                        logTofile(_eLogPtah, "[AddData]->[uiautomationcustominvokecontrol] ControlValue" + _controlName + " : Length is " + _controlName.Length);
                                        AutomationElement customcontrol = GetUIAutomationCustominvokecontrol(_searchBy, _controlName, _index);
                                        try
                                        {
                                            InvokePattern invkptn = (InvokePattern)customcontrol.GetCurrentPattern(InvokePattern.Pattern);
                                            if (_controlValue.ToLower() == "y" || _controlValue == "1")
                                            {
                                                invkptn.Invoke();
                                                uilog.AddTexttoColumn("Action Performed on Control", "Invoke:");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            uilog.AddTexttoColumn("Action Performed on Control", "Failed in Invoke:");
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationcustominvokecontrol]" + ex.Message.ToString());
                                        }

                                        break;

                                    case "uiautomationcustomclickcontrol":
                                    case "ucustomclickcontrol":
                                        AutomationElement customclickcontrol = GetUIAutomationCustominvokecontrol(_searchBy, _controlName, _index);
                                        try
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                switch (_controlValue.ToLower())
                                                {
                                                    case "l":
                                                    case "1":
                                                        ClickControl(customclickcontrol);
                                                        break;
                                                    case "r":
                                                        RightClickControl(customclickcontrol);
                                                        break;
                                                    case "d":
                                                        DoubleClickControl(customclickcontrol);
                                                        break;
                                                    default:
                                                        Console.WriteLine("No valid input provided for clicking customclickcontrol");
                                                        break;
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationcustomclickcontrol]" + ex.Message.ToString());
                                        }

                                        break;

                                    case "uiautomationcustomvaluecontrol":
                                    case "ucustomvaluecontrol":
                                        AutomationElement customvaluecontrol1 = GetUIAutomationCustomvaluecontrol(_searchBy, _controlName, _index);
                                        if (customvaluecontrol1 != null)
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationcustomvaluecontrol] Control was found ");
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationcustomvaluecontrol] Control was NOt found ");
                                        }
                                        try
                                        {
                                            ValuePattern invkptn = (ValuePattern)customvaluecontrol1.GetCurrentPattern(ValuePattern.Pattern);
                                            invkptn.SetValue(_controlValue);
                                        }
                                        catch (Exception ex)
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationcustomvaluecontrol]" + ex.Message.ToString());
                                        }

                                        break;
                                    case "uiautomationcustomrightclickcontrol":
                                    case "ucustomrightclickcontrol":
                                        AutomationElement customcontrolrightclick = GetUIAutomationCustominvokecontrol(_searchBy, _controlName, _index);
                                        if (customcontrolrightclick != null)
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationcustomrightclickcontrol] Control was found ");
                                            try
                                            {
                                                RightClickControl(customcontrolrightclick);
                                            }
                                            catch (Exception ex)
                                            {
                                                logTofile(_eLogPtah, "[AddData]->[uiautomationcustomrightclickcontrol]" + ex.Message.ToString());
                                            }
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationcustomrightclickcontrol] Control was NOt found ");
                                        }


                                        break;

                                    case "uiautomationultratabitem":
                                    case "uultratabitem":
                                    case "utabitem":
                                        AutomationElement ultratabitem = GetUIAutomationUltratab(_searchBy, _controlName, _index);
                                        logTofile(_eLogPtah, "[AddData]->[uiautomationultratabitem] Ultra Tab Name" + ultratabitem.Current.Name);
                                        AutomationElement objbuttonnext = null;
                                        if (objbuttonnext == null)
                                        {
                                            try
                                            {
                                                if (ultratabitem.Current.IsKeyboardFocusable == true)
                                                {
                                                    ultratabitem.SetFocus();
                                                }

                                                ClickControl(ultratabitem);
                                                uilog.AddTexttoColumn("Action Performed on Control", "Click Object: Autoit " + _controlName);
                                                logTofile(_eLogPtah, "Clicked the ultratabitem using clickcontrol ");

                                            }
                                            catch (Exception e)
                                            {

                                                SelectionItemPattern selpat = (SelectionItemPattern)ultratabitem.GetCurrentPattern(SelectionItemPattern.Pattern);
                                                selpat.Select();
                                                uilog.AddTexttoColumn("Action Performed on Control", "Click Object: Select Pattern (from Catch block) " + _controlName);
                                                logTofile(_eLogPtah, "Clicked using Selection Pattern " + e.Message);

                                            }
                                        }
                                        else
                                        // need to detect until no error
                                        {
                                            try
                                            {
                                                ClickControl(ultratabitem);
                                            }
                                            catch (Exception e)
                                            {
                                                SelectionItemPattern selpat = (SelectionItemPattern)ultratabitem.GetCurrentPattern(SelectionItemPattern.Pattern);
                                                selpat.Select();
                                                logTofile(_eLogPtah, "Catch block " + e.Message);
                                                logTofile(_eLogPtah, "Clicked using Selection Pattern ");


                                            }
                                        }
                                        break;
                                    case "uiautomationselectultratabitem":
                                    case "uselectultratabitem":
                                        AutomationElement ultratabitemSelect = GetUIAutomationUltratab(_searchBy, _controlName, _index);
                                        logTofile(_eLogPtah, "[AddData]->[uiautomationselectultratabitem] Ultra Tab Name" + ultratabitemSelect.Current.Name);

                                        try
                                        {
                                            if (ultratabitemSelect.Current.IsKeyboardFocusable == true)
                                            {
                                                ultratabitemSelect.SetFocus();
                                            }

                                            SelectionItemPattern selpat = (SelectionItemPattern)ultratabitemSelect.GetCurrentPattern(SelectionItemPattern.Pattern);
                                            selpat.Select();
                                            uilog.AddTexttoColumn("Action Performed on Control", "Clicked using Selection Pattern " + _controlName);
                                            logTofile(_eLogPtah, "Clicked using Selection Pattern ");

                                        }
                                        catch (Exception e)
                                        {
                                            logTofile(_eLogPtah, "Error encountered in uiautomationselectultratabitem: " + e.Message);
                                            throw new Exception("uiautomationselectultratabitem" + e.Message);
                                        }
                                        break;
                                    case "uiautomationbutton":
                                    case "ubutton":
                                        if (_controlValue.Length > 0)
                                        {
                                            if (uiAutomationWindow.Current.Name.Length == 0)
                                            {
                                                uiAutomationWindow = GetUIAutomationWindow(parentSearchBy, parentSearchValue);
                                            }


                                            AutomationElement button = GetUIAutomationbutton(_searchBy, _controlName, _index);

                                            string bname = button.Current.Name.ToString();
                                            logTofile(_eLogPtah, "[AddData][uiautomationbutton]:Retrived button contol  : " + bname);
                                            logTofile(_eLogPtah, "Clicking the button using Invoke pattern");


                                            try
                                            {
                                                InvokePattern invkbuttonptn = (InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern);
                                                logTofile(_eLogPtah, "Got the Invoke Pattern");
                                                if (Convert.IsDBNull(_controlValue) == false)

                                                    if (Int32.Parse(_controlValue) > 0)
                                                    {
                                                        logTofile(_eLogPtah, "[AddData][uiautomationbutton]: Button will be clicked : " + _controlValue.ToString() + ":Times");
                                                        for (int ib = 0; ib < Int32.Parse(_controlValue); ib++)
                                                        {
                                                            invkbuttonptn.Invoke();
                                                            System.Threading.Thread.Sleep(20);
                                                            logTofile(_eLogPtah, "[AddData][uiautomationbutton]:Clicked button : " + bname + "Times:" + ib.ToString());
                                                        }
                                                    }
                                                uilog.AddTexttoColumn("Action Performed on Control", "Invoke Pattern");
                                            }
                                            catch
                                            {
                                                logTofile(_eLogPtah, "No  Invoke Pattern hence using Click Control");
                                                ClickControl(button);
                                                uilog.AddTexttoColumn("Action Performed on Control", "Invoke Pattern");
                                            }
                                        }

                                        break;

                                    case "uiautomationtogglebutton":
                                    case "utogglebutton":
                                        logTofile(_eLogPtah, "Inside [uiautomationtogglebutton]:");
                                        if (_controlValue.Length > 0)
                                        {
                                            AutomationElement uiTogbutton = GetUIAutomationbutton(_searchBy, _controlName, _index);
                                            TogglePattern togPattern = (TogglePattern)uiTogbutton.GetCurrentPattern(TogglePattern.Pattern);
                                            ToggleState togstate = togPattern.Current.ToggleState;
                                            switch (_controlValue)
                                            {
                                                case "1":
                                                    if (togstate == ToggleState.Off)
                                                        togPattern.Toggle();
                                                    else
                                                        logTofile(_eLogPtah, "Button is already in On state");
                                                    break;
                                                case "0":
                                                    if (togstate == ToggleState.On)
                                                        togPattern.Toggle();
                                                    else
                                                        logTofile(_eLogPtah, "Button is already in Off state");
                                                    break;
                                                default:
                                                    logTofile(_eLogPtah, "Provide valid input for control value i.e. either 1 or 0");
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            logTofile(_eLogPtah, "Button for " + _controlName + " need not be clicked");
                                        }
                                        break;

                                    case "uiautomationribbonbutton":
                                    case "uribbonbutton":
                                        if (uiAutomationWindow.Current.Name.Length == 0)
                                        {
                                            uiAutomationWindow = GetUIAutomationWindow(parentSearchBy, parentSearchValue);
                                        }
                                        AutomationElement ribbonbutton = GetUIAutomationRibbonButton(_searchBy, _controlName, _index);
                                        logTofile(_eLogPtah, "[AddData][uiautomationribbonbutton]:Retrived Ribbonbutton contol");
                                        ClickControl(ribbonbutton);
                                        uilog.AddTexttoColumn("Action Performed on Control", "Click Control using autoit");
                                        System.Threading.Thread.Sleep(20);
                                        logTofile(_eLogPtah, "[AddData][uiautomationribbonbutton]:Pressed Ribbonbutton   : " + _logicalName);


                                        break;

                                    case "uiautomationtreeitemclick":
                                    case "utreeitemclick":
                                        if (_controlValue.Length > 0)
                                        {
                                            _controlName = _controlValue;
                                        }
                                        AutomationElement treeitem = GetUIAutomationtreeitem(_searchBy, _controlName, _index);
                                        ClickControl(treeitem);
                                        uilog.AddTexttoColumn("Action Performed on Control", "Click Tree item with text: " + _controlName);

                                        break;

                                    case "uiautomationtreeitem":
                                    case "utreeitem":
                                        {
                                            if (_controlValue.Length > 0)
                                            {
                                                AutomationElement treeviewitem = GetUIAutomationtreeitem(_searchBy, _controlName, _index);

                                                switch (_controlValue.ToLower())
                                                {
                                                    case "l":
                                                        ClickControl(treeviewitem);
                                                        break;
                                                    case "r":
                                                        RightClickControl(treeviewitem);
                                                        break;
                                                    case "d":
                                                        DoubleClickControl(treeviewitem);
                                                        break;
                                                    default:
                                                        Console.WriteLine("No valid input provided for clicking treeitem");
                                                        break;
                                                }
                                            }
                                            break;
                                        }

                                    //this method is used for both expanding and collapsing the Tree nnodes
                                    case "uiautomationtreeitemexpand":
                                    case "utreeitemexpand":
                                        if (_controlValue.Length > 0)
                                        {
                                            _controlName = _controlValue;
                                        }
                                        AutomationElement treeitemcolapsed = GetUIAutomationtreeitem(_searchBy, _controlName, _index);
                                        DoubleClickControl(treeitemcolapsed);
                                        break;

                                    case "uiautomationtreeitemexpandk2":
                                    case "utreeitemexpandk2":
                                        if (_controlValue.Length > 0)
                                        {
                                            AutomationElement treeviewitem = GetUIAutomationtreeitem(_searchBy, _controlName, _index);
                                            ExpandCollapsePattern collapsepat = (ExpandCollapsePattern)treeviewitem.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                            ExpandCollapseState state = collapsepat.Current.ExpandCollapseState;
                                            if (state == ExpandCollapseState.Collapsed || state == ExpandCollapseState.PartiallyExpanded)
                                            {
                                                collapsepat.Expand();
                                                logTofile(_eLogPtah, "Tree item expanded");
                                            }
                                            else
                                                logTofile(_eLogPtah, " tree item is already expanded");
                                        }
                                        break;

                                    case "uiautomationtreeitemcollapsek2":
                                    case "utreeitemcollapsek2":
                                        if (_controlValue.Length > 0)
                                        {
                                            AutomationElement treeviewitem = GetUIAutomationtreeitem(_searchBy, _controlName, _index);
                                            ExpandCollapsePattern expandpat = (ExpandCollapsePattern)treeviewitem.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                            ExpandCollapseState state = expandpat.Current.ExpandCollapseState;
                                            if (state == ExpandCollapseState.Expanded)
                                            {
                                                expandpat.Collapse();
                                                logTofile(_eLogPtah, "Tree item collapsed");
                                            }
                                            else
                                                logTofile(_eLogPtah, " tree item is already collapsed");
                                        }
                                        break;

                                    case "uiautomationsyncfusionpane":
                                    case "usyncfusionpane":
                                        AutomationElement syncfusionpane = GetUIAutomationsyncfusionpane(_searchBy, _controlName, _index);
                                        //   syncfusionpane.SetFocus();this mehtod dud not work hence using clickable point
                                        logTofile(_eLogPtah, "Datagrid : " + syncfusionpane.Current.Name);
                                        Console.WriteLine("[AddData][uiautomationsyncfusionpane]:-->before trying getclick");
                                        ClickControl(syncfusionpane);
                                        uilog.AddTexttoColumn("Action Performed on Control", "Click object");
                                        if (_controlValue != null)
                                        {
                                            char[] celldellim = new char[] { '|' };
                                            string[] arrcelladd = _controlValue.Split(celldellim);
                                            foreach (var item in arrcelladd)
                                            {
                                                // split control value in an array
                                                char[] delim = new char[] { ';' };
                                                string[] arr = item.Split(delim);

                                                string rowval = arr[0];
                                                string colval = arr[1];
                                                string offsetval = arr[2];
                                                string dataval = arr[3];
                                                if (_controlName1.ToLower() != "{tab}" && _controlName1.ToLower() != "{right}")
                                                    System.Console.WriteLine("uiautomationsyncfusionpane]->Wrong value in French value column of structure sheet");
                                                setcellvaleps(rowval, colval, offsetval, dataval, _controlName1);
                                            }
                                        }
                                        break;
                                    case "uiautomationsyncfusionpanereo":
                                    case "usyncfusionpanereo":
                                        AutomationElement syncfusionpanereo = GetUIAutomationsyncfusionpane(_searchBy, _controlName, _index);
                                        //   syncfusionpane.SetFocus();this mehtod dud not work hence using clickable point
                                        logTofile(_eLogPtah, "Datagrid : " + syncfusionpanereo.Current.Name);
                                        Console.WriteLine("[AddData][uiautomationsyncfusionpane]:-->before trying getclick");
                                        //ClickControl(syncfusionpanereo);
                                        if (_controlValue != null)
                                        {
                                            char[] celldellim = new char[] { '|' };
                                            string[] arrcelladd = _controlValue.Split(celldellim);
                                            foreach (var item in arrcelladd)
                                            {
                                                // split control value in an array
                                                char[] delim = new char[] { ';' };
                                                string[] arr = item.Split(delim);

                                                string rowval = arr[0];
                                                string colval = arr[1];
                                                string offsetval = arr[2];
                                                string dataval = arr[3];
                                                if (_controlName1.ToLower() != "{tab}" && _controlName1.ToLower() != "{right}")
                                                    System.Console.WriteLine("uiautomationsyncfusionpane]->Wrong value in French value column of structure sheet");
                                                setcellvaleps(rowval, colval, offsetval, dataval, _controlName1);
                                            }
                                        }
                                        break;
                                    case "uiautomationinfratable":
                                    case "uinfratable":
                                        AutomationElement infratable = GetUIAutomationInfraTableFlat(_searchBy, _controlName, _index);
                                        //   syncfusionpane.SetFocus();this mehtod dud not work hence using clickable point
                                        Console.WriteLine("[AddData][uiautomationsyncfusionpane]:-->before trying getclick");
                                        ClickControl(infratable);
                                        break;

                                    case "uiainfratablerow":
                                    case "uinfratablerow":
                                        if (Convert.IsDBNull(_controlName) == true || _controlName == "")
                                        {
                                            _controlName = _controlValue;
                                            logTofile(_eLogPtah, "[AddData]->[uiainfratablerow]: triyng to find infra row");
                                            AutomationElement infratablerow = GetUIAutomationGroupInfraTableRow(_searchBy, _controlName, _index);
                                            InvokePattern pat = (InvokePattern)infratablerow.GetCurrentPattern(InvokePattern.Pattern);
                                            logTofile(_eLogPtah, "[AddData]->[uiainfratablerow]:found row now ...trying  invoke patterns ");
                                            pat.Invoke();
                                        }
                                        else
                                        {
                                            System.Console.WriteLine("[AddData]->[uiainfratablerow]->Arguments,controlname shud be blank but has value " + _controlName);
                                        }
                                        break;
                                    case "uiainfratablerowflat":
                                    case "uinfratablerowflat":
                                        if (Convert.IsDBNull(_controlName) == true || _controlName == "")
                                        {
                                            _controlName = _controlValue;
                                            logTofile(_eLogPtah, "[AddData]->[uiainfratablerowflat] triyng to find infra row");
                                            AutomationElement infratablerow = GetUIAutomationFlatInfraTableRow(_searchBy, _controlName, _index);
                                            InvokePattern pat = (InvokePattern)infratablerow.GetCurrentPattern(InvokePattern.Pattern);
                                            logTofile(_eLogPtah, "[AddData]->[uiainfratablerowflat]  found row now ...trying  invoke patterns ");
                                            pat.Invoke();
                                        }
                                        else
                                        {
                                            System.Console.WriteLine("[AddData]->[uiainfratablerowflat]: ->Arguments,controlname shud be blank but has value " + _controlName);
                                        }
                                        break;
                                    case "uiainfratablerowflatnumber":
                                    case "uinfratablerowflatnumber":
                                        if (Convert.IsDBNull(_controlName) == true || _controlName == "")
                                        {
                                            _controlName = _controlValue;
                                            logTofile(_eLogPtah, "[AddData]->[uiainfratablerowflatnumber] triyng to find infra row");
                                            if (Convert.ToInt32(_index) != 1)
                                            {
                                                AutomationElement infratablerow = GetUIAutomationFlatInfraTableRowByNumber(_searchBy, _controlName, _index);

                                                logTofile(_eLogPtah, "[AddData]->[uiainfratablerowflatnumbers]  found row with index as " + _controlName);
                                                InvokePattern pat = (InvokePattern)infratablerow.GetCurrentPattern(InvokePattern.Pattern);
                                                logTofile(_eLogPtah, "[AddData]->[uiainfratablerowflatnumbers]  found row now ...trying  invoke patterns ");
                                                pat.Invoke();
                                            }
                                            System.Windows.Forms.SendKeys.Flush();
                                            System.Windows.Forms.SendKeys.SendWait("{Down}");
                                            System.Windows.Forms.SendKeys.Flush();
                                            System.Windows.Forms.SendKeys.SendWait("{Up}");


                                        }
                                        else
                                        {
                                            System.Console.WriteLine("[AddData]->[uiainfratablerowflat]: ->Arguments,controlname shud be blank but has value " + _controlName);
                                        }
                                        break;
                                    /*   case "wpfdatepicker":
                                           if (_controlValue != DBNull.Value.ToString())
                                           {
                                               wpfapp.GetWPFDatePicker(_globalWindow, _searchBy, _controlName).SetValue(_controlValue);
                                           }
                                           break;

                                       case "wpftabitem":
                                           if (_controlValue != DBNull.Value.ToString())
                                           {
                                               wpfapp.GetWPFTabItem(_globalWindow, _searchBy, _controlName).Select();
                                           }
                                           break;

                                       case "wpftree":
                                           if (_controlValue != DBNull.Value.ToString())
                                           {
                                               wpfapp.GetWPFTreeViewWindow(_globalWindow, _searchBy, _controlName, _index).SetValue(_controlValue);
                                           }
                                           break; */
                                    case "uiautomationheader":
                                    case "uheader":
                                        if (_controlValue != DBNull.Value.ToString())
                                        {

                                            AutomationElement header1 = GetUIAutomationHeader(_searchBy, _controlName, _index);
                                            try
                                            {
                                                InvokePattern hinvoke = (InvokePattern)header1.GetCurrentPattern(InvokePattern.Pattern);
                                                hinvoke.Invoke();
                                            }
                                            catch (Exception e)
                                            {
                                                logTofile(_eLogPtah, "Execption" + e.Message.ToString());

                                            }
                                        }
                                        break;

                                    case "uiautomationheaderitem":
                                    case "uheaderitem":
                                        if (_controlValue != DBNull.Value.ToString())
                                        {

                                            AutomationElement headeritem = GetUIAutomationHeaderitem(_searchBy, _controlName, _index);
                                            try
                                            {
                                                switch (_controlValue.ToLower())
                                                {
                                                    case "1":
                                                        ClickControl(headeritem);
                                                        break;
                                                    case "l":
                                                        ClickControl(headeritem);
                                                        break;
                                                    case "r":
                                                        RightClickControl(headeritem);
                                                        break;
                                                    case "d":
                                                        DoubleClickControl(headeritem);
                                                        break;
                                                    default:
                                                        Console.WriteLine("No valid input provided for clicking headeritem");
                                                        break;
                                                }
                                            }
                                            catch (Exception e)
                                            {
                                                logTofile(_eLogPtah, "Execption" + e.Message.ToString());

                                            }
                                        }
                                        break;
                                    case "uicustomcalendarmatbal":
                                        {

                                            if (_controlValue.Length > 0)
                                            {
                                                AutomationElement showcalbutton = GetUIAutomationbutton(_searchBy, _controlName, _index);
                                                inputDateInCalendarControl1(showcalbutton, _controlValue);

                                            }
                                            break;
                                        }
                                    default:
                                        throw new Exception("[AddData]->[System.Windows.Automation.ControlType]:Not a valid control type.");

                                }
                                #endregion   ControlTypes
                            }
                        }
                    }
                    uilog.commitrow();
                    uilog.CreateCSVfile(_reportsPath, uiAfileName);
                    uilog.ClearDataTable();
                }


                #endregion recordsinexcel
            }
            catch (Exception ex)
            {
                //todo add logging comments
                logTofile(_eLogPtah, "Erroring Line number in [Adddata] " + GetStacktrace(ex).ToString());

                throw new Exception(_error + "[AddData]:" + System.Environment.NewLine + ex.Message);

            }
        }
        //this function is called inside AddData(string testDataPath, string testCase)
        public void AddData(string testDataPath, string testCase)
        {
            try
            {
                string repeat = new string('=', 50);
                logTofile(_eLogPtah, repeat + " Start  of Add Data " + DateTime.Now.ToString() + repeat);
                logTofile(_eLogPtah, "Section : --> Excel Connection: Searching for excel file " + testDataPath + " Test case : " + testCase);
                testData.GetTestData(testDataPath, testCase);
                ptestDataPath = testDataPath;
                ptestCase = testCase;
                logTofile(_eLogPtah, "Section : --> Excel Connection: excel file" + testDataPath + "was found!!");
                AddData(0);
                logTofile(_eLogPtah, repeat + " End  of Add Data " + DateTime.Now.ToString() + repeat);
            }

            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        #endregion
        //this function adds data from an excel sheet
        #region VerifyData
        private void VerifyData(int rowposition1)
        {


            string parentType = "";
            string parentSearchBy = "";
            string parentSearchValue = "";
            string controlaction = "";
            var _controlType = "";
            var _logicalName = "";


            try
            {
                for (int i = 0; i < testData.Structure.Rows.Count; i++)
                {
                    parentType = testData.Structure.Rows[i]["ParentType"].ToString();
                    parentSearchBy = testData.Structure.Rows[i]["ParentSearchBy"].ToString().ToLower();
                    parentSearchValue = testData.Structure.Rows[i]["ParentSearchValue"].ToString();
                    controlaction = testData.Structure.Rows[i]["ParentSearchValue"].ToString();



                    if ((string)testData.Structure.Rows[i]["inputdata"].ToString().ToLower() == "y")
                    {
                        if (Convert.IsDBNull(testData.Structure.Rows[i]["ParentType"]) == false)
                        {
                            switch (parentType.Trim().ToLower())
                            {
                                case "window":
                                    {
                                        //wpfapp.GetWPFWindow(parentSearchValue);
                                        break;
                                    }
                                /*   case "groupbox":
                                       {
                                           if (_immediateParent.Trim().ToLower() == "window")

                                               wpfapp.GetWPFGroupBox(_globalWindow, parentSearchBy, parentSearchValue);
                                           else
                                               wpfapp.GetWPFGroupBox(_globalGroup, parentSearchBy, parentSearchValue);
                                           break;
                                       }
                                 * */

                                /* case "wpfmenu":
                                     {
                                         _globalMenu = wpfapp.GetWPFMenu(_globalWindow, parentSearchBy, parentSearchValue);
                                         _globalMenu.Click();
                                         break;
                                     } */
                                case "uiautomationwindow":
                                case "uwindow":
                                    {

                                        //  if (uiAutomationWindow == null || uiAutomationWindow.Current.Name != parentSearchValue)

                                        if (uiAutomationWindow == null)
                                        {
                                            uiAutomationWindow = GetUIAutomationWindow(parentSearchBy, parentSearchValue);

                                        }
                                        else
                                        {
                                            uiAutomationCurrentParent = uiAutomationWindow;
                                            logTofile(_eLogPtah, "[AddData]->[uiautomationwindow]->UI automationwindow was already set " + uiAutomationWindow.Current.Name.ToString());
                                        }
                                        break;

                                    }

                                case "uiautomationchildwindow":
                                case "uchildwindow":
                                    {


                                        uiAutomationCurrentParent = GetUIAutomationWindow(parentSearchBy, parentSearchValue);
                                        logTofile(_eLogPtah, "[AddData]->[uiautomationchildwindow]->UI automationchild window was already set " + uiAutomationCurrentParent.Current.Name.ToString());

                                        break;
                                    }
                                case "uiautomationtreeitem":
                                case "utreeitem":
                                    {
                                        uiAutomationCurrentParent = GetUIAutomationtreeitem(parentSearchBy, parentSearchValue, 0);
                                        break;
                                    }
                                default:
                                    throw new Exception("Not a valid parent type.");
                            }
                        }
                        var _action = "";
                        var _searchBy = "";
                        var _index = -1;
                        var _controlName = "";

                        if (Convert.IsDBNull(testData.Structure.Rows[i]["ControlName"]) == false)
                            _controlName = (string)testData.Structure.Rows[i]["ControlName"];

                        if (Convert.IsDBNull(testData.Structure.Rows[i]["Action"]) == false)
                        {
                            _action = (string)testData.Structure.Rows[i]["Action"];
                            Console.WriteLine("Action:" + _action);

                            switch (_action.Trim().ToLower())
                            {
                                case "keyboard":
                                    System.Windows.Forms.SendKeys.Flush();
                                    System.Windows.Forms.SendKeys.SendWait("{" + _controlName + "}");
                                    break;
                                case "wait":
                                    Console.WriteLine("Waiting for : " + _controlName);
                                    Thread.Sleep(int.Parse(_controlName) * 1000);
                                    break;

                                /*   case "pagedown":
                                       Console.WriteLine("pagedown");
                                       _globalWindow.Focus();
                                       Thread.Sleep(1000);
                                       _globalWindow.Keyboard.PressSpecialKey(White.Core.WindowsAPI.KeyboardInput.SpecialKeys.PAGEDOWN);
                                       break;

                                   case "pageup":
                                       _globalWindow.Focus();
                                       _globalWindow.Keyboard.PressSpecialKey(White.Core.WindowsAPI.KeyboardInput.SpecialKeys.PAGEUP);
                                       break; */
                                case "refresh":
                                    break;
                                default:
                                    throw new Exception("Valid action types are keyboard, wait, pagedown, pageup");
                            }
                        }
                        if (Convert.IsDBNull(testData.Structure.Rows[i]["ControlType"]) == false)
                        {
                            if (IsColumnPresent("System.Windows.Automation.ControlType"))
                            {
                                if (Convert.IsDBNull(testData.Structure.Rows[i]["System.Windows.Automation.ControlType"]) == false)
                                {
                                    _controlType = (string)testData.Structure.Rows[i]["System.Windows.Automation.ControlType"].ToString().ToLower();
                                }
                            }
                            else
                            {
                                _controlType = (string)testData.Structure.Rows[i]["ControlType"].ToString().ToLower();
                            }
                            _logicalName = (string)testData.Structure.Rows[i]["FieldName"].ToString();
                            _searchBy = (string)testData.Structure.Rows[i]["SearchBy"];

                            Console.WriteLine(_logicalName);

                            if (Convert.IsDBNull(testData.Structure.Rows[i]["Index"]) == false)
                            {
                                _index = int.Parse(testData.Structure.Rows[i]["Index"].ToString());
                            }
                            _testcase = (string)testData.Data.Rows[0]["testcase"].ToString();
                            string _controlValue = null;
                            if (_logicalName.Length > 0)
                            {
                                _controlValue = (string)testData.Data.Rows[0][_logicalName].ToString();
                            }
                            if (_logicalName.Length > 0 && _controlValue.Length == 0)
                            {

                            }
                            else
                            {
                                switch (_controlType.Trim().ToLower())
                                {

                                    //  case "wpfmenuitem":
                                    //      wpfapp.GetWPFMenuItem(_globalMenu, _controlName).Click();
                                    //      break;

                                    case "uiautomationsyncfusionpane":
                                    case "usyncfusionpane":
                                        AutomationElement syncfusionpane = GetUIAutomationsyncfusionpane(_searchBy, _controlName, _index);
                                        //   syncfusionpane.SetFocus();this mehtod dud not work hence using clickable point

                                        ClickControl(syncfusionpane);
                                        logTofile(_eLogPtah, "[VerifyData]: Inside Syncfusiion pane");

                                        char[] celldellim = new char[] { ';' };
                                        string[] arr = _controlValue.Split(celldellim);
                                        string testdataPath = arr[0];
                                        string paramfiletcase = arr[1];
                                        string expectedfileName = arr[2];
                                        string tcase = arr[3];
                                        string otptfile = arr[4];
                                        logTofile(_eLogPtah, "[VerifyData]: Trying to update paramfile");

                                        action3.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcase, "expectedFile", _testDataPath + expectedfileName);
                                        action3.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcase, "testcaseID", tcase);
                                        action3.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcase, "tempFilePath", @"C:\created.txt");
                                        action3.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcase, "resultspath", otptfile);
                                        if (File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls")))
                                        {
                                            System.IO.File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls"));
                                            System.Threading.Thread.Sleep(2000);
                                        }
                                        System.IO.File.Copy(_testDataPath + @"paramFile.xls", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls"));

                                        //put information of testdata dynamicaly in a vbs file to read 
                                        logTofile(_eLogPtah, "[VerifyData]: uiautomationsyncfusionpane updated paramfile trying to execute vbs");
                                        RunVBS(_testDataPath + @"VerifySyncFusionGridData.vbs");
                                        System.IO.File.Delete(@"C:\created.txt");
                                        if (File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls")))
                                        {
                                            System.IO.File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls"));
                                            System.Threading.Thread.Sleep(2000);
                                        }
                                        break;
                                    case "uiautomationsyncfusionpanereo":
                                    case "usyncfusionpanereo":
                                        AutomationElement syncfusionpanereo = GetUIAutomationsyncfusionpane(_searchBy, _controlName, _index);
                                        //   syncfusionpane.SetFocus();this mehtod dud not work hence using clickable point

                                        //    ClickControl(syncfusionpane);
                                        logTofile(_eLogPtah, "[VerifyData]: Inside Syncfusiion pane");

                                        char[] celldellimreo = new char[] { ';' };
                                        string[] arrreo = _controlValue.Split(celldellimreo);
                                        string testdataPathreo = arrreo[0];
                                        string paramfiletcasereo = arrreo[1];
                                        string expectedfileNamereo = arrreo[2];
                                        string tcasereo = arrreo[3];
                                        string otptfilereo = arrreo[4];
                                        logTofile(_eLogPtah, "[VerifyData]: Trying to update paramfile");

                                        action3.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcasereo, "expectedFile", _testDataPath + expectedfileNamereo);
                                        action3.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcasereo, "testcaseID", tcasereo);
                                        action3.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcasereo, "tempFilePath", @"C:\created.txt");
                                        action3.UpdateTestData(_testDataPath + @"paramFile.xls", paramfiletcasereo, "resultspath", otptfilereo);
                                        if (File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls")))
                                        {
                                            System.IO.File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls"));
                                            System.Threading.Thread.Sleep(2000);
                                        }
                                        System.IO.File.Copy(_testDataPath + @"paramFile.xls", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls"));

                                        //put information of testdata dynamicaly in a vbs file to read 
                                        logTofile(_eLogPtah, "[VerifyData]: uiautomationsyncfusionpane updated paramfile trying to execute vbs");
                                        RunVBS(_testDataPath + @"VerifySyncFusionGridData.vbs");
                                        System.IO.File.Delete(@"C:\created.txt");
                                        if (File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls")))
                                        {
                                            System.IO.File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "paramFile.xls"));
                                            System.Threading.Thread.Sleep(2000);
                                        }
                                        break;

                                    default:
                                        throw new Exception("Not a valid control type.");

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //todo add logging comments
                throw new Exception(_error + "AddData:" + System.Environment.NewLine + ex.Message);
            }
        }
        public void VerifyData(string testDataPath, string testCase)
        {
            try
            {
                testData.GetTestData(testDataPath, testCase);

                VerifyData(0);
            }

            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        #endregion
        /*
         *
        public ListView GetDataGridFromTemplate(string templateFile)
        {
            ListView dataGrid = null;
            try
            {
                testData.GetVerificationData(templateFile, "T1");
                testData.ExpectedData = testData.Data.Copy();
                string parentSearchBy = "";
                string parentSearchValue = "";
                string parentType = "";
                string gridSearchBy = "";
                string gridSearchValue = "";
                //string gridType = "";
                parentSearchBy = (String)testData.Template.Rows[0]["ParentSearchBy"];
                parentSearchValue = (String)testData.Template.Rows[0]["ParentSearchValue"];
                parentType = (String)testData.Template.Rows[0]["ParentControl"];
                gridSearchBy = (String)testData.Template.Rows[0]["SearchBy"];
                gridSearchValue = (String)testData.Template.Rows[0]["SearchValue"];
                //gridType=(String)_template.Rows[0]["GridType"];
                switch (parentType.Trim().ToLower())
                {
                    case "window":
                        {
                            wpfapp.GetWPFWindow(parentSearchValue);
                            break;
                        }
                    case "groupbox":
                        {
                            if (_immediateParent.Trim().ToLower() == "window")
                                wpfapp.GetWPFGroupBox(_globalWindow, parentSearchBy, parentSearchValue);

                            else
                                wpfapp.GetWPFGroupBox(_globalGroup, parentSearchBy, parentSearchValue);
                        }
                        break;

                    default:
                        throw new Exception("Window & GroupBox are valid parentType");
                }

                if (gridSearchBy.Trim().ToLower() == "automationid")
                {
                    switch (parentType.Trim().ToLower())
                    {
                        case "window":
                            {
                                dataGrid = wpfapp.GetWPFDataGrid(_globalWindow, gridSearchBy, gridSearchValue);
                                break;
                            }
                        case "groupbox":
                            {
                                //action.GetWPFDataGrid(action., gridSearchValue);
                                break;
                            }
                        default:
                            throw new Exception(_searchTxtAuto);
                    }
                }
                else if (gridSearchBy.Trim().ToLower() == "text")
                {
                    switch (parentType.Trim().ToLower())
                    {
                        case "window":
                            {
                                dataGrid = wpfapp.GetWPFDataGrid(_globalWindow, gridSearchBy, gridSearchValue);
                                break;
                            }
                        case "groupbox":
                            {
                                //t = GetWPFDataGridView(_groupbox, WpfAction.SearchBy.Text, gridSearchValue);
                                break;
                            }
                        default:
                            throw new Exception("Window & GroupBox are valid parentType");
                    }
                }
                if (dataGrid == null)
                {
                    throw new Exception("ListView Not found");
                }
                else
                    return dataGrid;
            }

            catch (SystemException e)
            {
                logTofile(_eLogPtah, "Erroring Line number in getDataGridFromTemplate " + GetStacktrace(e).ToString());
                throw new Exception(e.Message);
            }
        }
        //this function returns the listview from a template sheet
        public void GetActualDataNonUniform(ListView dataGrid)
        {
            try
            {
                if (testData.Template.Rows.Count <= 0)
                {
                    throw new Exception("No data available in template sheet");
                }
                DataTable dt = new DataTable();
                for (int i = 0; i < testData.Template.Rows.Count; i++)
                {
                    if (Convert.IsDBNull(testData.Template.Rows[i]["FieldName"]) == false)
                    {
                        dt.Columns.Add((string)testData.Template.Rows[i]["FieldName"]);
                    }
                }
                DataRow dr;
                dr = dt.NewRow();
                for (int j = 0; j < testData.Template.Rows.Count; j++)
                {
                    try
                    {
                        int _row = int.Parse(testData.Template.Rows[j]["Row"].ToString());
                        string _fieldname = (string)(testData.Template.Rows[j]["FieldName"]);
                        if (dataGrid.Rows[_row].Cells.Count > 0)
                        {
                            dr[j] = dataGrid.Rows[_row].Cells[_fieldname].Text;
                        }
                    }
                    catch (Exception)
                    {
                        throw new Exception(_error + "GetActualDataNonuniform: " + System.Environment.NewLine + " Unable to read the row or column values in template sheet");
                    }
                }
                dt.Rows.Add(dr);
                ActualData = dt;
                testData.ActualData = ActualData;
            }
            catch (Exception ex)
            {
                throw new Exception(_error + "GetActualDataNonUniform: " + System.Environment.NewLine + ex.Message);
            }
        } */
        private bool regexpMatch(string strtext, string strpattern)
        {
            // Instance method:
            try
            {
                Regex reg = new Regex(strpattern);

                // MessageBox.Show(strtext + "--" + strpattern);

                if (reg.IsMatch(strtext))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Erroring Line number in regexpMatch " + GetStacktrace(ex).ToString());
                return false;
                throw new Exception("error on getclickablepoints :" + ex.Message);
            }
        }
        private void ClickControl(AutomationElement control)
        {
            logTofile(_eLogPtah, "ClickControl: outside of try");
            try
            {

                logTofile(_eLogPtah, "Inside Left click method control type is : " + control.Current.LocalizedControlType);
                System.Windows.Point clickpoint1 = control.GetClickablePoint();
                logTofile(_eLogPtah, "Got clickable Points ");
                double x = clickpoint1.X;
                double y = clickpoint1.Y;
                int x1 = Convert.ToInt32(x);
                int y1 = Convert.ToInt32(y);
                logTofile(_eLogPtah, "clickable Points " + x1.ToString() + " " + y1.ToString());
                at.MouseMove(x1, y1, -1);
                try
                {
                    at.MouseClick("LEFT", x1, y1, 1);

                }
                catch (Exception e)
                {
                    logTofile(_eLogPtah, "error on mouseclick :" + e.Message);
                }

            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Erroring Line numner in ClickControl " + GetStacktrace(ex).ToString());
                throw new Exception("error on getclickablepoints :" + ex.Message);

            }
        }
        private void DoubleClickControl(AutomationElement control)
        {
            try
            {
                logTofile(_eLogPtah, "Inside Double click method");
                System.Windows.Point clickpoint1 = control.GetClickablePoint();
                double x = clickpoint1.X;
                double y = clickpoint1.Y;
                int x1 = Convert.ToInt32(x);
                int y1 = Convert.ToInt32(y);
                logTofile(_eLogPtah, "Clickable points " + x1 + " " + y1);
                at.MouseClick("LEFT", x1, y1, 2);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void RightClickControl(AutomationElement control)
        {
            try
            {
                logTofile(_eLogPtah, "Inside Right click method");
                System.Windows.Point clickpoint1 = control.GetClickablePoint();
                double x = clickpoint1.X;
                double y = clickpoint1.Y;
                int x1 = Convert.ToInt32(x);
                int y1 = Convert.ToInt32(y);
                logTofile(_eLogPtah, "Clickable points " + x1 + " " + y1);
                at.MouseClick("RIGHT", x1, y1, 1);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public void setcellvaleps(string rowpos, string colpos, string offset, string strdata, string movenextkey)
        {
            try
            {

                System.Windows.Forms.SendKeys.Flush();
                System.Windows.Forms.SendKeys.SendWait("^{Home}");
                Thread.Sleep(500);
                for (int i = 1; i < Int32.Parse(rowpos); i++)
                {
                    System.Windows.Forms.SendKeys.Flush();
                    System.Windows.Forms.SendKeys.SendWait("{Down}");
                    Thread.Sleep(500);
                }

                for (int i = 1; i < Int32.Parse(colpos); i++)
                {
                    System.Windows.Forms.SendKeys.Flush();
                    System.Windows.Forms.SendKeys.SendWait(movenextkey);
                    Thread.Sleep(500);
                }



                if (strdata != "")
                {
                    if (strdata.Contains("{"))
                    {
                        if (strdata == "{SPACE}")
                        {
                            strdata = " ";
                        }
                    }
                    else
                    {
                        System.Windows.Forms.SendKeys.Flush();
                        System.Windows.Forms.SendKeys.SendWait("{INSERT}");
                        System.Windows.Forms.SendKeys.Flush();
                        System.Windows.Forms.SendKeys.SendWait("+{END}");
                        System.Windows.Forms.SendKeys.Flush();
                        Thread.Sleep(500);
                        System.Windows.Forms.SendKeys.SendWait("+{HOME}");
                        System.Windows.Forms.SendKeys.Flush();
                        Thread.Sleep(500);
                        System.Windows.Forms.SendKeys.SendWait("{DELETE}");
                        Thread.Sleep(500);
                        System.Windows.Forms.SendKeys.Flush();
                        System.Windows.Forms.SendKeys.SendWait(strdata);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);


            }
        }

        public void logTofile(string spath, string stxtMsg)
        {
            if (UseDetaillog == true)
            {
                System.IO.File.AppendAllText(spath, System.DateTime.Now + ":" + stxtMsg + Environment.NewLine);

                try
                {
                    Console.WriteLine(stxtMsg);
                }
                catch (Exception ex)
                {
                    System.Threading.Thread.Sleep(1);
                    throw new Exception(ex.Message);


                }
            }

        }
        private void RunVBS(string vbsFilepath)
        {
            var proc = System.Diagnostics.Process.Start(vbsFilepath);
            proc.WaitForExit();
        }
        private ValuePattern SupportsValuePattern(AutomationElement ae)
        {
            try
            {
                ValuePattern valpat = (ValuePattern)ae.GetCurrentPattern(ValuePattern.Pattern);
                AutomationPattern[] p1 = ae.GetSupportedPatterns();
                logTofile(_eLogPtah, "patterns lenght for ae =" + p1.Length);
                //for (int jj = 0; jj < p1.Length; jj++)
                //{
                // }

                return valpat;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
                //  return null;
            }
        }
        private TableItemPattern SupportsTableItemPattern(AutomationElement ae)
        {
            try
            {
                TableItemPattern talpat = (TableItemPattern)ae.GetCurrentPattern(TableItemPattern.Pattern);
                AutomationPattern[] p1 = ae.GetSupportedPatterns();
                logTofile(_eLogPtah, "patterns lenght for Tableitem patten was e =" + p1.Length);
                //for (int jj = 0; jj < p1.Length; jj++)
                //{
                // }

                return talpat;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
                // return null;
            }
        }
        #region DataTableVerification
        private DataTable GetActualDataForm(string testDataFile)
        {
            logTofile(_eLogPtah, "GetActualDataForm started: Test Data File" + testDataFile);
            string parentType = "";
            string parentSearchBy = "";
            string parentSearchValue = "";
            string controlaction = "";
            var _controlType = "";
            var _logicalName = "";
            var _property = "";
            try
            {

                logTofile(_eLogPtah, "Testdata Tempate Row Count" + testData.Template.Rows.Count.ToString());

                if (testData.Template.Rows.Count <= 0)
                {
                    throw new Exception("No data available in template sheet");
                }
                DataTable dt = new DataTable();
                string colNamesArray = "";
                string colNametoAdd = "";
                for (int i = 0; i < testData.Template.Rows.Count; i++)
                {
                    if (Convert.IsDBNull(testData.Template.Rows[i]["FieldName"]) == false)
                    {

                        colNametoAdd = (string)testData.Template.Rows[i]["FieldName"];

                        logTofile(_eLogPtah, "Column to Add:" + colNametoAdd);

                        if (colNamesArray.Contains(colNametoAdd) == false)
                        {
                            logTofile(_eLogPtah, "Column Added:" + colNametoAdd);

                            colNamesArray = colNamesArray + colNametoAdd + ";";
                            dt.Columns.Add((string)testData.Template.Rows[i]["FieldName"]);
                        }
                    }
                }
                DataRow dr = dt.NewRow();
                logTofile(_eLogPtah, "[ GetActualDataForm]:Columns were added for actual datatable in mememory (for reporting)");
                try
                {
                    #region interationloop
                    string _controlValue = null;
                    for (int i = 0; i < testData.Template.Rows.Count; i++)
                    {
                        parentType = testData.Template.Rows[i]["ParentType"].ToString();
                        parentSearchBy = testData.Template.Rows[i]["ParentSearchBy"].ToString().ToLower();
                        parentSearchValue = testData.Template.Rows[i]["ParentSearchValue"].ToString();
                        controlaction = testData.Template.Rows[i]["ParentSearchValue"].ToString();
                        if (Convert.IsDBNull(testData.Template.Rows[i]["FieldName"]) == false)
                        {
                            _logicalName = (string)testData.Template.Rows[i]["FieldName"].ToString();
                            logTofile(_eLogPtah, "[ GetActualDataForm]:Logical Name to look for:  " + _logicalName);
                        }
                        if (_logicalName.Length > 0)
                        {

                            _controlValue = (string)testData.ExpectedData.Rows[0][_logicalName].ToString();
                            logTofile(_eLogPtah, "[ GetActualDataForm]:ControlValue to Verify from Datasheet was read from datasheet" + _controlValue);
                        }
                        if (IsColumnPresent("Property"))
                        {
                            if (Convert.IsDBNull(testData.Template.Rows[i]["Property"]) == false)
                            {
                                _property = (string)testData.Template.Rows[i]["Property"].ToString();
                                logTofile(_eLogPtah, "[ GetActualDataForm]:Property to look for:  " + _property);
                            }
                        }

                        if ((string)testData.Template.Rows[i]["inputdata"].ToString().ToLower() == "y")
                        {
                            if (!String.IsNullOrEmpty(_controlValue))
                            {
                                #region ParentType
                                if (Convert.IsDBNull(testData.Template.Rows[i]["ParentType"]) == false)
                                {
                                    switch (parentType.Trim().ToLower())
                                    {
                                        case "window":
                                            {
                                                //  wpfapp._application = _application;
                                                //  _globalWindow = wpfapp.GetWPFWindow(parentSearchValue);
                                                //  _globalWindow.Click();

                                                break;
                                            }
                                        case "groupbox":
                                        /*   {
                                               if (_immediateParent.Trim().ToLower() == "window")

                                                   wpfapp.GetWPFGroupBox(_globalWindow, parentSearchBy, parentSearchValue);
                                               else
                                                   wpfapp.GetWPFGroupBox(_globalGroup, parentSearchBy, parentSearchValue);
                                               break;
                                           } */


                                        case "uiautomationwindow":
                                        case "uwindow":
                                            {
                                                uiAutomationWindow = GetUIAutomationWindow(parentSearchBy, parentSearchValue);
                                                break;
                                            }
                                        case "commonhieararchy":
                                            {
                                                string parentH = "";
                                                string parentHvalue = "";
                                                string parentHSearchby = "";
                                                testDataHieararchy.GetTestData(hrchyfile, parentSearchValue);
                                                # region HiearachySheet
                                                for (int ih = 0; ih < testDataHieararchy.Data.Rows.Count; ih++)
                                                {
                                                    parentH = testDataHieararchy.Data.Rows[ih]["Parent"].ToString();
                                                    parentHvalue = testDataHieararchy.Data.Rows[ih]["Value"].ToString();
                                                    parentHSearchby = testDataHieararchy.Data.Rows[ih]["HSearchBy"].ToString();
                                                    switch (parentH.ToString().ToLower())
                                                    {
                                                        case "uiautomationwindow":
                                                            {
                                                                if (uiAutomationWindow == null)
                                                                {
                                                                    uiAutomationWindow = GetUIAutomationWindow(parentHSearchby, parentHvalue);
                                                                }
                                                                else
                                                                {

                                                                    uiAutomationCurrentParent = uiAutomationWindow;
                                                                    logTofile(_eLogPtah, "[GetActualDataForm]->[commonhieararchy]->[uiautomationwindow] : loaded with UIautomaiton window Value");
                                                                    logTofile(_eLogPtah, "[GetActualDataForm]->[commonhieararchy]->[uiautomationwindow] :UI automationwindow was already set and hehce Will not be Reloaded unless you force by some means:COMMONH");
                                                                }
                                                                break;
                                                            }
                                                        case "uiautomationpane":
                                                            {
                                                                GetDescenDentPaneWithName(parentHvalue);
                                                                break;
                                                            }
                                                        case "uiautomationchildpane":
                                                            {
                                                                GetChildPane(Int32.Parse(parentHvalue));
                                                                break;
                                                            }
                                                    }
                                                }
                                                # endregion HiearachySheet
                                                break;
                                            }
                                        case "uiautomationpane":
                                        case "upane":
                                            {
                                                if (_controlValue.Length > 0)
                                                {
                                                    logTofile(_eLogPtah, "[GetActaualDataForm]:Hrchy from main sheet: loking for " + parentSearchValue);
                                                    GetDescenDentPaneWithName(parentSearchValue);
                                                }
                                                break;
                                            }
                                        case "uiautomationchildpane":
                                        case "uchildpane":
                                            {
                                                if (_controlValue.Length > 0)
                                                {
                                                    GetChildPane(Int32.Parse(parentSearchValue));
                                                }
                                                break;
                                            }

                                        case "uiautomationtreeitem":
                                        case "utreeitem":
                                            {
                                                uiAutomationCurrentParent = GetUIAutomationtreeitem(parentSearchBy, parentSearchValue, 0);
                                                break;
                                            }

                                        default:
                                            throw new Exception("Not a valid parent type.");
                                    }
                                }
                                #endregion
                            }

                            var _searchBy = "";
                            var _index = -1;
                            var _controlName = "";

                            if (Convert.IsDBNull(testData.Template.Rows[i]["ControlName"]) == false)
                                _controlName = (string)testData.Template.Rows[i]["ControlName"];


                            if (Convert.IsDBNull(testData.Template.Rows[i]["ControlType"]) == false)
                            {
                                _controlType = (string)testData.Template.Rows[i]["ControlType"].ToString().ToLower();
                                //   _logicalName = (string)testData.Template.Rows[i]["FieldName"].ToString();
                                if (Convert.IsDBNull(testData.Template.Rows[i]["SearchBy"]) == false)
                                    _searchBy = (string)testData.Template.Rows[i]["SearchBy"];

                                Console.WriteLine(_logicalName);

                                if (Convert.IsDBNull(testData.Template.Rows[i]["Index"]) == false)
                                {
                                    _index = int.Parse(testData.Template.Rows[i]["Index"].ToString());
                                }
                                _testcase = (string)testData.ExpectedData.Rows[0]["testcase"].ToString();

                                if (_logicalName.Length > 0)
                                {
                                    logTofile(_eLogPtah, "[ GetActualDataForm]: Trying to Get control Value for " + _logicalName);
                                    //  _controlValue = (string)testData.ExpectedData.Rows[0][_logicalName].ToString();
                                }
                                if (_logicalName.Length > 0 && _controlValue.Length == 0)
                                {
                                    logTofile(_eLogPtah, "[ GetActualDataForm]:" + "Control Value :" + _controlValue);
                                    logTofile(_eLogPtah, "[ GetActualDataForm]: Doing nothing for FieldName (Skiiped):=====> " + _logicalName);
                                }
                                else
                                {
                                    logTofile(_eLogPtah, "[ GetActualDataForm]: Trying to Get control Value for:=======> " + _logicalName + "Control Type was : =================>" + _controlType);
                                    switch (_controlType.Trim().ToLower())
                                    {



                                        /*   case "wpftoolstrip":
                                               wpfapp.GetWPFToolStrip(_globalWindow, _controlName).Focus();
                                               break;

                                           case "wpflistbox":
                                               wpfapp.GetWPFListBox(_globalWindow, _searchBy, _controlName).Focus();
                                               string lstnamecol = "";
                                               int itemscount = wpfapp.GetWPFListBox(_globalWindow, _searchBy, _controlName).Items.Count;
                                               for (int ilt = 0; ilt < itemscount - 1; ilt++)
                                               {
                                                   lstnamecol = lstnamecol + ";" + wpfapp.GetWPFListBox(_globalWindow, _searchBy, _controlName).Items[ilt].Name.ToString();
                                               }
                                               lstnamecol = lstnamecol.Substring(2, lstnamecol.Length);
                                               dr[_logicalName] = (string)lstnamecol;
                                               break;

                                           case "wpflabel":
                                               string labelName = wpfapp.GetWPFLabel(_globalWindow, _searchBy, _controlName).Text.ToString();
                                               testData.UpdateTestData(testData.TestDataFile, testData.TestCase, _logicalName, labelName);
                                               break;

                                           case "wpflistview":
                                               wpfapp.GetWPFDataGrid(_globalWindow, _searchBy, _controlName).Focus();
                                               break;

                                           case "wpfcombobox":
                                               logTofile(_eLogPtah, "before reaching combo ");
                                               wpfapp.GetWPFComboBox(_globalWindow, _searchBy, _controlName, _index).Focus();
                                               int wpfcombocount = wpfapp.GetWPFComboBox(_globalWindow, _searchBy, _controlName, _index).Items.Count;
                                               logTofile(_eLogPtah, "items count inside comob box" + wpfcombocount);
                                               string actvalcmb = wpfapp.GetWPFComboBox(_globalWindow, _searchBy, _controlName, _index).SelectedItem.Name.ToString();
                                               if (actvalcmb == "")
                                               {
                                                   logTofile(_eLogPtah, "Could not read combobox value from White ");
                                               }
                                               dr[_logicalName] = (string)actvalcmb.ToString();
                                               Thread.Sleep(1000);
                                               break; */
                                        case "uiautomationcombobox":
                                        case "ucombobox":
                                            AutomationElement combo = GetUIAutomationComboBox(_searchBy, _controlName, _index);
                                            try
                                            {

                                                combo.SetFocus();
                                                logTofile(_eLogPtah, "Checking for expandcollapse pattern ");
                                                ExpandCollapsePattern expandPat = (ExpandCollapsePattern)combo.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                                if (expandPat != null)
                                                {
                                                    logTofile(_eLogPtah, "Expanding the combobox");
                                                    expandPat.Expand();
                                                    Thread.Sleep(100);
                                                }
                                            }
                                            catch
                                            {
                                            }

                                            try
                                            {
                                                ClickControl(combo);
                                            }
                                            catch
                                            {
                                            }

                                            Thread.Sleep(1000);

                                            //Control value is item to select
                                            AutomationElementCollection comboitems = combo.FindAll(TreeScope.Children,
                                                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.ListItem));
                                            logTofile(_eLogPtah, "[GetActualDataForm][iautomationcombobox]:items cout  veryifing " + _controlValue + "is :" + comboitems.Count);
                                            for (int icb = 0; icb <= comboitems.Count - 1; icb++)
                                            {
                                                logTofile(_eLogPtah, "[GetActualDataForm][iautomationcombobox]:checkingfor items selection tru or not" + comboitems[icb].Current.Name.ToString());
                                                SelectionItemPattern selpat = (SelectionItemPattern)comboitems[icb].GetCurrentPattern(SelectionItemPattern.Pattern);
                                                if (selpat.Current.IsSelected == true)
                                                {
                                                    logTofile(_eLogPtah, "Selected item in Combobox was :" + comboitems[icb].Current.Name.ToString());

                                                    if (TreeWalker.ControlViewWalker.GetFirstChild(comboitems[icb]) != null)
                                                    {
                                                        AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(comboitems[icb]);
                                                        string _controltype = elementNode.Current.LocalizedControlType.ToString();
                                                        logTofile(_eLogPtah, elementNode.Current.Name.ToLower() + ":" + _controlValue.ToLower());
                                                        if (_controltype.ToLower() == "text" && elementNode.Current.Name.ToLower() == _controlValue.ToLower())
                                                        {
                                                            logTofile(_eLogPtah, "[GetActualDataForm][uiautomationcombobox]:match found for -->" + _controlValue.ToLower());
                                                            dr[_logicalName] = _controlValue;
                                                            uilog.AddTexttoColumn("Action Performed on Control", "Value has been read from text node " + _controlValue);
                                                            break;
                                                        }

                                                    }
                                                    else
                                                    {
                                                        dr[_logicalName] = (string)comboitems[icb].Current.Name.ToString();
                                                        uilog.AddTexttoColumn("Action Performed on Control", "Value has been read by selection pattern " + _controlValue);
                                                        break;
                                                    }
                                                }
                                                #region old code
                                                //if (comboitems[icb].Current.Name.ToLower() == _controlValue.ToLower()) //if listitemname matches 
                                                //{

                                                //    SelectionItemPattern selpat = (SelectionItemPattern)comboitems[icb].GetCurrentPattern(SelectionItemPattern.Pattern);
                                                //    if (selpat.Current.IsSelected == true)
                                                //    {
                                                //        logTofile(_eLogPtah, "Selected item in Combobox was :" + comboitems[icb].Current.Name.ToString());
                                                //        dr[_logicalName] = (string)comboitems[icb].Current.Name.ToString();
                                                //        uilog.AddTexttoColumn("Action Performed on Control", "Value has been read by selection pattern " + _controlValue);
                                                //    }

                                                //}
                                                #endregion old code
                                                #region checking with text of the listitem
                                                //else if (TreeWalker.ControlViewWalker.GetFirstChild(comboitems[icb]) != null)
                                                //{
                                                //    AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(comboitems[icb]);
                                                //    string _controltype = elementNode.Current.System.Windows.Automation.ControlType.LocalizedControlType.ToString();
                                                //    logTofile(_eLogPtah, elementNode.Current.Name.ToLower() + ":" + _controlValue.ToLower());
                                                //    if (_controltype.ToLower() == "text" && elementNode.Current.Name.ToLower() == _controlValue.ToLower())
                                                //    {
                                                //        // ClickControl(elementNode);
                                                //        logTofile(_eLogPtah, "[GetActualDataForm][uiautomationcombobox]:match found for -->" + _controlValue.ToLower());

                                                //        dr[_logicalName] = _controlValue;
                                                //        uilog.AddTexttoColumn("Action Performed on Control", "Value has been read from text node " + _controlValue);
                                                //        break;
                                                //    }
                                                //}
                                                #endregion checking with text of the listitem

                                            }
                                            #region clickCombo
                                            //try
                                            //{
                                            //    ClickControl(combo);
                                            //}
                                            //catch
                                            //{
                                            //}
                                            #endregion clickcombo
                                            break;
                                        /*  case "wpfcheckbox":
                                              string actvalchk = (string)wpfapp.GetWPFCheckBox(_globalWindow, _searchBy, _controlName, _index).IsSelected.ToString();
                                              dr[_logicalName] = (string)actvalchk;
                                              break; */

                                        case "uiautomationcheckbox":
                                        case "ucheckbox":
                                            logTofile(_eLogPtah, "Verifying the toolge state for checkboxes");
                                            logTofile(_eLogPtah, "[GetActualDataForm][uiautomationcheeckbox]: for " + _logicalName);
                                            AutomationElement uiCheckBox = GetUIAutomationCheckBox(_searchBy, _controlName, _index);
                                            logTofile(_eLogPtah, "[GetActualDataForm][uiautomationcheeckbox]: got checkbox  " + _logicalName);
                                            try
                                            {
                                                TogglePattern togPattern = (TogglePattern)uiCheckBox.GetCurrentPattern(TogglePattern.Pattern);
                                                logTofile(_eLogPtah, "Checkbox value is -->" + togPattern.Current.ToggleState.ToString());
                                                dr[_logicalName] = (string)togPattern.Current.ToggleState.ToString();
                                            }
                                            catch (Exception ex)
                                            {
                                                logTofile(_eLogPtah, "Exception in Toggle patern" + ex.Message.ToString());
                                            }
                                            break;

                                        case "uiautomationradiobutton":
                                        case "uradiobutton":
                                            AutomationElement uiRadio = GetUIAutomationRadioButton(_searchBy, _controlName, _index);
                                            SelectionItemPattern selpatt = (SelectionItemPattern)uiRadio.GetCurrentPattern(SelectionItemPattern.Pattern);
                                            logTofile(_eLogPtah, "RadioButton Button value is -->" + selpatt.Current.IsSelected.ToString());
                                            dr[_logicalName] = (string)selpatt.Current.IsSelected.ToString();
                                            break;

                                        /*   case "wpfbutton":
                                               wpfapp.GetWPFButton(_globalWindow, _searchBy, _controlName, _index).Click();
                                               break;

                                           case "wpfradiobutton":
                                               bool actvalrb = wpfapp.GetWPFRadioButton(_globalWindow, _searchBy, _controlName, _index).IsSelected;

                                               dr[_logicalName] = (string)actvalrb.ToString();
                                               logTofile(_eLogPtah, "Inseretd Radio button value to Datable also for " + _logicalName + "value isneted was " + _controlValue);
                                               Thread.Sleep(1000);
                                               break;

                                           case "wpftextbox":
                                               if (_controlValue != DBNull.Value.ToString())
                                               {
                                                   logTofile(_eLogPtah, "Trying to Verify TextBox" + _logicalName);
                                                   string acttext = wpfapp.GetWPFTextBox(_globalWindow, _searchBy, _controlName, _index).Text.ToString();
                                                   dr[_logicalName] = (string)acttext;
                                                   logTofile(_eLogPtah, "Text box value :" + _logicalName);

                                               }
                                               break; */
                                        case "uiautomationedit":
                                        case "uedit":
                                            if (_controlValue != DBNull.Value.ToString())
                                            {
                                                logTofile(_eLogPtah, "[GetActualDataForm]:[uiautomationedit]:Trying to Verify UIautomationedit" + _logicalName);
                                                AutomationElement ueditbox = GetUIAutomationEdit(_searchBy, _controlName, _index);
                                                ValuePattern valpat = (ValuePattern)ueditbox.GetCurrentPattern(ValuePattern.Pattern);
                                                string acttext = valpat.Current.Value;
                                                dr[_logicalName] = (string)acttext;

                                            }
                                            break;
                                        case "uiautomationtext":
                                        case "utext":
                                            {
                                                if (_controlValue.Length > 0)
                                                {
                                                    logTofile(_eLogPtah, "[GetActualDataForm]:[uiautomationtext]:Trying to Verify UIautomationtext" + _logicalName);
                                                    AutomationElement utext = GetUIAutomationText(_searchBy, _controlName, _index);
                                                    string acttext = utext.Current.Name.ToString();
                                                    dr[_logicalName] = (string)acttext;
                                                }
                                                break;
                                            }

                                        /*    case "maskedwpftextbox":
                                                wpfapp.GetWPFTextBox(_globalWindow, _searchBy, _controlName, _index);
                                                Console.WriteLine("inside maskedwpftextbox");
                                                var charArray = _controlValue.Select(q => new string(q, 1)).ToArray();
                                                for (int k = 0; k < charArray.Count(); k++)
                                                {
                                                    System.Windows.Forms.SendKeys.Flush();
                                                    System.Windows.Forms.SendKeys.SendWait(charArray[k]);
                                                }
                                                break; */
                                        /*  case "wpfmultilinetextbox":
                                              logTofile(_eLogPtah, "[GetActualDataForm][wpfmultilinetextbox]:-->lookinng for " + _logicalName);

                                              if (_controlValue != DBNull.Value.ToString())
                                              {
                                                  SearchCriteria _tbsearchcriteria = SearchCriteria.ByText(_controlName);
                                                  if (_globalWindow == null)
                                                  {
                                                      logTofile(_eLogPtah, " [GetActualDataForm]:[case::->wpfmultilinetextbox]- >global window was not set ..trying to set it  first ");
                                                      GetUIAutomationWindow(_searchBy, _controlName);
                                                  }
                                                  logTofile(_eLogPtah, " [GetActualDataForm]:[case::->wpfmultilinetextbox]- >global window was not set ..trying to set it  first ");

                                                  var tbeditor = _globalWindow.Get<MultilineTextBox>(_tbsearchcriteria);
                                                  dr[_logicalName] = (string)tbeditor.Text;
                                                  logTofile(_eLogPtah, " [GetActualDataForm]:[case::->wpfmultilinetextbox]- >Value was inserted to datatable");


                                              }
                                              break; */
                                        case "uiautomationtextarea":
                                        case "utextarea":
                                            if (_controlValue.Length > 0)
                                            {
                                                logTofile(_eLogPtah, "[GetActualDataForm][uiautomationtextarea]:-->Hunting for: " + _logicalName);
                                                AutomationElement textarea = GetUIAutomationTextarea(_searchBy, _controlName, _index);
                                                if (textarea == null)
                                                {
                                                    logTofile(_eLogPtah, "For Control " + _logicalName + "Control was not detected Failure to find control");
                                                }
                                                else
                                                {
                                                    logTofile(_eLogPtah, "For Control " + _logicalName + "Control was detected:Success");
                                                }
                                                TextPattern txtptn = (TextPattern)textarea.GetCurrentPattern(TextPattern.Pattern);

                                                string rtt = txtptn.DocumentRange.GetText(10000);
                                                logTofile(_eLogPtah, " [GetActualDataForm]:[case::->uiautomationtextarea]- >Text was Read from  " + _logicalName + "text read was :" + rtt);
                                                // string ttval = txtptn.DocumentRange.GetAttributeValue(AutomationTextAttribute.
                                                // string rtt = (string)textarea.GetCurrentPropertyValue(ValuePattern.ValueProperty);
                                                dr[_logicalName] = (string)rtt;
                                                logTofile(_eLogPtah, " [GetActualDataForm]:[case::->uiautomationtextarea]- >Value got inserted to datatable for " + _logicalName);
                                                Thread.Sleep(1000);
                                            }
                                            break;
                                        case "uiautomationcustominvokecontrol":
                                        case "ucustominvokecontrol":
                                            AutomationElement customcontrol = GetUIAutomationCustominvokecontrol(_searchBy, _controlName, _index);
                                            InvokePattern invkptn = (InvokePattern)customcontrol.GetCurrentPattern(InvokePattern.Pattern);
                                            if (_controlValue.ToLower() == "y")
                                            {
                                                invkptn.Invoke();
                                            }

                                            break;
                                        case "uiautomationlistitems":
                                        case "ulistitems":
                                            string uilist = "";
                                            uilist = GetUIAutomationListContent(_searchBy, _controlName, _index);
                                            dr[_logicalName] = (string)uilist;
                                            break;
                                        case "ulistitemcollection":
                                            string uilistitems = "";
                                            uilistitems = GetUIAutomationListCollection(_searchBy, _controlName, _index);
                                            dr[_logicalName] = (string)uilistitems;
                                            break;
                                        case "uiautomationtreeitems":
                                            string uitrees = GetUIAutomationTreeContent(_searchBy, _controlName, _index);
                                            dr[_logicalName] = (string)uitrees;
                                            break;

                                        case "uiautomationtreeitem":
                                        case "utreeitem":
                                            if (_controlValue.Length > 0)
                                            {
                                                string actval = null;
                                                AutomationElement uitreeitem = GetUIAutomationtreeitem(_searchBy, _controlName, _index);
                                                switch (_searchBy.ToLower())
                                                {
                                                    case "name":
                                                    case "text":
                                                        actval = uitreeitem.Current.Name.ToString();
                                                        break;
                                                    case "automationid":
                                                        actval = uitreeitem.Current.AutomationId.ToString();
                                                        break;
                                                    default:
                                                        logTofile(_eLogPtah, "invalid searchby criteria");
                                                        break;
                                                }
                                                dr[_logicalName] = (string)actval;
                                            }
                                            break;

                                        case "uiautomationselectultratabitem":
                                        case "uselectultratabitem":
                                            AutomationElement ultratabitemSelect = GetUIAutomationUltratab(_searchBy, _controlName, _index);
                                            logTofile(_eLogPtah, "[GetActualDataForm]->[uiautomationselectultratabitem] Ultra Tab Name" + ultratabitemSelect.Current.Name);

                                            try
                                            {
                                                if (ultratabitemSelect.Current.IsKeyboardFocusable == true)
                                                {
                                                    ultratabitemSelect.SetFocus();
                                                }

                                                SelectionItemPattern selpat = (SelectionItemPattern)ultratabitemSelect.GetCurrentPattern(SelectionItemPattern.Pattern);
                                                selpat.Select();
                                                uilog.AddTexttoColumn("Action Performed on Control", "Clicked using Selection Pattern " + _controlName);
                                                logTofile(_eLogPtah, "Clicked using Selection Pattern ");

                                            }
                                            catch (Exception e)
                                            {
                                                logTofile(_eLogPtah, "Error encountered in uiautomationselectultratabitem: " + e.Message);
                                                throw new Exception("uiautomationselectultratabitem" + e.Message);
                                            }
                                            break;
                                        case "uiautomationultratabitem":
                                        case "uultratabitem":
                                            AutomationElement ultratabitem = GetUIAutomationUltratab(_searchBy, _controlName, _index);
                                            logTofile(_eLogPtah, "[VerifyData]->[uiautomationultratabitem]" + _searchBy + "search value" + _controlName + "index" + _index);
                                            AutomationElement objbuttonnext = null;
                                            if (objbuttonnext == null)
                                            {
                                                try
                                                {
                                                    logTofile(_eLogPtah, " Current Parent : " + uiAutomationCurrentParent.Current.Name.ToString());
                                                    if (ultratabitem.Current.IsKeyboardFocusable == true)
                                                    {
                                                        ultratabitem.SetFocus();
                                                    }
                                                    ClickControl(ultratabitem);
                                                }
                                                catch (Exception e)
                                                {
                                                    logTofile(_eLogPtah, "Error encountered :" + e.Message);
                                                    throw new Exception(e.Message);
                                                }
                                            }
                                            else
                                            // need to detect until no error
                                            {
                                                try
                                                {
                                                    ClickControl(ultratabitem);
                                                }
                                                catch (Exception e)
                                                {
                                                    logTofile(_eLogPtah, "Error encountered :" + e.Message);
                                                    throw new Exception(e.Message);
                                                }
                                            }
                                            break;

                                        case "uiautomationbutton":
                                        case "ubutton":
                                            if (uiAutomationWindow.Current.Name.Length == 0)
                                            {
                                                uiAutomationWindow = GetUIAutomationWindow(parentSearchBy, parentSearchValue);
                                            }
                                            AutomationElement button = GetUIAutomationbutton(_searchBy, _controlName, _index);
                                            string bname = button.Current.Name.ToString();
                                            logTofile(_eLogPtah, "[AddData][uiautomationbutton]:Retrived button contol  : " + bname);

                                            InvokePattern invkbuttonptn = (InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern);
                                            if (Convert.IsDBNull(_controlValue) == false)

                                                if (Int32.Parse(_controlValue) > 0)
                                                {
                                                    for (int ib = 0; ib < Int32.Parse(_controlValue); ib++)
                                                    {

                                                        try
                                                        {
                                                            invkbuttonptn.Invoke();
                                                            System.Threading.Thread.Sleep(20);
                                                            logTofile(_eLogPtah, "[AddData][uiautomationbutton]:Pressed button : " + bname);
                                                            dr[_logicalName] = _controlValue;
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            logTofile(_eLogPtah, " Error Encountered: " + ex.Message);
                                                            throw new Exception(ex.Message);
                                                        }
                                                    }
                                                }

                                            break;

                                        case "uiautomationbuttonverify":
                                        case "ubuttonverify":
                                            AutomationElement uiautomationbutton_verify = GetUIAutomationbutton(_searchBy, _controlName, _index);
                                            logTofile(_eLogPtah, _property);
                                            switch (_property.Trim().ToLower())
                                            {
                                                case "isenabled":
                                                    {
                                                        actresult_K2 = ((bool)uiautomationbutton_verify.GetCurrentPropertyValue(AutomationElement.IsEnabledProperty));
                                                        logTofile(_eLogPtah, "actresult value is : " + actresult_K2);
                                                        if (actresult_K2.ToString() == "True")
                                                            logTofile(_eLogPtah, "Button is enabled");
                                                        else
                                                            logTofile(_eLogPtah, "Button is disabled");
                                                        dr[_logicalName] = actresult_K2.ToString();
                                                        break;
                                                    }
                                                case "name":
                                                case "text":
                                                    {
                                                        string actval = null;
                                                        actval = uiautomationbutton_verify.Current.Name.ToString();
                                                        logTofile(_eLogPtah, "Name of a button :" + actval);
                                                        dr[_logicalName] = actval;
                                                        break;
                                                    }
                                                case "togglestate":
                                                    {
                                                        string actval = null;
                                                        TogglePattern togPattern = (TogglePattern)uiautomationbutton_verify.GetCurrentPattern(TogglePattern.Pattern);
                                                        actval = togPattern.Current.ToggleState.ToString();
                                                        logTofile(_eLogPtah, "Toggle state of a button: " + actval);
                                                        dr[_logicalName] = actval;
                                                        break;
                                                    }


                                            }
                                            break;

                                        case "uiautomationtabverify":
                                        case "utabverify":

                                            if (_controlValue.Length > 0)
                                            {
                                                string actval = null;
                                                AutomationElement uiautomationtab_verify = GetUIAutomationUltratab(_searchBy, _controlName, _index);
                                                switch (_searchBy.ToLower())
                                                {
                                                    case "name":
                                                    case "text":
                                                        actval = uiautomationtab_verify.Current.Name.ToString();
                                                        break;
                                                    case "automationid":
                                                        actval = uiautomationtab_verify.Current.AutomationId.ToString();
                                                        break;
                                                    default:
                                                        break;

                                                }
                                                dr[_logicalName] = (string)actval;
                                            }
                                            logTofile(_eLogPtah, "Verifying Proprrty Value:");
                                            logTofile(_eLogPtah, "looikng for Property Value" + _property);
                                            logTofile(_eLogPtah, "Was Column Property present in strcuture sheet:" + IsColumnPresent(_property));
                                            if (IsColumnPresent("Property"))
                                            {
                                                AutomationElement uiautomationtab_verify = GetUIAutomationUltratab(_searchBy, _controlName, _index);
                                                Boolean attrresult;
                                                switch (_property.Trim().ToLower())
                                                {
                                                    case "isenabled":
                                                        {
                                                            attrresult = ((bool)uiautomationtab_verify.GetCurrentPropertyValue(AutomationElement.IsEnabledProperty));
                                                            logTofile(_eLogPtah, "actresult value is : " + attrresult);
                                                            if (attrresult.ToString() == "True")
                                                                logTofile(_eLogPtah, "tab is enabled");
                                                            else
                                                                logTofile(_eLogPtah, "tab is disabled");
                                                            dr[_logicalName] = attrresult.ToString();
                                                            break;
                                                        }


                                                    default:
                                                        {
                                                            logTofile(_eLogPtah, "No Attributes are supplied for verification: ");
                                                            break;
                                                        }
                                                }
                                            }

                                            break;

                                        case "uiautomationribbonbutton":
                                        case "uribbonbutton":
                                            if (uiAutomationWindow.Current.Name.Length == 0)
                                            {
                                                uiAutomationWindow = GetUIAutomationWindow(parentSearchBy, parentSearchValue);
                                            }
                                            AutomationElement ribbonbutton = GetUIAutomationRibbonButton(_searchBy, _controlName, _index);
                                            logTofile(_eLogPtah, "[AddData][uiautomationribbonbutton]:Retrived Ribbonbutton contol  : ");
                                            ClickControl(ribbonbutton);
                                            System.Threading.Thread.Sleep(20);
                                            logTofile(_eLogPtah, "[AddData][uiautomationribbonbutton]:Pressed Ribbonbutton   : " + _logicalName);


                                            break;

                                        case "uiautomationtreeitemclick":
                                        case "utreeitemclick":
                                            if (_controlValue.Length > 0)
                                            {
                                                _controlName = _controlValue;
                                            }
                                            AutomationElement treeitem = GetUIAutomationtreeitem(_searchBy, _controlName, _index);
                                            ClickControl(treeitem);
                                            break;
                                        case "uiautomationtreeitemexpand":
                                        case "utreeitemexpand":
                                            if (_controlValue.Length > 0)
                                            {
                                                _controlName = _controlValue;
                                            }
                                            AutomationElement treeitemcolapsed = GetUIAutomationtreeitem(_searchBy, _controlName, _index);
                                            DoubleClickControl(treeitemcolapsed);
                                            break;


                                        case "uiautomationsyncfusionpane":
                                        case "usyncfusionpane":
                                            AutomationElement syncfusionpane = GetUIAutomationsyncfusionpane(_searchBy, _controlName, _index);
                                            //   syncfusionpane.SetFocus();this mehtod dud not work hence using clickable point

                                            ClickControl(syncfusionpane);
                                            Console.WriteLine("found syncfusion pane ");

                                            char[] celldellim = new char[] { ';' };
                                            string[] arr = _controlValue.Split(celldellim);
                                            string testdataPath = arr[0];
                                            string paramfiletcase = arr[1];
                                            string expectedfileName = arr[2];
                                            string tcase = arr[3];

                                            action3.UpdateTestData(testdataPath + @"paramFile.xls", paramfiletcase, "expectedFile", testdataPath + expectedfileName);
                                            action3.UpdateTestData(testdataPath + @"paramFile.xls", paramfiletcase, "testcaseID", tcase);
                                            action3.UpdateTestData(testdataPath + @"paramFile.xls", paramfiletcase, "tempFilePath", @"C:\created.txt");
                                            RunVBS(testdataPath + @"\VerifySyncFusionGridData.vbs");
                                            break;


                                        /*    case "wpfdatepicker":
                                                if (_controlValue != DBNull.Value.ToString())
                                                {
                                                    wpfapp.GetWPFDatePicker(_globalWindow, _searchBy, _controlName).SetValue(_controlValue);
                                                }
                                                break; */

                                        /*  case "wpftabitem":
                                              if (_controlValue != DBNull.Value.ToString())
                                              {
                                                  wpfapp.GetWPFTabItem(_globalWindow, _searchBy, _controlName).Select();
                                              }
                                              break;

                                          case "wpftree":
                                              if (_controlValue != DBNull.Value.ToString())
                                              {
                                                  wpfapp.GetWPFTreeViewWindow(_globalWindow, _searchBy, _controlName, _index).SetValue(_controlValue);
                                              }
                                              break; */

                                        case "uiautomationdialogtextcontrol":
                                        case "udialogtextcontrol":
                                            if (_controlValue != DBNull.Value.ToString())
                                            {
                                                logTofile(_eLogPtah, "Inside dialgo text conctrol");
                                                AutomationElement dlgtxtcontrol = GetUIAutomationDialogTextControl(_searchBy, _controlName, _index);
                                                logTofile(_eLogPtah, "got dialgo text conctrol");
                                                string acttext = dlgtxtcontrol.Current.Name.ToString();
                                                logTofile(_eLogPtah, "[uiautomationdialogtextcontrol]: got dialgo text conctrol");
                                                logTofile(_eLogPtah, "got value: " + acttext);
                                                dr[_logicalName] = (string)acttext;
                                                logTofile(_eLogPtah, "added data to data table from now on helper should take care");
                                            }
                                            break;
                                        case "uicustomcalendarmatbal":
                                        case "ucustomcalendarmatbal":
                                            {
                                                if (_controlValue.Length > 0)
                                                {
                                                    //AutomationElement showcalbutton = GetUIAutomationbutton(_searchBy, _controlName, _index);
                                                    //inputDateInCalendarControl1(showcalbutton, _controlValue);

                                                }
                                                break;
                                            }
                                        case "uiautomationwindow":
                                        case "uwindow":
                                            {
                                                if (_controlValue.Length > 0)
                                                {
                                                    string actval = null;
                                                    AutomationElement uiwindow = GetUIAutomationWindow(_searchBy, _controlName);
                                                    switch (_searchBy.ToLower())
                                                    {
                                                        case "title":
                                                            actval = uiwindow.Current.Name.ToString();
                                                            break;
                                                        case "automationid":
                                                            actval = uiwindow.Current.AutomationId.ToString();
                                                            break;
                                                        default:
                                                            break;

                                                    }
                                                    dr[_logicalName] = (string)actval;
                                                }

                                                break;
                                            }

                                        default:
                                            throw new Exception("Not a valid control type.");

                                    } // Switch                                     

                                }

                            }

                        }


                    }
                    #endregion

                    dt.Rows.Add(dr);
                    logTofile(_eLogPtah, "GetActualDataForm completed");
                    return dt;

                }

                catch (Exception ex)
                {
                    logTofile(_eLogPtah, "Genreic error in any of GetActaulDataForm Case: " + ex.Message.ToString());
                    logTofile(_eLogPtah, "[getActualdataform] Line number : " + GetStacktrace(ex).ToString());
                    throw new Exception("Genreic Error in any of GetActaulDataForm Case");
                    //  return null;

                }



            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Genreic Error in any of GetActaulDataForm Case: " + ex.Message.ToString());
                logTofile(_eLogPtah, "[getActualdataform] Line number : " + GetStacktrace(ex).ToString());
                // return null;
                throw new Exception(_error + "GetActualData:" + System.Environment.NewLine + ex.Message);
            }
        }
        private DataTable GetActualDataIEPaneTable(string testDataFile)
        {
            string TableType = "";
            string SearchBy = "";
            string SearchValue = "";
            int tindex = -1;
            var _logicalName = "";//coulmnname/fieldname  alias 
            DataTable dt = new DataTable();
            try
            {
                if (testData.Template.Rows.Count <= 0)
                {
                    throw new Exception("No data available in template sheet");
                }

                DataRow dr = dt.NewRow();
                TableType = testData.Template.Rows[0]["TableType"].ToString();
                SearchBy = testData.Template.Rows[0]["SearchBy"].ToString().ToLower();
                SearchValue = testData.Template.Rows[0]["SearchValue"].ToString();
                if (SearchBy.ToLower() == "index")
                {
                    tindex = Int32.Parse(SearchValue);
                }
                for (int i = 0; i < testData.Template.Rows.Count; i++)
                {
                    if (Convert.IsDBNull(testData.Template.Rows[i]["FieldName"]) == false)
                    {
                        dt.Columns.Add((string)testData.Template.Rows[i]["FieldName"]);
                    }
                }
                int cellCountStart = 0;
                int colcount = testData.Template.Rows.Count;
                AutomationElement iepanetable = GetUIAutomationIEPaneTable("", "", tindex);
                AutomationElementCollection cells = iepanetable.FindAll(TreeScope.Children,
                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom));
                System.Console.WriteLine("Cells count  Inside targeted table  = " + cells.Count.ToString());
                logTofile(_eLogPtah, "[GetActualDataIEPaneTable]Cells count  Inside targeted table  = " + cells.Count.ToString());
                var ieVersion = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Internet Explorer").GetValue("Version");
                logTofile(_eLogPtah, "IE verrsion For WellFlo Reports and mini reports is" + ieVersion);
                if (GetOSArchitecture() == "64")
                {
                    logTofile(_eLogPtah, "Os is 64 bit so headers start from zero----questionable");
                    if (ieVersion.ToString().IndexOf("9.0") != -1) // IE 9.0 browser
                    {
                        logTofile(_eLogPtah, "Browser is Ie 9.0 ");
                        cellCountStart = colcount;
                    }
                    else if (ieVersion.ToString().IndexOf("10.0") != -1) // IE 10.0 browser
                    {
                        logTofile(_eLogPtah, "Browser is Ie 10.0 ");
                        cellCountStart = 0;
                    }
                    else                       // IE 8.0  or older browser
                    {
                        logTofile(_eLogPtah, "Browser is Ie 8.0 or older version hopefuly not IE 11 !!! ");
                        cellCountStart = colcount;
                    }

                }
                else
                {
                    cellCountStart = colcount;
                }
                int ir;
                dr = dt.NewRow();

                for (int ip = cellCountStart; ip < cells.Count; ip++)
                {
                    AutomationElement txt = cells[ip].FindFirst(TreeScope.Descendants,
                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Text));
                    int resutl = Math.DivRem(ip, colcount, out ir);

                    _logicalName = (string)testData.Template.Rows[ir]["FieldName"].ToString();
                    if (txt != null)
                    {
                        string val = txt.Current.Name;
                        System.Console.WriteLine("Cells value for cell " + ip + " = " + val);
                        dr[_logicalName] = (string)val;
                    }
                    else
                    {
                        string val = "";
                        System.Console.WriteLine("Cells value for cell " + ip + " = " + val);
                        dr[_logicalName] = (string)val;
                    }

                    if (ir == colcount - 1)
                    {

                        dt.Rows.Add(dr);
                        dr = dt.NewRow();

                    }
                }



                return dt;
            }

            catch (Exception ex)
            {
                System.Console.WriteLine("Execption Message : " + ex.Message.ToString());

                logTofile(_eLogPtah, "[GetActualDataIEPaneTable] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message.ToString());
            }




        }
        private DataTable GetActualDataGridContent(string testDataFile)
        {
            string TableType = "";
            string SearchBy = "";
            string SearchValue = "";
            int tindex = -1;
            var _logicalName = "";//coulmnname/fieldname  alias 
            DataTable dtwpf = new DataTable();
            try
            {
                if (testData.Template.Rows.Count <= 0)
                {
                    throw new Exception("No data available in template sheet");
                }
                DataRow dr = dtwpf.NewRow();
                TableType = testData.Template.Rows[0]["TableType"].ToString();
                SearchBy = testData.Template.Rows[0]["SearchBy"].ToString().ToLower();
                SearchValue = testData.Template.Rows[0]["SearchValue"].ToString();
                if (SearchBy.ToLower() == "index")
                {
                    tindex = Int32.Parse(SearchValue);
                }
                for (int i = 0; i < testData.Template.Rows.Count; i++)
                {
                    if (Convert.IsDBNull(testData.Template.Rows[i]["FieldName"]) == false)
                    {
                        dtwpf.Columns.Add((string)testData.Template.Rows[i]["FieldName"]);
                    }
                }
                AutomationElement datagridtable = GetUIAutomationDataGrid(SearchBy, SearchValue, tindex);
                if (datagridtable != null)
                {
                    logTofile(_eLogPtah, "[GetActualDataGridContent]:Datagrid Object has been detected!");
                }
                AutomationElementCollection datarows = datagridtable.FindAll(TreeScope.Children,
                        new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataItem));
                logTofile(_eLogPtah, "[GetActualDataGridContent]: DataGrid Rows count is =" + datarows.Count.ToString());
                dr = dtwpf.NewRow();
                for (int ip = 0; ip < datarows.Count; ip++)
                {
                    AutomationElementCollection cells = datarows[ip].FindAll(TreeScope.Children,
                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom));
                    logTofile(_eLogPtah, "[GetActualDataGridContent]: Cells or columns count " + cells.Count);
                    for (int i = 0; i < testData.Template.Rows.Count; i++)
                    {

                        logTofile(_eLogPtah, "[GetActualDataGridContent]: Trying Value patern " + i);
                        ValuePattern valpat = (ValuePattern)cells[i].GetCurrentPattern(ValuePattern.Pattern);
                        _logicalName = testData.Template.Rows[i]["FieldName"].ToString();
                        logTofile(_eLogPtah, "[GetActualDataGridContent]: Field: " + _logicalName);
                        dr[_logicalName] = (string)valpat.Current.Value;
                        logTofile(_eLogPtah, "[GetActualDataGridContent]: Read Value: " + (string)valpat.Current.Value + " for Row:" + i.ToString());

                    }
                    dtwpf.Rows.Add(dr);
                    dr = dtwpf.NewRow();
                }
                return dtwpf;
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[GetActualDataGridContent]:Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[GetActualDataGridContent] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message.ToString());


                // return null;
            }

        }
        /// <summary>
        /// This function is to verify the grid in which cell does not have value but its immediate child has . Also child element's control type is not uniform
        /// </summary>
        /// <param name="testDataFile"></param>
        /// <returns></returns>

        private DataTable GetActualDataGrid2Content(string testDataFile)
        {
            string TableType = "";
            string SearchBy = "";
            string SearchValue = "";
            int tindex = -1;
            var _logicalName = "";//coulmnname/fieldname  alias 
            DataTable dtwpf = new DataTable();
            try
            {
                if (testData.Template.Rows.Count <= 0)
                {
                    throw new Exception("No data available in template sheet");
                }
                DataRow dr = dtwpf.NewRow();
                TableType = testData.Template.Rows[0]["TableType"].ToString();
                SearchBy = testData.Template.Rows[0]["SearchBy"].ToString().ToLower();
                SearchValue = testData.Template.Rows[0]["SearchValue"].ToString();
                if (SearchBy.ToLower() == "index")
                {
                    tindex = Int32.Parse(SearchValue);
                }
                for (int i = 0; i < testData.Template.Rows.Count; i++)
                {
                    if (Convert.IsDBNull(testData.Template.Rows[i]["FieldName"]) == false)
                    {
                        dtwpf.Columns.Add((string)testData.Template.Rows[i]["FieldName"]);
                    }
                }
                AutomationElement datagridtable = GetUIAutomationDataGrid(SearchBy, SearchValue, tindex);
                if (datagridtable != null)
                {
                    logTofile(_eLogPtah, "[GetActualDataGridContent]:Datagrid Object has been detected!");
                }
                AutomationElementCollection datarows = datagridtable.FindAll(TreeScope.Children,
                        new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.DataItem));
                logTofile(_eLogPtah, "[GetActualDataGridContent]: DataGrid Rows count is =" + datarows.Count.ToString());
                dr = dtwpf.NewRow();
                for (int ip = 0; ip < datarows.Count; ip++)
                {
                    AutomationElementCollection cells = datarows[ip].FindAll(TreeScope.Children,
                    new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom));
                    logTofile(_eLogPtah, "[GetActualDataGridContent]: Cells or columns count " + cells.Count);

                    for (int i = 0; i < cells.Count; i++)
                    {
                        logTofile(_eLogPtah, "Current coulmn: " + i);
                        logTofile(_eLogPtah, "[GetActualDataGridContent]: Trying Value patern " + i);

                        ValuePattern valpat = (ValuePattern)cells[i].GetCurrentPattern(ValuePattern.Pattern);
                        if ((string)valpat.Current.Value == "")
                        {

                            AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(cells[i]);
                            if (elementNode != null)
                            {
                                string _controltype = elementNode.Current.LocalizedControlType.ToString(); ;
                                logTofile(_eLogPtah, "Control type: " + _controltype);
                                switch (_controltype.ToLower())
                                {
                                    case "check box":
                                        TogglePattern togPattern = (TogglePattern)elementNode.GetCurrentPattern(TogglePattern.Pattern);
                                        logTofile(_eLogPtah, "Checkbox value is -->" + togPattern.Current.ToggleState.ToString());
                                        _logicalName = testData.Template.Rows[i]["FieldName"].ToString();
                                        dr[_logicalName] = (string)togPattern.Current.ToggleState.ToString();
                                        break;

                                    case "button":
                                        logTofile(_eLogPtah, "no value available");
                                        _logicalName = testData.Template.Rows[i]["FieldName"].ToString();
                                        dr[_logicalName] = "N/A";
                                        break;

                                    default:
                                        ValuePattern val = (ValuePattern)elementNode.GetCurrentPattern(ValuePattern.Pattern);
                                        logTofile(_eLogPtah, "Value is -->" + val.Current.Value.ToString());
                                        _logicalName = testData.Template.Rows[i]["FieldName"].ToString();
                                        dr[_logicalName] = (string)val.Current.Value;
                                        break;
                                }
                            }
                            else
                            {
                                logTofile(_eLogPtah, "no value available");
                                _logicalName = testData.Template.Rows[i]["FieldName"].ToString();
                                dr[_logicalName] = "N/A";
                            }


                        }
                        else
                        {
                            _logicalName = testData.Template.Rows[i]["FieldName"].ToString();
                            dr[_logicalName] = (string)valpat.Current.Value;
                        }
                    }
                    dtwpf.Rows.Add(dr);
                    dr = dtwpf.NewRow();
                }
                return dtwpf;
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[GetActualDataGrid2Content]:Execption Message : " + ex.Message.ToString());
                throw new Exception(ex.Message.ToString());
                // return null;
            }

        }



        public void VerifyDataGridContent(string testDataFile, string testcase, string section, string customcolumnfile, string resultFilePath)
        {
            try
            {
                Helper.ReportsManagement rppt1 = new Helper.ReportsManagement();

                testData.ExpectedData.Clear();
                testData.ActualData.Clear();
                testData.Template.Clear();
                testData.GetVerificationData(testDataFile, testcase);
                testData.ActualData = GetActualDataGridContent(testDataFile);
                testData.UpdateReporterSheet(customcolumnfile, "testcase", testcase);
                testData.UpdateReporterSheet(customcolumnfile, "section", "NA");
                testData.UpdateReporterSheet(customcolumnfile, "webtable", section);
                testData.CompareDataForm();
                rppt1.ResultTable = testData.ResultTable;
                rppt1.ReportPath = resultFilePath;
                rppt1.GenerateReport(customcolumnfile);
                testData.ResultTable.Clear();
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[VerifyDataGridContent]:Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[VerifyDataGridContent] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// This function is to verify the grid in which cell does not have value but its immediate child has . Also child element's control type is not uniform.
        /// </summary>
        /// <param name="testDataFile"></param>
        /// <param name="testcase"></param>
        /// <param name="section"></param>
        /// <param name="customcolumnfile"></param>
        /// <param name="resultFilePath"></param>
        public void VerifyDataGrid2Content(string testDataFile, string testcase, string section, string customcolumnfile, string resultFilePath)
        {
            try
            {
                Helper.ReportsManagement rppt1 = new Helper.ReportsManagement();

                testData.ExpectedData.Clear();
                testData.ActualData.Clear();
                testData.Template.Clear();
                testData.GetVerificationData(testDataFile, testcase);
                testData.ActualData = GetActualDataGrid2Content(testDataFile);
                testData.UpdateReporterSheet(customcolumnfile, "testcase", testcase);
                testData.UpdateReporterSheet(customcolumnfile, "section", "NA");
                testData.UpdateReporterSheet(customcolumnfile, "webtable", section);
                testData.CompareDataForm();
                rppt1.ResultTable = testData.ResultTable;
                rppt1.ReportPath = resultFilePath;
                rppt1.GenerateReport(customcolumnfile);
                testData.ResultTable.Clear();
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[VerifyDataGrid2Content]:Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[VerifyDataGrid2Content] Line number : " + GetStacktrace(ex).ToString());
            }
        }
        public void verifyDataForm(string testDataFile, string testcase, string section, string customcolumnfile, string resultFilePath)
        {
            try
            {
                Helper.ReportsManagement rppt1 = new Helper.ReportsManagement();

                try
                {
                    testData.ExpectedData.Clear();
                }
                catch (Exception ex)
                {

                    logTofile(_eLogPtah, "Error occured in Expected Data: " + ex.Message);
                }

                logTofile(_eLogPtah, "Cleared Expected Data");
                try
                {
                    testData.ActualData.Clear();
                }
                catch (Exception ex)
                {
                    logTofile(_eLogPtah, "Error occured in Actual Data: " + ex.Message);

                }
                logTofile(_eLogPtah, "Cleared Actual Data");

                try
                {
                    testData.Template.Clear();
                }
                catch (Exception ex)
                {

                    logTofile(_eLogPtah, "Error occured in template  Data: " + ex.Message);
                }


                logTofile(_eLogPtah, "Cleared Template Data");

                testData.GetVerificationDataForm(testDataFile, testcase);
                logTofile(_eLogPtah, "GetVerificationDataForm completed");
                testData.ActualData = GetActualDataForm(testDataFile);
                logTofile(_eLogPtah, "GetActualDataForm completed");
                testData.UpdateReporterSheet(customcolumnfile, "testcase", testcase);
                testData.UpdateReporterSheet(customcolumnfile, "section", section);
                testData.UpdateReporterSheet(customcolumnfile, "webtable", "NA");
                logTofile(_eLogPtah, "Reporter sheet updated");
                testData.CompareDataForm();
                logTofile(_eLogPtah, "comparing data completed");
                rppt1.ResultTable = testData.ResultTable;
                rppt1.ReportPath = resultFilePath;
                rppt1.GenerateReport(customcolumnfile);
                logTofile(_eLogPtah, "Report generated");
                testData.ResultTable.Clear();
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[verifyDataForm]:Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[verifyDataForm] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message);
            }


        }
        public void verifyDataIEPaneTable(string testDataFile, string testcase, string webtable, string customcolumnfile, string resultFilePath)
        {
            try
            {
                Helper.ReportsManagement rppt1 = new Helper.ReportsManagement();

                testData.ExpectedData.Clear();
                testData.ActualData.Clear();
                testData.Template.Clear();
                testData.Data.Clear();

                testData.GetVerificationData(testDataFile, testcase);
                testData.ActualData = GetActualDataIEPaneTable(testDataFile);
                testData.UpdateReporterSheet(customcolumnfile, "testcase", testcase);
                testData.UpdateReporterSheet(customcolumnfile, "section", "NA");
                testData.UpdateReporterSheet(customcolumnfile, "webtable", webtable);
                testData.CompareData();
                rppt1.ResultTable = testData.ResultTable;
                rppt1.ReportPath = resultFilePath;
                rppt1.GenerateReport(customcolumnfile);
                testData.ResultTable.Clear();
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[VerifyDataGrid2Content]:Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[VerifyDataGrid2Content] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message);

            }


        }
        private DataTable GetActualDataInfragisticsTableFlat(string testDataFile)
        {


            string TableType = "";
            string SearchBy = "";
            string SearchValue = "";
            var _logicalName = "";//coulmnname/fieldname  alias 
            DataTable dt = new DataTable();
            try
            {
                if (testData.Template.Rows.Count <= 0)
                {
                    throw new Exception("[GetActualDataInfragisticsTableFlat]:No data available in template sheet");
                }

                DataRow dr = dt.NewRow();
                TableType = testData.Template.Rows[0]["TableType"].ToString();
                SearchBy = testData.Template.Rows[0]["SearchBy"].ToString().ToLower();
                SearchValue = testData.Template.Rows[0]["SearchValue"].ToString();
                string colnamesarray = "";
                string colnameitem = "";
                for (int i = 0; i < testData.Template.Rows.Count; i++)
                {
                    if (Convert.IsDBNull(testData.Template.Rows[i]["FieldName"]) == false)
                    {
                        dt.Columns.Add((string)testData.Template.Rows[i]["FieldName"]);
                        colnameitem = (string)testData.Template.Rows[i]["FieldName"];
                        colnamesarray = colnamesarray + colnameitem + ";";
                    }
                }

                int colcount = testData.Template.Rows.Count;
                AutomationElement infratable = GetUIAutomationInfraTableFlat("", SearchValue, -1);


                if (infratable == null)
                {
                    logTofile(_eLogPtah, "[GetActualDataInfragisticsTableFlat]:table not found");
                    System.Console.WriteLine("[GetActualDataInfragisticsTableFlat]:table not found");
                }
                else
                {
                    logTofile(_eLogPtah, "[GetActualDataInfragisticsTableFlat]: found Infra. table!! now.. finding childern of type  custom ....");
                    System.Console.WriteLine("[GetActualDataInfragisticsTableFlat]: found Infra. table!! now.. finding childern of type  custom ....");
                    AutomationElementCollection tablerows = infratable.FindAll(TreeScope.Children,
                                              new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom));
                    System.Console.WriteLine(tablerows.Count);

                    char[] celldellim = new char[] { ';' };

                    string[] arr = colnamesarray.Split(celldellim);
                    int actrowcount = 0;
                    bool rflag = false;
                    for (int rcnt = 0; rcnt < tablerows.Count; rcnt++)
                    {
                        AutomationElementCollection rowcells = tablerows[rcnt].FindAll(TreeScope.Children,
                                              new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Custom));
                        string actcellval = "";
                        string rowdata = "";

                        for (int colinfcnt = 0; colinfcnt < rowcells.Count; colinfcnt++)
                        {
                            ValuePattern v1 = SupportsValuePattern(rowcells[colinfcnt]);
                            string colname = rowcells[colinfcnt].Current.Name;
                            colname = colname.Replace("\n", "_");
                            if (v1 != null)
                            {
                                ValuePattern vals = (ValuePattern)rowcells[colinfcnt].GetCurrentPattern(ValuePattern.Pattern);
                                actcellval = vals.Current.Value;
                                _logicalName = colname;
                                for (int i = 0; i < arr.Length; i++)
                                {
                                    if (colname == arr[i])
                                    {

                                        rowdata = rowdata + actcellval + "|";
                                        dr[_logicalName] = actcellval;
                                        rflag = true;

                                    }
                                }
                            }

                        }
                        if (rflag)
                        {
                            actrowcount = actrowcount + 1;
                            dt.Rows.Add(dr);
                            dr = dt.NewRow();
                        }
                        logTofile(_eLogPtah, "[GetActualDataInfragisticsTableFlat]: The row data is for " + actrowcount + ":" + rowdata);
                        rflag = false;


                    }

                }



                return dt;
            }

            catch (Exception ex)
            {
                System.Console.WriteLine("[GetActualDataInfragisticsTableFlat]: Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[GetActualDataInfragisticsTableFlat] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message);
            }






        }
        public void verifyInfragisticsTableFlat(string testDataFile, string testcase, string customcolumnfile, string resultFilePath)
        {
            try
            {
                Helper.ReportsManagement rppt1 = new Helper.ReportsManagement();

                testData.ExpectedData.Clear();
                testData.ActualData.Clear();
                testData.Template.Clear();
                testData.Data.Clear();

                testData.GetVerificationData(testDataFile, testcase);
                testData.ActualData = GetActualDataInfragisticsTableFlat(testDataFile);
                testData.UpdateReporterSheet(customcolumnfile, "testcase", testcase);
                testData.CompareDataForm();
                rppt1.ResultTable = testData.ResultTable;
                rppt1.ReportPath = resultFilePath;
                rppt1.GenerateReport(customcolumnfile);
                testData.ResultTable.Clear();
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[verifyInfragisticsTableFlat]: Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[verifyInfragisticsTableFlat] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message);
            }


        }
        #endregion
        public void closeGlobalWindow()
        {
            try
            {
                //_globalWindow.Close();
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[closeGlobalWindow]: Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[closeGlobalWindow] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message);
            }
        }
        private void inputDateInCalendarControl1(AutomationElement calendarbtn, string dateval)
        {
            try
            {
                InvokePattern invk = (InvokePattern)calendarbtn.GetCurrentPattern(InvokePattern.Pattern);
                invk.Invoke();
                char[] celldellim = new char[] { '-' };
                string[] arrcelladd = dateval.Split(celldellim);
                string yearval = arrcelladd[2];
                string monthval = arrcelladd[1];
                string dayval = arrcelladd[0];

                AutomationElement partHeader = GetUIAutomationbutton("automationid", "PART_HeaderButton", -1);
                ClickControl(partHeader);
                Thread.Sleep(1000);

                ClickControl(partHeader);

                int curyear = System.DateTime.Now.Year;
                int searchyear = Int32.Parse(yearval);
                if (curyear < searchyear)
                {
                    AutomationElement partNext = GetUIAutomationbutton("automationid", "PART_NextButton", -1);
                    AutomationElement yearbutton = GetUIAutomationbutton("text", yearval, -1);
                    do
                    {
                        yearbutton = GetUIAutomationbutton("text", yearval, -1);
                        if (yearbutton == null)
                        {
                            ClickControl(partNext);
                            logTofile(_eLogPtah, "year button not found clicking next ");
                        }
                    } while (yearbutton == null);
                    logTofile(_eLogPtah, "Finally year button was found");
                    ClickControl(yearbutton);
                }
                else
                {
                    AutomationElement partBack = GetUIAutomationbutton("automationid", "PART_PreviousButton", -1);
                    AutomationElement yearbutton = GetUIAutomationbutton("text", yearval, -1);
                    do
                    {
                        yearbutton = GetUIAutomationbutton("text", yearval, -1);
                        if (yearbutton == null)
                        {
                            ClickControl(partBack);
                            logTofile(_eLogPtah, "year button not found clicking back ");
                        }
                    } while (yearbutton == null);
                    logTofile(_eLogPtah, "Finally year button was found");
                    ClickControl(yearbutton);
                    logTofile(_eLogPtah, "clcikded button using click control");
                }
                logTofile(_eLogPtah, "passing month as " + monthval);
                Thread.Sleep(1000);
                AutomationElement monthbutton = GetUIAutomationbutton("text", monthval + ", " + yearval, -1);
                ClickControl(monthbutton);
                logTofile(_eLogPtah, "passing day as " + dayval);
                Thread.Sleep(1000);
                string dayval1 = monthval + " " + dayval + ", " + yearval;
                Console.WriteLine("Stringed Date is : " + dayval1);
                DateTime dt = Convert.ToDateTime(dayval1);
                CultureInfo ci = CultureInfo.CurrentCulture;
                Console.WriteLine(ci.DateTimeFormat.LongDatePattern.ToString());
                string ftt = "";
                ftt = ci.DateTimeFormat.LongDatePattern.ToString();
                Console.WriteLine("Curret date in format of machine: final ===" + dt.ToString(ftt));
                string dayval2 = dt.ToString(ftt);
                logTofile(_eLogPtah, "input day format was in name: " + dateval);
                AutomationElement daybutton = GetUIAutomationbutton("text", dayval2, -1);
                ClickControl(daybutton);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[inputDateInCalendarControl1]: Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[inputDateInCalendarControl1] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message);
            }

        }

        private string GetOSArchitecture()
        {
            string osbit = "";
            try
            {
                SelectQuery query = new SelectQuery(@"Select * from Win32_Processor");

                //initialize the searcher with the query it is supposed to execute
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query))
                {
                    //execute the query
                    foreach (ManagementObject process in searcher.Get())
                    {

                        //print process properties
                        // Console.WriteLine("/*********Processor Information ***************/");
                        //  Console.WriteLine("{0}{1}", "Addres Bit 32Bt Or 64 bit :", process["AddressWidth"]);
                        osbit = process["AddressWidth"].ToString();
                    }

                }
                return osbit;
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("[GetOSArchitecture]: Execption Message : " + ex.Message.ToString());
                logTofile(_eLogPtah, "[GetOSArchitecture] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception(ex.Message);
            }
        }
        private static void KeyBoardEnter(string strvalue)
        {

            System.Windows.Forms.SendKeys.SendWait("{Home}");
            Thread.Sleep(500);
            System.Windows.Forms.SendKeys.SendWait("+{End}");
            Thread.Sleep(500);
            System.Windows.Forms.SendKeys.SendWait("{Del}");
            Thread.Sleep(500);
            System.Windows.Forms.SendKeys.SendWait(strvalue);
        }
        #region GenericControlMethods
        private AutomationElement GetControlByLabel(System.Windows.Automation.ControlType controlType, string searchValue)
        {
            logTofile(_eLogPtah, "Searching " + controlType.LocalizedControlType + " by: " + searchValue);

            try
            {
                logTofile(_eLogPtah, "[GetControlByLabel]: Inside Try.");
                AutomationElement control = null;
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");
                    AutomationElement currentRoot = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants
                                         , new System.Windows.Automation.AndCondition(new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Text)
                                         , new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, searchValue)));


                    var foundControl = false;
                    while (foundControl == false)
                    {

                        control = TreeWalker.ControlViewWalker.GetNextSibling(currentRoot);

                        logTofile(_eLogPtah, "Current root is " + control.Current.ControlType.ProgrammaticName);

                        if (control.Current.ControlType == controlType)
                        {
                            foundControl = true;
                            logTofile(_eLogPtah, "Found " + controlType.LocalizedControlType + " searched by " + searchValue);
                        }
                        else
                        {
                            currentRoot = control;
                        }
                    }
                    if (control != null)
                    {
                        break;
                    }
                }
                if (control != null)
                {
                    uilog.AddTexttoColumn("Control Detected", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("Control Detected", "No");
                }


                return control;
            }
            catch (Exception ex)
            {
                uilog.AddTexttoColumn("Control Detected", "No");
                logTofile(_eLogPtah, "Exception in  function GetControlByLabel : " + ex.Message);
                logTofile(_eLogPtah, "[GetControlByLabel] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception("Exception in  function GetControlByLabel : " + ex.Message);
            }
        }
        private AutomationElement GetControlByIndex(System.Windows.Automation.ControlType controlType, int index)
        {


            try
            {
                logTofile(_eLogPtah, "Inside Function GetControlByIndex");
                AutomationElement control = null;
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");

                    AutomationElementCollection controlCol = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, controlType));


                    logTofile(_eLogPtah, "Total " + controlType.LocalizedControlType + " controls found: " + controlCol.Count.ToString());

                    logTofile(_eLogPtah, "Searching  for " + controlType.LocalizedControlType + " with index : " + index.ToString());



                    for (int k = 0; k < controlCol.Count; k++)
                    {
                        logTofile(_eLogPtah, "Bounding Rectangle: (" + k.ToString() + ")" + controlCol[k].Current.BoundingRectangle.ToString());

                        if (index == k)
                        {
                            control = controlCol[k];
                            logTofile(_eLogPtah, " Index found");
                            break;
                        }
                        else
                        {

                            logTofile(_eLogPtah, " Index not found");
                        }

                    }

                    if (control != null)
                    {
                        break;
                    }
                }
                if (control != null)
                {
                    uilog.AddTexttoColumn("Control Detected", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("Control Detected", "No");
                }

                return control;

            }
            catch (Exception ex)
            {

                logTofile(_eLogPtah, "Exception in  function GetControlByIndex : " + ex.Message);
                logTofile(_eLogPtah, "[GetControlByIndex] Line number : " + GetStacktrace(ex).ToString());
                throw new Exception("Exception in  function GetControlByIndex : " + ex.Message);
            }

        }

        private AutomationElement GetControlByName(System.Windows.Automation.ControlType controlType, string SearchValue)
        {
            if (SearchValue == null || SearchValue == "")
            {
                logTofile(_eLogPtah, "Error: No Control Name was specifed in excel sheet Please check excel sheet");
            }
            logTofile(_eLogPtah, "Inside Function GetControlByName:");
            AutomationElement GetControlByName = null;
            logTofile(_eLogPtah, "Searching  for " + controlType.LocalizedControlType + " with Name: " + SearchValue);
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");
                    GetControlByName = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                     new System.Windows.Automation.AndCondition(
                                           new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
                                       new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, SearchValue, PropertyConditionFlags.IgnoreCase)));
                    if (GetControlByName != null)
                    {
                        break;
                    }
                }

                logTofile(_eLogPtah, "Ensuring that Controls was returned:");

                if (GetControlByName != null)
                {
                    logTofile(_eLogPtah, "Name of Control If Control(object) was found :" + GetControlByName.Current.Name);
                    uilog.AddTexttoColumn("Control Detected", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("Control Detected", "No");
                }
                return GetControlByName;

            }

            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Execption from [GetControlByName]" + ex.Message.ToString());
                logTofile(_eLogPtah, "Line number is Source File where Error is encountered is:" + GetStacktrace(ex));
                throw new Exception("Execption from [GetControlByName]" + ex.Message.ToString());

            }

        }
        //private AutomationElement GetControlByNameAndIndex(System.Windows.Automation.ControlType controlType, string SearchValue,int index)
        //{
        //    if (SearchValue == null || SearchValue == "")
        //    {
        //        logTofile(_eLogPtah, "Error: No Control Name was specifed in excel sheet Please check excel sheet");
        //    }
        //    logTofile(_eLogPtah, "Inside Function GetControlByName:");
        //    AutomationElement GetControlByNameAndIndex = null;
        //    logTofile(_eLogPtah, "Searching  for " + controlType.LocalizedControlType + " with Name: " + SearchValue);
        //    try
        //    {
        //        for (var i = 0; i < _Attempts; i++)
        //        {
        //            System.Threading.Thread.Sleep(1);
        //            logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");
        //             AutomationElementCollection allelem = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
        //                             new System.Windows.Automation.AndCondition(
        //                                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
        //                               new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, SearchValue, PropertyConditionFlags.IgnoreCase)));
        //             int cnt = 0;
        //             foreach (AutomationElement elem in allelem)
        //             {
        //                 if (cnt == index)
        //                 {
        //                     GetControlByNameAndIndex = elem;
        //                     break;
        //                 }
        //             }


        //             if (GetControlByNameAndIndex != null)
        //            {
        //                break;
        //            }
        //        }

        //        logTofile(_eLogPtah, "Ensuring that Controls was returned:");

        //        if (GetControlByNameAndIndex != null)
        //        {
        //            logTofile(_eLogPtah, "Name of Control If Control(object) was found :" + GetControlByNameAndIndex.Current.Name);
        //            uilog.AddTexttoColumn("Control Detected", "Yes");
        //        }
        //        else
        //        {
        //            uilog.AddTexttoColumn("Control Detected", "No");
        //        }
        //        return GetControlByNameAndIndex;

        //    }

        //    catch (Exception ex)
        //    {
        //        logTofile(_eLogPtah, "Execption from [GetControlByName]" + ex.Message.ToString());
        //        logTofile(_eLogPtah, "Line number is Source File where Error is encountered is:" + GetStacktrace(ex));
        //        throw new Exception("Execption from [GetControlByName]" + ex.Message.ToString());

        //    }

        //}

        private AutomationElement GetControlByAutomationId(System.Windows.Automation.ControlType controlType, string SearchValue)
        {
            AutomationElement GetControlByAutomationId = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");


                    GetControlByAutomationId = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                      new System.Windows.Automation.AndCondition(
                                            new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
                                        new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, SearchValue)));
                    if (GetControlByAutomationId != null)
                    {
                        break;
                    }
                }
                if (GetControlByAutomationId != null)
                {
                    uilog.AddTexttoColumn("Control Detected", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("Control Detected", "No");
                }

                return GetControlByAutomationId;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetControlByAutomationId] Error in Fucntion " + ex.Message.ToString());
                return GetControlByAutomationId;
                throw new Exception("Exception in  function GetControlByAutomationId : " + ex.Message);
            }

        }
        private AutomationElement GetControlByNameFromCollectionAndIndex(System.Windows.Automation.ControlType controlType, string SearchValue, int index)
        {
            AutomationElement GetControlByNameFromCollectionAndIndex = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");

                    AutomationElementCollection aecollection = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                            new System.Windows.Automation.AndCondition(
                                            new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
                                            new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, SearchValue)));
                    logTofile(_eLogPtah, "Count of chidlitems: " + aecollection.Count);
                    if (index == -1)
                    {
                        GetControlByNameFromCollectionAndIndex = aecollection[0];
                    }
                    else
                    {
                        GetControlByNameFromCollectionAndIndex = aecollection[index];
                    }

                    if (GetControlByNameFromCollectionAndIndex != null)
                    {
                        break;
                    }
                }
                if (GetControlByNameFromCollectionAndIndex != null)
                {
                    uilog.AddTexttoColumn("Control Detected", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("Control Detected", "No");
                }
                return GetControlByNameFromCollectionAndIndex;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetControlByNameFromCollectionAndIndex] Error in Fucntion " + ex.Message.ToString());
                return GetControlByNameFromCollectionAndIndex;
                throw new Exception("Exception in  function GetControlByNameFromCollectionAndIndex : " + ex.Message);
            }
        }

        private AutomationElement GetControlByAutomationIdFromCollectionAndIndex(System.Windows.Automation.ControlType controlType, string SearchValue, int index)
        {
            AutomationElement GetControlByNameFromCollectionAndIndex = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");
                    AutomationElementCollection aecollection = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                            new System.Windows.Automation.AndCondition(
                                            new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
                                            new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, SearchValue)));
                    GetControlByNameFromCollectionAndIndex = aecollection[index];
                    if (GetControlByNameFromCollectionAndIndex != null)
                    {
                        break;
                    }
                }
                if (GetControlByNameFromCollectionAndIndex != null)
                {
                    uilog.AddTexttoColumn("Control Detected", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("Control Detected", "No");
                }

                return GetControlByNameFromCollectionAndIndex;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetControlByAutomationIdFromCollectionAndIndex] Error in Fucntion " + ex.Message.ToString());
                return GetControlByNameFromCollectionAndIndex;
                throw new Exception("Exception in  function GetControlByAutomationIdFromCollectionAndIndex : " + ex.Message);
            }
        }
        private AutomationElement GetControlByHelpTextFromCollectionAndIndex(System.Windows.Automation.ControlType controlType, string SearchValue, int index)
        {
            AutomationElement GetControlByHelpTextFromCollectionAndIndex = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");
                    AutomationElementCollection aecollection = uiAutomationCurrentParent.FindAll(TreeScope.Descendants,
                                            new System.Windows.Automation.AndCondition(
                                            new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
                                            new System.Windows.Automation.PropertyCondition(AutomationElement.HelpTextProperty, SearchValue)));
                    GetControlByHelpTextFromCollectionAndIndex = aecollection[index];
                    if (GetControlByHelpTextFromCollectionAndIndex != null)
                    {
                        break;
                    }
                }
                if (GetControlByHelpTextFromCollectionAndIndex != null)
                {
                    uilog.AddTexttoColumn("Control Detected", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("Control Detected", "No");
                }

                return GetControlByHelpTextFromCollectionAndIndex;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "[GetControlByAutomationIdFromCollectionAndIndex] Error in Fucntion " + ex.Message.ToString());
                return GetControlByHelpTextFromCollectionAndIndex;
                throw new Exception("Exception in  function GetControlByAutomationIdFromCollectionAndIndex : " + ex.Message);
            }
        }
        private AutomationElement GetControlByHelpText(System.Windows.Automation.ControlType controlType, string SearchValue)
        {
            AutomationElement GetControlByHelpText = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");
                    GetControlByHelpText = uiAutomationCurrentParent.FindFirst(TreeScope.Descendants,
                                     new System.Windows.Automation.AndCondition(
                                           new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
                                       new System.Windows.Automation.PropertyCondition(AutomationElement.HelpTextProperty, SearchValue)));

                    if (GetControlByHelpText != null)
                    {
                        break;
                    }
                }
                if (GetControlByHelpText != null)
                {
                    uilog.AddTexttoColumn("Control Detected", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("Control Detected", "No");
                }

                return GetControlByHelpText;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Error in Finding Control by helptext" + ex.Message.ToString());
                return GetControlByHelpText;
                throw new Exception("Error in GetControlByHelpText refer line number of code above");

            }
        }

        private AutomationElement GetControlByClassNameandIndex(string classname, int index)
        {
            AutomationElement GetControlByClassNameandIndex = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");
                    AutomationElementCollection allelems =
                     uiAutomationCurrentParent.FindAll(TreeScope.Descendants,

                                           new System.Windows.Automation.PropertyCondition(AutomationElement.ClassNameProperty, classname));
                    int cnnt = 0;
                    foreach (AutomationElement elem in allelems)
                    {
                        if (cnnt == index)
                        {
                            GetControlByClassNameandIndex = elem;
                            break;
                        }
                        cnnt++;

                    }
                    if (GetControlByClassNameandIndex != null)
                    {
                        break;
                    }

                }
                if (GetControlByClassNameandIndex != null)
                {
                    uilog.AddTexttoColumn("Control Detected", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("Control Detected", "No");
                }

                return GetControlByClassNameandIndex;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Error in Finding Control by ClassName and index" + ex.Message.ToString());
                return GetControlByClassNameandIndex;
                throw new Exception("Error in GetControlByHelpText refer line number of code above");

            }
        }

        private AutomationElement GetWindowByName(string SearchValue)
        {
            AutomationElement GetWindowByName = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");

                    GetWindowByName = uiAutomationapp.FindFirst(TreeScope.Descendants,
                          new System.Windows.Automation.AndCondition(
                                       new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                        new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, SearchValue.Trim()), new System.Windows.Automation.PropertyCondition(AutomationElement.ProcessIdProperty, _processId)));


                    if (GetWindowByName != null)
                    {
                        logTofile(_eLogPtah, "Window was found only by Strict Name :" + GetWindowByName.Current.Name);
                        break;
                    }
                    else // Also check Fi we can get Window without using Process ID;
                    {
                        GetWindowByName = uiAutomationapp.FindFirst(TreeScope.Descendants,
                      new System.Windows.Automation.AndCondition(
                                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, SearchValue.Trim())));


                        if (GetWindowByName != null)
                        {

                            break;
                        }
                    }
                }
                if (GetWindowByName != null)
                {
                    uilog.AddTexttoColumn("WindowConstructed", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("WindowConstructed", "No");
                }
                return GetWindowByName;
            }
            catch (Exception ex)
            {
                logTofile(_eLogPtah, "Could not find Window only by Strict Name" + ex.Message.ToString());
                uilog.AddTexttoColumn("WindowConstructed", "No");
                return GetWindowByName;
                throw new Exception("Error in GetWindowByName refer line number of code above");
            }

        }
        private AutomationElement GetWindowByProcessId()
        {
            AutomationElement GetWindowByName = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");

                    GetWindowByName = uiAutomationapp.FindFirst(TreeScope.Descendants,
                          new System.Windows.Automation.AndCondition(
                                       new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                        new System.Windows.Automation.PropertyCondition(AutomationElement.ProcessIdProperty, _processId)));

                    if (GetWindowByName != null)
                    {
                        logTofile(_eLogPtah, "Window was found only by Strict Name :" + GetWindowByName.Current.Name);
                        break;
                    }
                }
                if (GetWindowByName != null)
                {
                    uilog.AddTexttoColumn("WindowConstructed", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("WindowConstructed", "No");
                }
                return GetWindowByName;
            }
            catch (Exception ex)
            {
                uilog.AddTexttoColumn("WindowConstructed", "No");
                logTofile(_eLogPtah, "Error in Finding Window only by Strict Name" + ex.Message.ToString());
                return GetWindowByName;
                throw new Exception("Error in GetWindowByName refer line number of code above");
            }

        }

        private AutomationElement GetWindowByPartialName(string SearchValue)
        {

            AutomationElement GetWindowByPartialName = null;
            try
            {
                for (var icnt = 0; icnt < _Attempts; icnt++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + icnt.ToString() + " times");

                    AutomationElementCollection windows = uiAutomationapp.FindAll(TreeScope.Children,
                   new System.Windows.Automation.AndCondition(
                                       new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                        new System.Windows.Automation.PropertyCondition(AutomationElement.ProcessIdProperty, _processId)));

                    for (int i = 0; i < windows.Count; i++)
                    {
                        if (windows[i].Current.Name.ToLower().Trim().Contains(SearchValue.ToLower().Trim()))
                        {
                            GetWindowByPartialName = windows[i];
                            break;
                        }
                    }
                    if (GetWindowByPartialName != null)
                    {
                        break;
                    }
                }

                if (GetWindowByPartialName != null)
                {
                    uilog.AddTexttoColumn("WindowConstructed", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("WindowConstructed", "No");
                }

                return GetWindowByPartialName;
            }
            catch (Exception ex)
            {
                uilog.AddTexttoColumn("WindowConstructed", "No");
                logTofile(_eLogPtah, "Error in Finding window usng partial text also !!!" + ex.Message.ToString());
                logTofile(_eLogPtah, "[GetWindowByPartialName]: Line number is Source File where Error is encountered is:" + GetStacktrace(ex));
                return GetWindowByPartialName;
                throw new Exception("Error in GetWindowByPartialName refer line number of code above.");
            }

        }
        private AutomationElement GetWindowByNameAndIndex(string SearchValue, int index)
        {
            AutomationElement GetWindowByNameAndIndex = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");

                    AutomationElementCollection windows = uiAutomationapp.FindAll(TreeScope.Descendants,
                          new System.Windows.Automation.AndCondition(
                                       new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                        new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, SearchValue.Trim()), new System.Windows.Automation.PropertyCondition(AutomationElement.ProcessIdProperty, _processId)));
                    GetWindowByNameAndIndex = windows[index];


                    if (GetWindowByNameAndIndex != null)
                    {
                        logTofile(_eLogPtah, "Window was found only by Strict Name :" + GetWindowByNameAndIndex.Current.Name);
                        break;
                    }
                }
                if (GetWindowByNameAndIndex != null)
                {
                    uilog.AddTexttoColumn("WindowConstructed", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("WindowConstructed", "No");
                }
                return GetWindowByNameAndIndex;
            }
            catch (Exception ex)
            {
                uilog.AddTexttoColumn("WindowConstructed", "No");
                logTofile(_eLogPtah, "Error in Finding Window only by Strict Name" + ex.Message.ToString());
                logTofile(_eLogPtah, "[GetWindowByNameAndIndex]: Line number is Source File where Error is encountered is:" + GetStacktrace(ex));
                return GetWindowByNameAndIndex;
                throw new Exception("Error in Finding Window only by Strict Name");
            }

        }

        private AutomationElement GetWindowByAutomationId(string SearchValue)
        {
            AutomationElement GetWindowByAutomationId = null;
            try
            {
                for (var i = 0; i < _Attempts; i++)
                {
                    System.Threading.Thread.Sleep(1);
                    logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");
                    GetWindowByAutomationId = uiAutomationapp.FindFirst(TreeScope.Descendants,
                          new System.Windows.Automation.AndCondition(
                                       new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                        new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, SearchValue.Trim()), new System.Windows.Automation.PropertyCondition(AutomationElement.ProcessIdProperty, _processId)));
                    //   logTofile(_eLogPtah, "Window was found only by Strict Automation ID :" + GetWindowByAutomationId.Current.Name);




                    if (GetWindowByAutomationId != null)
                    {

                        break;
                    }
                    else // Also check Fi we can get Window without using Process ID;
                    {
                        GetWindowByAutomationId = uiAutomationapp.FindFirst(TreeScope.Descendants,
                      new System.Windows.Automation.AndCondition(
                                   new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Window),
                                    new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, SearchValue.Trim())));


                        if (GetWindowByAutomationId != null)
                        {

                            break;
                        }
                    }


                }
                if (GetWindowByAutomationId != null)
                {
                    uilog.AddTexttoColumn("WindowConstructed", "Yes");
                }
                else
                {
                    uilog.AddTexttoColumn("WindowConstructed", "No");
                }
                return GetWindowByAutomationId;
            }
            catch (Exception ex)
            {

                uilog.AddTexttoColumn("WindowConstructed", "No");
                logTofile(_eLogPtah, "Error in Finding Window only by Strict Automation ID" + ex.Message.ToString());
                throw new Exception("Error in Finding Window only by Strict Name");
            }

        }

        private AutomationElement GetWindowByPartialAutomationId(string SearchValue)
        {
            return null;
        }

        private AutomationElement GetWindowByAutomationIdAndIndex(string SearchValue)
        {
            return null;
        }
        #endregion
        public bool IsColumnPresent(string colname)
        {
            try
            {
                bool IsColumnPresent = false;
                string colNameString = "";
                for (int ic = 0; ic < testData.Template.Columns.Count; ic++)
                {
                    colNameString = colNameString + testData.Template.Columns[ic].Caption.ToString() + ";";
                }
                logTofile(_eLogPtah, "[IsColumnPresent] : The Column names string is : " + colNameString);
                if (colNameString.Contains(colname))
                {
                    IsColumnPresent = true;
                }

                return IsColumnPresent;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

        private string GetStacktrace(Exception ex)
        {

            var st = new StackTrace(ex, true);
            var testFrame = st.GetFrame(0);
            for (int i = 0; i < st.FrameCount; i++)
            {
                testFrame = st.GetFrame(i);
                if (testFrame.GetFileName() != null)
                {
                    if (testFrame.GetFileName().ToString().Contains("UIAutomation.cs") == true)
                    {
                        break;
                    }
                }

            }

            var frame = testFrame;
            // Get the line number from the stack frame
            var line = testFrame.GetFileLineNumber();
            return line.ToString();
        }


        private void WaitForAttempts(AutomationElement elem)
        {
            for (var i = 0; i < _Attempts; i++)
            {
                System.Threading.Thread.Sleep(1);
                logTofile(_eLogPtah, "Searched for " + i.ToString() + " times");
                if (elem != null)
                {
                    break;
                }
            }
        }

        public string dllversion()
        {
            return "V_2.0.0.10";
        }


        public void getdatafromExcelToClipboard(string filepath, string ws, string range)
        {
            Excel.Application objxls = new Excel.Application();
            objxls.DisplayAlerts = false;
            Excel.Workbook wb = objxls.Workbooks.Open(filepath);
            Excel.Sheets wkshts = wb.Worksheets;
            Excel.Worksheet wkshti = null;

            foreach (Excel.Worksheet indsht in wkshts)
            {
                if (indsht.Name.ToLower() == ws.ToLower())
                {
                    wkshti = indsht;
                    break;
                }
            }
            Excel.Range rng = (Excel.Range)wkshti.get_Range(range);
            rng.Select();
            objxls.Selection.Copy();
            System.Windows.Forms.Clipboard.GetData(System.Windows.Forms.DataFormats.Text);
            wb.Close();
            objxls.Quit();
        }

    }    //Class 


    public class UIAutomationLog
    {
        public DataTable _logTable = new DataTable();
        public DataRow dr;


        public string uiAutoamtionreportPath
        {
            get;
            set;
        }
        public string fileName
        {
            get;
            set;
        }

        public void AddHeaders()
        {

            if (IsColumnPresent("FunctionName") == false)
            {
                _logTable.Columns.Add("FunctionName");
            }
            if (IsColumnPresent("StructureSheetName") == false)
            {
                _logTable.Columns.Add("StructureSheetName");
            }
            if (IsColumnPresent("TestCaseID") == false)
            {
                _logTable.Columns.Add("TestCaseID");
            }
            if (IsColumnPresent("ParentType") == false)
            {
                _logTable.Columns.Add("ParentType");
            }
            if (IsColumnPresent("ParentSearchBy") == false)
            {
                _logTable.Columns.Add("ParentSearchBy");
            }
            if (IsColumnPresent("ParentSearchValue") == false)
            {
                _logTable.Columns.Add("ParentSearchValue");
            }
            if (IsColumnPresent("ControlType") == false)
            {

                _logTable.Columns.Add("ControlType");
            }
            if (IsColumnPresent("FieldName") == false)
            {
                _logTable.Columns.Add("FieldName");
            }
            if (IsColumnPresent("Action") == false)
            {
                _logTable.Columns.Add("Action");
            }

            if (IsColumnPresent("Index") == false)
            {
                _logTable.Columns.Add("Index");
            }
            if (IsColumnPresent("SearchBy") == false)
            {
                _logTable.Columns.Add("SearchBy");
            }
            if (IsColumnPresent("ControlName") == false)
            {
                _logTable.Columns.Add("ControlName");
            }
            if (IsColumnPresent("ControlValue") == false)
            {
                _logTable.Columns.Add("ControlValue");
            }
            if (IsColumnPresent("WindowConstructed") == false)
            {
                _logTable.Columns.Add("WindowConstructed");
            }
            if (IsColumnPresent("Control Detected") == false)
            {
                _logTable.Columns.Add("Control Detected");
            }
            if (IsColumnPresent("Action Performed on Control") == false)
            {
                _logTable.Columns.Add("Action Performed on Control");
            }
            if (IsColumnPresent("Value Entered in Control") == false)
            {
                _logTable.Columns.Add("Value Entered in Control");
            }

        }

        public void createnewrow()
        {
            dr = _logTable.NewRow();
        }
        public void commitrow()
        {
            _logTable.Rows.Add(dr);
        }
        public void AddTexttoColumn(string columnName, string strtext)
        {
            if (_logTable.Columns.Count > 0)
            {
                if (strtext.Length <= 0)
                {
                    strtext = "_";
                }
                dr[columnName] = strtext;
            }
        }

        public void CreateCSVfile(string uiAutoamtionreportPath, string fileName)
        {
            string ReportPath = uiAutoamtionreportPath + fileName + "Log.csv";
            // StreamWriter writer = new StreamWriter(ReportPath, true);
            //  if (System.IO.File.Exists(ReportPath))
            //  {

            //  }
            using (StreamWriter writer = new StreamWriter(ReportPath, true))
            {
                if (writer.BaseStream.Length == 0)
                {
                    foreach (DataColumn column in _logTable.Columns)
                    {


                        writer.Write('\u0022' + column.ColumnName + '\u0022' + ",");

                    }
                    writer.WriteLine();
                }
                for (int i = 0; i < _logTable.Rows.Count; i++)
                {


                    foreach (DataColumn column in _logTable.Columns)
                    {
                        if (_logTable.Rows[i][column.ColumnName] != DBNull.Value && _logTable.Rows[i][column.ColumnName].ToString().Length != 0)
                        {
                            writer.Write('\u0022' + (string)_logTable.Rows[i][column.ColumnName] + '\u0022' + ",");
                        }
                        else
                        {
                            writer.Write('\u0022' + "_" + '\u0022' + ",");
                        }
                    }
                    writer.WriteLine();
                }
            }


        }

        private bool IsColumnPresent(string colname)
        {
            try
            {
                bool IsColumnPresent = false;
                string colNameString = "";
                for (int ic = 0; ic < _logTable.Columns.Count; ic++)
                {
                    colNameString = colNameString + _logTable.Columns[ic].Caption.ToString() + ";";
                }
                if (colNameString.Contains(colname))
                {
                    IsColumnPresent = true;
                }

                return IsColumnPresent;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public void ClearDataTable()
        {
            _logTable.Clear();
        }
    }


}
