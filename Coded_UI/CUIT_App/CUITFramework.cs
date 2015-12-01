using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Globalization;
using System.Drawing;
using System.Reflection;
using System.IO;
using System.Xml;
using Microsoft.Win32;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
using Microsoft.VisualStudio.TestTools.UITest.Framework;
using Microsoft.VisualStudio.TestTools.UITest.Playback;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
using Helper;
using System.Diagnostics;
using System.Drawing.Imaging;




namespace CUIT_app
{
    /// <summary>
    /// Summary description for TestWellFlo
    /// </summary>

    public class clsCUIT_app
    {
        private string _error = "Error in Function ";
        TestDataManagement testData = new TestDataManagement();
        ReportsManagement report = new ReportsManagement();
        public WinWindow _globalWindow { get; set; }
        public String _eLogPtah { get; set; }
        public String _toleranceImage = @"D:\PO\PAN\Screenshots\difference.png";
        //this is used in generic exception inside catch block
        public void Drag(string searchBy, string SearchValue)
        {
            Console.WriteLine("Inside function Drag:");
            int x;
            int y;
            int newx;
            int newy;
            string[] oldCoordinates = searchBy.Split(',');
            string[] newcoordinates = SearchValue.Split(',');
            x = int.Parse(oldCoordinates[0]);
            y = int.Parse(oldCoordinates[1]);
            newx = int.Parse(newcoordinates[0]);
            newy = int.Parse(newcoordinates[1]);
            try
            {
                Playback.Initialize();
                Point oldPoint = new Point(x, y);
                Point newPoint = new Point(newx, newy);
                Mouse.Hover(oldPoint);
                System.Threading.Thread.Sleep(1000);
                Mouse.StartDragging();
                Mouse.StopDragging(newPoint);
                Playback.Cleanup();
            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in Drag and in line: " + line);
                throw new Exception(_error + System.Environment.NewLine + ex.Message);

            }


        }
        public void Drag(UITestControl startElement, UITestControl Endelement)
        {
            Console.WriteLine("Inside function Drag:");

            try
            {
                Playback.Initialize();
                Mouse.Click(startElement);
                Mouse.StartDragging();
                Mouse.StopDragging(Endelement);
                Playback.Cleanup();
            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in Drag and in line: " + line);
                throw new Exception(_error + System.Environment.NewLine + ex.Message);

            }


        }

        public WinWindow GetCUITWindow(string searchBy, string SearchValue)
        {
            Console.WriteLine("Inside function getCUITWindow:");
            try
            {
                Playback.Initialize();
                WinWindow GetCUITWindow = new WinWindow();

                switch (searchBy.ToLower())
                {
                    case "title":
                    case "name":
                    case "text":
                        {
                            Console.WriteLine("Searching window by title");
                            GetCUITWindow.SearchProperties[WinWindow.PropertyNames.Name] = SearchValue;
                            break;
                        }
                    case "parttext":
                        {
                            Console.WriteLine("Searching window by title");
                            GetCUITWindow.SearchProperties.Add(WinWindow.PropertyNames.Name, SearchValue, PropertyExpressionOperator.Contains);
                            break;
                        }
                    case "automationid":
                        {
                            GetCUITWindow.SearchProperties.Add(WinWindow.PropertyNames.ControlName, SearchValue);
                            UITestControlCollection checkboxCollection = GetCUITWindow.FindMatchingControls();
                            GetCUITWindow = (WinWindow)checkboxCollection[0];
                            break;
                        }
                    //case "controlid":
                    //    {
                    //        GetCUITWindow.SearchProperties.Add(WinWindow.PropertyNames.ControlId, SearchValue);
                    //        UITestControlCollection checkboxCollection = GetCUITWindow.FindMatchingControls();
                    //        GetCUITWindow = (WinWindow)checkboxCollection[0];
                    //        break;
                    //    }
                    default:
                        {
                            throw new Exception(_error + "CUITWindow:" + "only name and autoid are valid for Widnow");
                        }
                }

                Playback.Cleanup();
                Console.WriteLine("Found Window and exiting function getCUITWindow:");
                return GetCUITWindow;

            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITWindow and in line: " + line);
                throw new Exception(_error + "CUITWindow:" + System.Environment.NewLine + ex.Message);

            }


        }
        //this function is called inside CUITWindow() and its overloaded method

        public WinButton GetCUITButton(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITButton:");
            WinButton GetCUITButton = new WinButton(w);

            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {

                                GetCUITButton.SearchProperties[WinButton.PropertyNames.Name] = searchValue;
                            }
                            else
                            {

                                GetCUITButton.SearchProperties.Add(WinButton.PropertyNames.Name, searchValue);
                                UITestControlCollection buttonCollection = GetCUITButton.FindMatchingControls();
                                GetCUITButton = (WinButton)buttonCollection[index];

                            }
                            break;
                        }
                    case "regulartext":
                        {
                            if (index == -1)
                            {

                                GetCUITButton.SearchProperties.Add(WinButton.PropertyNames.Name, searchValue, PropertyExpressionOperator.Contains);
                                UITestControlCollection buttonCollection = GetCUITButton.FindMatchingControls();
                                GetCUITButton = (WinButton)buttonCollection[0];
                            }
                            else
                            {

                                GetCUITButton.SearchProperties.Add(WinButton.PropertyNames.Name, searchValue);
                                UITestControlCollection buttonCollection = GetCUITButton.FindMatchingControls();
                                GetCUITButton = (WinButton)buttonCollection[index];

                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITButton.SearchProperties.Add(WinButton.PropertyNames.ControlName, searchValue);
                                UITestControlCollection buttonCollection = GetCUITButton.FindMatchingControls();
                                GetCUITButton = (WinButton)buttonCollection[0];
                            }
                            else
                            {

                                GetCUITButton.SearchProperties.Add(WinButton.PropertyNames.ControlName, searchValue);
                                UITestControlCollection buttonCollection = GetCUITButton.FindMatchingControls();
                                GetCUITButton = (WinButton)buttonCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);

                }
                Playback.Cleanup();
                Console.WriteLine("Found Button and exiting function GetCUITButton");
                return GetCUITButton;

            }
            catch (Exception e)
            {
                // Get stack trace for the exception with source file information
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITButton and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message + " in line " + line);
            }
        }
        public WinRadioButton GetCUITRadioButton(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITRadioButton");
            WinRadioButton GetCUITRadioButton = new WinRadioButton(w);
            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITRadioButton.SearchProperties[WinRadioButton.PropertyNames.Name] = searchValue;

                            }
                            else
                            {
                                GetCUITRadioButton.SearchProperties.Add(WinRadioButton.PropertyNames.Name, searchValue);
                                UITestControlCollection radioButtonCollection = GetCUITRadioButton.FindMatchingControls();
                                GetCUITRadioButton = (WinRadioButton)radioButtonCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITRadioButton.SearchProperties.Add(WinRadioButton.PropertyNames.ControlName, searchValue);
                                UITestControlCollection radioButtonCollection = GetCUITRadioButton.FindMatchingControls();
                                GetCUITRadioButton = (WinRadioButton)radioButtonCollection[0];
                            }
                            else
                            {
                                GetCUITRadioButton.SearchProperties.Add(WinRadioButton.PropertyNames.ControlName, searchValue);
                                UITestControlCollection radioButtonCollection = GetCUITRadioButton.FindMatchingControls();
                                GetCUITRadioButton = (WinRadioButton)radioButtonCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }
                Playback.Cleanup();
                Console.WriteLine("Found Radio Button and exiting function GetCUITRadioButton");
                return GetCUITRadioButton;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITRadioButton and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinComboBox GetCUITComboBox(WinWindow w, string searchBy, string searchValue, int index)
        {
            WinComboBox GetCUITComboBox = new WinComboBox(w);
            Console.WriteLine("Inside function GetCUITComboBox");
            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITComboBox.SearchProperties[WinComboBox.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITComboBox.SearchProperties.Add(WinComboBox.PropertyNames.Name, searchValue);
                                UITestControlCollection comboboxCollection = GetCUITComboBox.FindMatchingControls();
                                GetCUITComboBox = (WinComboBox)comboboxCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITComboBox.SearchProperties.Add(WinComboBox.PropertyNames.ControlName, searchValue);
                                UITestControlCollection comboboxCollection = GetCUITComboBox.FindMatchingControls();
                                GetCUITComboBox = (WinComboBox)comboboxCollection[0];
                            }
                            else
                            {
                                GetCUITComboBox.SearchProperties.Add(WinComboBox.PropertyNames.ControlName, searchValue);
                                UITestControlCollection comboboxCollection = GetCUITComboBox.FindMatchingControls();
                                GetCUITComboBox = (WinComboBox)comboboxCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }
                Playback.Cleanup();
                Console.WriteLine("Found Combobox and exiting function GetCUITComboBox");
                return GetCUITComboBox;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITComboBox and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinEdit GetCUITEdit(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITEdit");
            Playback.Initialize();
            UITestControl cntl = new UITestControl(w);
            cntl.TechnologyName = "MSAA";
            cntl.SearchProperties.Add("ControlType", "Edit");
           // WinEdit GetCUITEdit = new WinEdit(w);
            WinEdit GetCUITEdit = null;
            
            try
            {

                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                               // GetCUITEdit.SearchProperties[WinEdit.PropertyNames.Name] = searchValue;
                                
                                cntl.SearchProperties.Add("Name", searchValue);
                                UITestControlCollection editCollection = cntl.FindMatchingControls();
                                GetCUITEdit = (WinEdit)editCollection[0];
                            }
                            else
                            {
                                //GetCUITEdit.SearchProperties.Add(WinEdit.PropertyNames.Name, searchValue);
                                //UITestControlCollection editCollection = GetCUITEdit.FindMatchingControls();
                                //GetCUITEdit = (WinEdit)editCollection[index];
                                cntl.SearchProperties.Add("ControlType", "Edit");
                                cntl.SearchProperties.Add("Name", searchValue);
                                UITestControlCollection editCollection = cntl.FindMatchingControls();
                                GetCUITEdit = (WinEdit)editCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                //GetCUITEdit.SearchProperties.Add(WinEdit.PropertyNames.ControlName, searchValue);
                                //UITestControlCollection editCollection = GetCUITEdit.FindMatchingControls();
                                //GetCUITEdit = (WinEdit)editCollection[0];
                               
                                UITestControlCollection editCollection = cntl.FindMatchingControls();
                               int i = 0;
                                foreach (UITestControl edt in editCollection)
                                {
                                    if (((WinEdit)edt).ControlName == searchValue)
                                    {
                                        GetCUITEdit = (WinEdit)edt;
                                        break;
                                    }
                                    i++;
                                }
            
                                
                            }
                            else
                            {
                                //GetCUITEdit.SearchProperties.Add(WinEdit.PropertyNames.ControlName, searchValue);
                                //UITestControlCollection editCollection = GetCUITEdit.FindMatchingControls();
                                //GetCUITEdit = (WinEdit)editCollection[index];
                               
                                UITestControlCollection editCollection = cntl.FindMatchingControls();
                                int i = 0;
                                int j = 0;
                                foreach (UITestControl edt in editCollection)
                                {
                                    if ( ((WinEdit)edt).ControlName == searchValue && j==index)
                                    {
                                        GetCUITEdit = (WinEdit)edt;
                                        j++;
                                        break;
                                    }
                                    i++;
                                }

                                
                            }
                            break;
                        }
                    case "notext":
                        {
                            if (index == -1)
                            {
                                UITestControlCollection editCollection = cntl.FindMatchingControls();
                                GetCUITEdit = (WinEdit)editCollection[0];
                            }
                            else
                            {
                                UITestControlCollection editCollection = cntl.FindMatchingControls();
                                GetCUITEdit = (WinEdit)editCollection[index];
                            }
                            
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }
                Playback.Cleanup();
                Console.WriteLine("Found Text box and exiting function GetCUITEdit");
                return GetCUITEdit;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITEdit and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinRow GetCUITDataRow(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITDataRow");
            WinRow GetCUITDataRow = new WinRow(w);
            Playback.Initialize();
            try
            {

                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITDataRow.SearchProperties[WinRow.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITDataRow.SearchProperties.Add(WinRow.PropertyNames.Name, searchValue);
                                UITestControlCollection editCollection = GetCUITDataRow.FindMatchingControls();
                                GetCUITDataRow = (WinRow)editCollection[index];

                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITDataRow.SearchProperties.Add(WinRow.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataRow.FindMatchingControls();
                                GetCUITDataRow = (WinRow)editCollection[0];
                            }
                            else
                            {
                                GetCUITDataRow.SearchProperties.Add(WinRow.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataRow.FindMatchingControls();
                                GetCUITDataRow = (WinRow)editCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }
                Playback.Cleanup();
                Console.WriteLine("Found Data Row and exiting function GetCUITDataRow");
                return GetCUITDataRow;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITEdit and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinRowHeader GetCUITDataRowHeader(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITDataRow");
            WinRowHeader GetCUITDataRowHeader = new WinRowHeader(w);
            Playback.Initialize();
            try
            {

                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITDataRowHeader.SearchProperties[WinRowHeader.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITDataRowHeader.SearchProperties.Add(WinRowHeader.PropertyNames.Name, searchValue);
                                UITestControlCollection editCollection = GetCUITDataRowHeader.FindMatchingControls();
                                GetCUITDataRowHeader = (WinRowHeader)editCollection[index];

                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITDataRowHeader.SearchProperties.Add(WinRowHeader.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataRowHeader.FindMatchingControls();
                                GetCUITDataRowHeader = (WinRowHeader)editCollection[0];
                            }
                            else
                            {
                                GetCUITDataRowHeader.SearchProperties.Add(WinRowHeader.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataRowHeader.FindMatchingControls();
                                GetCUITDataRowHeader = (WinRowHeader)editCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }
                Playback.Cleanup();
                Console.WriteLine("Found Data Row and exiting function GetCUITDataRow");
                return GetCUITDataRowHeader;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITEdit and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinColumnHeader GetCUITDataColumnHeader(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITDataRow");
            WinColumnHeader GetCUITDataColumnHeader = new WinColumnHeader(w);
            Playback.Initialize();
            try
            {

                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITDataColumnHeader.SearchProperties[WinColumnHeader.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITDataColumnHeader.SearchProperties.Add(WinColumnHeader.PropertyNames.Name, searchValue);
                                UITestControlCollection editCollection = GetCUITDataColumnHeader.FindMatchingControls();
                                GetCUITDataColumnHeader = (WinColumnHeader)editCollection[index];

                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITDataColumnHeader.SearchProperties.Add(WinColumnHeader.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataColumnHeader.FindMatchingControls();
                                GetCUITDataColumnHeader = (WinColumnHeader)editCollection[0];
                            }
                            else
                            {
                                GetCUITDataColumnHeader.SearchProperties.Add(WinColumnHeader.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataColumnHeader.FindMatchingControls();
                                GetCUITDataColumnHeader = (WinColumnHeader)editCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }
                Playback.Cleanup();
                Console.WriteLine("Found Data Row and exiting function GetCUITDataRow");
                return GetCUITDataColumnHeader;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITEdit and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinTable GetCUITDataTable(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITDataRow");
            WinTable GetCUITDataTable = new WinTable(w);
            Playback.Initialize();
            try
            {

                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITDataTable.SearchProperties[WinTable.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITDataTable.SearchProperties.Add(WinTable.PropertyNames.Name, searchValue);
                                UITestControlCollection editCollection = GetCUITDataTable.FindMatchingControls();
                                GetCUITDataTable = (WinTable)editCollection[index];

                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITDataTable.SearchProperties.Add(WinTable.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataTable.FindMatchingControls();
                                GetCUITDataTable = (WinTable)editCollection[0];
                            }
                            else
                            {
                                GetCUITDataTable.SearchProperties.Add(WinTable.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataTable.FindMatchingControls();
                                GetCUITDataTable = (WinTable)editCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }
                Playback.Cleanup();
                Console.WriteLine("Found Data Row and exiting function GetCUITDataRow");
                return GetCUITDataTable;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITEdit and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinCell GetCUITDataCell(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITDataCell");
            WinCell GetCUITDataCell = new WinCell(w);
            Playback.Initialize();
            try
            {

                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITDataCell.SearchProperties[WinCell.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITDataCell.SearchProperties.Add(WinCell.PropertyNames.Name, searchValue);
                                UITestControlCollection editCollection = GetCUITDataCell.FindMatchingControls();
                                GetCUITDataCell = (WinCell)editCollection[index];

                            }
                            break;
                        }
                    case "regulartext":
                        {
                            if (index == -1)
                            {
                                GetCUITDataCell.SearchProperties.Add(WinCell.PropertyNames.Name, searchValue, PropertyExpressionOperator.Contains);
                                UITestControlCollection editCollection = GetCUITDataCell.FindMatchingControls();
                                GetCUITDataCell = (WinCell)editCollection[0];
                            }
                            else
                            {
                                GetCUITDataCell.SearchProperties.Add(WinCell.PropertyNames.Name, searchValue, PropertyExpressionOperator.Contains);
                                UITestControlCollection editCollection = GetCUITDataCell.FindMatchingControls();
                                GetCUITDataCell = (WinCell)editCollection[index];

                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITDataCell.SearchProperties.Add(WinCell.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataCell.FindMatchingControls();
                                GetCUITDataCell = (WinCell)editCollection[0];
                            }
                            else
                            {
                                GetCUITDataCell.SearchProperties.Add(WinCell.PropertyNames.ControlName, searchValue);
                                UITestControlCollection editCollection = GetCUITDataCell.FindMatchingControls();
                                GetCUITDataCell = (WinCell)editCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }
                Playback.Cleanup();
                Console.WriteLine("Found Data Cell and exiting function GetCUITDataCell");
                return GetCUITDataCell;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITEdit and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }

        public WinTabPage GetCUITTabpage(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITTabpage");
            WinTabPage GetCUITTabpage = new WinTabPage(w);
            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITTabpage.SearchProperties[WinTabPage.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITTabpage.SearchProperties.Add(WinTabPage.PropertyNames.Name, searchValue);
                                UITestControlCollection tabPageCollection = GetCUITTabpage.FindMatchingControls();
                                GetCUITTabpage = (WinTabPage)tabPageCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITTabpage.SearchProperties.Add(WinTabPage.PropertyNames.ControlName, searchValue);
                                UITestControlCollection tabPageCollection = GetCUITTabpage.FindMatchingControls();
                                GetCUITTabpage = (WinTabPage)tabPageCollection[0];
                            }
                            else
                            {
                                GetCUITTabpage.SearchProperties.Add(WinTabPage.PropertyNames.ControlName, searchValue);
                                UITestControlCollection tabPageCollection = GetCUITTabpage.FindMatchingControls();
                                GetCUITTabpage = (WinTabPage)tabPageCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }

                Playback.Cleanup();
                Console.WriteLine("Found Tab and exiting function GetCUITTabpage");
                return GetCUITTabpage;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITTabpage and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinCheckBox GetCUITCHeckbox(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITCHeckbox");
            WinCheckBox GetCUITCHeckbox = new WinCheckBox(w);
            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITCHeckbox.SearchProperties[WinCheckBox.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITCHeckbox.SearchProperties.Add(WinCheckBox.PropertyNames.Name, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITCHeckbox.FindMatchingControls();
                                GetCUITCHeckbox = (WinCheckBox)checkboxCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITCHeckbox.SearchProperties.Add(WinCheckBox.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITCHeckbox.FindMatchingControls();
                                GetCUITCHeckbox = (WinCheckBox)checkboxCollection[0];
                            }
                            else
                            {
                                GetCUITCHeckbox.SearchProperties.Add(WinCheckBox.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITCHeckbox.FindMatchingControls();
                                GetCUITCHeckbox = (WinCheckBox)checkboxCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }

                Playback.Cleanup();
                Console.WriteLine("Found Checkbox and exiting function GetCUITCHeckbox");
                return GetCUITCHeckbox;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITCHeckbox and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinListItem GetCUITListItem(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITListItem");
            WinListItem GetCUITListItem = new WinListItem(w);

            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITListItem.SearchProperties[WinListItem.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITListItem.SearchProperties.Add(WinListItem.PropertyNames.Name, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITListItem.FindMatchingControls();
                                GetCUITListItem = (WinListItem)checkboxCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITListItem.SearchProperties.Add(WinListItem.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITListItem.FindMatchingControls();
                                GetCUITListItem = (WinListItem)checkboxCollection[0];
                            }
                            else
                            {
                                GetCUITListItem.SearchProperties.Add(WinListItem.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITListItem.FindMatchingControls();
                                GetCUITListItem = (WinListItem)checkboxCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }

                Playback.Cleanup();
                Console.WriteLine("Found List Item and exiting function GetCUITListItem");
                return GetCUITListItem;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITCHeckbox and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinMenuItem GetCUITMenuItem(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITMenuItem");
            WinMenuItem GetCUITMenuItem = new WinMenuItem(w);
            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITMenuItem.SearchProperties[WinMenuItem.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITMenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.Name, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITMenuItem.FindMatchingControls();
                                GetCUITMenuItem = (WinMenuItem)checkboxCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITMenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITMenuItem.FindMatchingControls();
                                GetCUITMenuItem = (WinMenuItem)checkboxCollection[0];
                            }
                            else
                            {
                                GetCUITMenuItem.SearchProperties.Add(WinMenuItem.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITMenuItem.FindMatchingControls();
                                GetCUITMenuItem = (WinMenuItem)checkboxCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }

                Playback.Cleanup();
                Console.WriteLine("Found MenuItem and exiting function GetCUITMenuItem");
                return GetCUITMenuItem;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITCHeckbox and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinTreeItem GetCUITTreeItem(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITMenuItem");
            WinTreeItem GetCUITTreeItem = new WinTreeItem(w);
            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITTreeItem.SearchProperties.Add(WinTreeItem.PropertyNames.Name, searchValue, PropertyExpressionOperator.Contains);
                                UITestControlCollection checkboxCollection = GetCUITTreeItem.FindMatchingControls();
                                GetCUITTreeItem = (WinTreeItem)checkboxCollection[0];
                            }
                            else
                            {
                                GetCUITTreeItem.SearchProperties.Add(WinTreeItem.PropertyNames.Name, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITTreeItem.FindMatchingControls();
                                GetCUITTreeItem = (WinTreeItem)checkboxCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITTreeItem.SearchProperties.Add(WinTreeItem.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITTreeItem.FindMatchingControls();
                                GetCUITTreeItem = (WinTreeItem)checkboxCollection[0];
                            }
                            else
                            {
                                GetCUITTreeItem.SearchProperties.Add(WinTreeItem.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITTreeItem.FindMatchingControls();
                                GetCUITTreeItem = (WinTreeItem)checkboxCollection[index];
                            }
                            break;
                        }
                    case "regulartext":
                        {
                            if (index == -1)
                            {

                                GetCUITTreeItem.SearchProperties.Add(WinTreeItem.PropertyNames.Name, searchValue, PropertyExpressionOperator.Contains);
                                UITestControlCollection checkboxCollection = GetCUITTreeItem.FindMatchingControls();
                                GetCUITTreeItem = (WinTreeItem)checkboxCollection[0];
                            }
                            else
                            {

                                GetCUITTreeItem.SearchProperties.Add(WinTreeItem.PropertyNames.Name, searchValue, PropertyExpressionOperator.Contains);
                                UITestControlCollection buttonCollection = GetCUITTreeItem.FindMatchingControls();
                                GetCUITTreeItem = (WinTreeItem)buttonCollection[index];

                            }
                            break;
                        }


                    default:
                        throw new Exception(_error);
                }

                Playback.Cleanup();
                Console.WriteLine("Found MenuItem and exiting function GetCUITMenuItem");
                return GetCUITTreeItem;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITCHeckbox and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinList GetCUITList(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITList");
            WinList GetCUITList = new WinList(w);

            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITList.SearchProperties[WinList.PropertyNames.Name] = searchValue;
                            }
                            else
                            {
                                GetCUITList.SearchProperties.Add(WinList.PropertyNames.Name, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITList.FindMatchingControls();
                                GetCUITList = (WinList)checkboxCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITList.SearchProperties.Add(WinList.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITList.FindMatchingControls();
                                GetCUITList = (WinList)checkboxCollection[0];
                            }
                            else
                            {
                                GetCUITList.SearchProperties.Add(WinList.PropertyNames.ControlName, searchValue);
                                UITestControlCollection checkboxCollection = GetCUITList.FindMatchingControls();
                                GetCUITList = (WinList)checkboxCollection[index];
                            }
                            break;
                        }
                    case "notext":
                        {
                            UITestControlCollection checkboxCollection = w.GetChildren();
                            for (int i = 0; i < checkboxCollection.Count; i++)
                            {
                                UITestControl control = checkboxCollection[i];
                                if (control.ControlType.ToString().ToLower() == "winlist")
                                {
                                    GetCUITList = (WinList)control;
                                }
                            }
                                
                                
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }

                Playback.Cleanup();
                Console.WriteLine("Found List and exiting function GetCUITList");
                return GetCUITList;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITCHeckbox and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }
        public WinText GetCUITTextcontrol(WinWindow w, string searchBy, string searchValue, int index)
        {
            Console.WriteLine("Inside function GetCUITTextcontrol");
            WinText GetCUITTextcontrol = new WinText(w);

            try
            {
                Playback.Initialize();
                switch (searchBy.Trim().ToLower())
                {
                    case "text":
                        {
                            if (index == -1)
                            {
                                GetCUITTextcontrol.SearchProperties.Add(WinText.PropertyNames.Name, searchValue, PropertyExpressionOperator.Contains);
                                UITestControlCollection textCollection = GetCUITTextcontrol.FindMatchingControls();
                                GetCUITTextcontrol = (WinText)textCollection[0];
                            }
                            else
                            {
                                GetCUITTextcontrol.SearchProperties.Add(WinText.PropertyNames.Name, searchValue);
                                UITestControlCollection textCollection = GetCUITTextcontrol.FindMatchingControls();
                                GetCUITTextcontrol = (WinText)textCollection[index];
                            }
                            break;
                        }

                    case "automationid":
                        {
                            if (index == -1)
                            {
                                GetCUITTextcontrol.SearchProperties.Add(WinText.PropertyNames.ControlName, searchValue);
                                UITestControlCollection textCollection = GetCUITTextcontrol.FindMatchingControls();
                                GetCUITTextcontrol = (WinText)textCollection[0];
                            }
                            else
                            {
                                GetCUITTextcontrol.SearchProperties.Add(WinText.PropertyNames.ControlName, searchValue);
                                UITestControlCollection textCollection = GetCUITTextcontrol.FindMatchingControls();
                                GetCUITTextcontrol = (WinText)textCollection[index];
                            }
                            break;
                        }

                    default:
                        throw new Exception(_error);
                }

                Playback.Cleanup();
                Console.WriteLine("Found TextControl and exiting function GetCUITTextcontrol");
                return GetCUITTextcontrol;

            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in GetCUITCHeckbox and in line: " + line);
                throw new Exception(_error + "CUITRButton:" + System.Environment.NewLine + e.Message);
            }
        }

        private void AddData(int rowPosition)
        {
            string parentType = "";
            string parentSearchBy = "";
            string parentSearchValue = "";
            var _controlType = "";
            var _logicalName = "";
            try
            {
                for (int i = 0; i < testData.Structure.Rows.Count; i++)
                {
                    parentType = testData.Structure.Rows[i]["ParentType"].ToString();
                    parentSearchBy = testData.Structure.Rows[i]["ParentSearchBy"].ToString().ToLower();
                    parentSearchValue = testData.Structure.Rows[i]["ParentSearchValue"].ToString();

                    if ((string)testData.Structure.Rows[i]["inputdata"].ToString().ToLower() == "y")
                    {
                        if (Convert.IsDBNull(testData.Structure.Rows[i]["ParentType"]) == false)
                        {
                            switch (parentType.Trim().ToLower())
                            {
                                case "cuitwindow":
                                    {
                                        _globalWindow = GetCUITWindow(parentSearchBy, parentSearchValue);
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
                                default:
                                    throw new Exception("Valid action types are keyboard, wait, pagedown, pageup");
                            }
                        }
                        if (Convert.IsDBNull(testData.Structure.Rows[i]["ControlType"]) == false)
                        {
                            _controlType = (string)testData.Structure.Rows[i]["ControlType"].ToString().ToLower();
                            _logicalName = (string)testData.Structure.Rows[i]["FieldName"].ToString();
                            _searchBy = (string)testData.Structure.Rows[i]["SearchBy"];


                            Console.WriteLine(_logicalName);

                            if (Convert.IsDBNull(testData.Structure.Rows[i]["Index"]) == false)
                            {
                                _index = int.Parse(testData.Structure.Rows[i]["Index"].ToString());
                            }
                            Console.WriteLine("index " + _index);
                            string _controlValue = null;
                            if (_logicalName.Length > 0)
                            {
                                _controlValue = (string)testData.Data.Rows[rowPosition][_logicalName].ToString();
                            }
                            Console.WriteLine("controlValue" + _controlValue);
                            if (_logicalName.Length > 0 && _controlValue.Length == 0)
                            {

                            }
                            else
                            {
                                switch (_controlType.Trim().ToLower())
                                {
                                    case "cuitbutton":
                                        Console.WriteLine("trying to get cuit button");
                                        WinButton button = GetCUITButton(_globalWindow, _searchBy, _controlName, _index);
                                        Console.WriteLine("got cuit button");
                                        Playback.Initialize();
                                        Mouse.Click(button);
                                        Playback.Cleanup();
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
                var st = new StackTrace(ex, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in AddData and in line: " + line);
                throw new Exception(_error + "AddData:" + System.Environment.NewLine + ex.Message);
            }
        }
        //this function is called inside AddData(string testDataPath, string testCase)

        public void AddData(string testDataPath, string testCase)
        {
            try
            {
                testData.GetTestData(testDataPath, testCase);

                AddData(0);
            }

            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                var testFrame = st.GetFrame(0);
                for (int i = 0; i < st.FrameCount; i++)
                {
                    testFrame = st.GetFrame(i);
                    if (testFrame.GetFileName() != null)
                    {
                        if (testFrame.GetFileName().ToString().Contains("CUITFramework.cs") == true)
                        {
                            break;
                        }
                    }

                }
                // Get the top stack frame
                var frame = testFrame;
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                logTofile(_eLogPtah, "Error in AddData and in line: " + line);
                throw new Exception(ex.Message);
            }
        }
        //this function adds data from an excel sheet

        public void logTofile(string spath, string stxtMsg)
        {
            System.IO.File.AppendAllText(spath, System.DateTime.Now + ":" + stxtMsg + Environment.NewLine);

            try
            {
                Console.WriteLine(stxtMsg);
            }
            catch
            {

            }

        }
        

    }
}
