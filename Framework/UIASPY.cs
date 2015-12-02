using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace CUTE
{
    public partial class Form1 : Form
    {
        string _txtBaseFolder = "";
        string _testAppPath = "";
        string _testAppTitle = "";
        string _currentParentName = "";
        DataTable children = new DataTable();
        int _rowIndex;
        string excelPath = "";
        DataGridViewComboBoxColumn actionColumn = new DataGridViewComboBoxColumn();


        public Form1()
        {
            InitializeComponent();
            tabMain.TabPages.Remove(tbNewProject);
            tabMain.TabPages.Remove(tbObjectSpy);
            tabMain.TabPages.Remove(tbOpenProject);
            tabMain.TabPages.Remove(tbSettings);
            children.Columns.Add("PARENTTYPE");
            children.Columns.Add("PARENT");
            children.Columns.Add("PARENTID");
            children.Columns.Add("CONTROLTYPE");
            children.Columns.Add("NAME");
            children.Columns.Add("AUTOMATIONID");
            children.Columns.Add("DATA");
            actionColumn.Name = "ACTION";
            List<string> actionList = new List<string>();
            actionList.Add("Select Action");
            actionList.Add("Wait");
            actionList.Add("KeyBoard");
            actionList.Add("ClearWindow");
            actionColumn.DataSource = actionList;
            dgvObjects.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        }

        private void tsbtnNew_Click(object sender, EventArgs e)
        {
            tabMain.TabPages.Add(tbNewProject);
            tabMain.TabPages.Remove(tbObjectSpy);
            tabMain.TabPages.Remove(tbOpenProject);
            tabMain.TabPages.Remove(tbSettings);
        }

        private void tsbtnOpen_Click(object sender, EventArgs e)
        {
            tabMain.TabPages.Remove(tbNewProject);
            tabMain.TabPages.Remove(tbObjectSpy);
            tabMain.TabPages.Add(tbOpenProject);
            tabMain.TabPages.Remove(tbSettings);
        }

        private void tsbtnConfigure_Click(object sender, EventArgs e)
        {
            tabMain.TabPages.Remove(tbNewProject);
            tabMain.TabPages.Remove(tbObjectSpy);
            tabMain.TabPages.Remove(tbOpenProject);
            tabMain.TabPages.Add(tbSettings);
        }

        private void tsbtnSpy_Click(object sender, EventArgs e)
        {
            tabMain.TabPages.Remove(tbNewProject);
            tabMain.TabPages.Add(tbObjectSpy);
            tabMain.TabPages.Remove(tbOpenProject);
            tabMain.TabPages.Remove(tbSettings);
            MessageBox.Show("Please Place cursor on the control and press shift key to get its identification properties");


        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (this.Text == "NewProject")
            {
                MessageBox.Show("Please save the existing project before creating another project");
            }
            else
            {
                FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
                DialogResult result = folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtBasefolder.Text = folderBrowserDialog1.SelectedPath;
                    _txtBaseFolder = txtBasefolder.Text;

                }
            }
        }
       protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
       //  protected override bool    ProcessKeyPreview(ref Message msg, Keys keyData)
        {
            this.Focus();
            
          if (keyData == (Keys.ShiftKey | Keys.Shift))
           //    if (keyData == (Keys.Left))
            //    if (keyData == (Keys.Alt | Keys.F8))

            {
                MessageBox.Show("OK Control Got Focuseed");
                System.Windows.Point point = new System.Windows.Point(MousePosition.X, MousePosition.Y);
                AutomationElement element = AutomationElement.FromPoint(point);
                System.Windows.Rect boundingRect1 = (System.Windows.Rect)element.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty); ;
                object boundingRectNoDefault = element.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty, true);
                if (boundingRectNoDefault == AutomationElement.NotSupported)
                {
                    // TODO Handle the case where you do not wish to proceed using the default value.
                }
                else
                {
                    boundingRect1 = (System.Windows.Rect)boundingRectNoDefault;

                }
                //System.Windows.Point points = boundingRect1.TopLeft;
                //System.Drawing.Point topleft = new System.Drawing.Point(Convert.ToInt32(points.X), Convert.ToInt32(points.Y));
                //System.Windows.Size size = boundingRect1.Size;
                //System.Drawing.Size windowSize = new System.Drawing.Size(Convert.ToInt32(size.Width), Convert.ToInt32(size.Height));
                //System.Drawing.Rectangle rect = new System.Drawing.Rectangle(topleft, windowSize);
                //ControlPaint.DrawReversibleFrame(rect, SystemColors.Highlight, FrameStyle.Thick);
                TreeWalker walker = TreeWalker.ControlViewWalker;
                string autoIdString;
                object autoIdNoDefault = element.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty, true);
                object nameDefault = element.GetCurrentPropertyValue(AutomationElement.NameProperty, true);
                object controlTypeDefault = element.GetCurrentPropertyValue(AutomationElement.LocalizedControlTypeProperty, true);
                AutomationElement parent = walker.GetParent(element);
                string parentName = parent.GetCurrentPropertyValue(AutomationElement.NameProperty, true) as string;
                AutomationElement currentParent = null;
                do
                {

                    if (parent == AutomationElement.RootElement)
                    {
                        break;
                    }
                    else
                    {
                        currentParent = parent;
                        parent = walker.GetParent(parent);
                    }
                }
                while (true);
                string currentParentName = currentParent.GetCurrentPropertyValue(AutomationElement.NameProperty, true) as string;
                if (_currentParentName == "")
                {
                    DataRow dr = children.NewRow();
                    _currentParentName = currentParentName;
                    string currentParentType = currentParent.GetCurrentPropertyValue(AutomationElement.LocalizedControlTypeProperty, true) as string;
                    string currentParentID = currentParent.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty, true) as string;
                    dr["PARENTTYPE"] = currentParentType;
                    dr["PARENT"] = currentParentName;
                    if (currentParentID != null)
                    {
                        dr["PARENTID"] = currentParentID;
                    }
                    if (autoIdNoDefault == AutomationElement.NotSupported)
                    {
                        txtControlId.Text = "Not Supported";
                        dr["AUTOMATIONID"] = "";
                        // TODO Handle the case where you do not wish to proceed using the default value.
                    }
                    else
                    {
                        autoIdString = autoIdNoDefault as string;
                        txtControlId.Text = autoIdString;
                        Size sizeId = TextRenderer.MeasureText(txtControlId.Text, txtControlId.Font);
                        txtControlId.Width = sizeId.Width;
                        dr["AUTOMATIONID"] = autoIdString;

                    }
                    if (nameDefault == AutomationElement.NotSupported)
                    {
                        txtControlName.Text = "Not Supported";
                        dr["NAME"] = "";
                    }
                    else
                    {
                        autoIdString = nameDefault as string;
                        txtControlName.Text = autoIdString;
                        Size sizeName = TextRenderer.MeasureText(txtControlName.Text, txtControlName.Font);
                        txtControlName.Width = sizeName.Width;
                        dr["NAME"] = autoIdString;

                    }
                    if (controlTypeDefault == AutomationElement.NotSupported)
                    {
                        txtControlType.Text = "Not Supported";
                        dr["CONTROLTYPE"] = "";

                    }
                    else
                    {
                        autoIdString = controlTypeDefault as string;
                        txtControlType.Text = autoIdString;
                        Size sizeType = TextRenderer.MeasureText(txtControlType.Text, txtControlType.Font);
                        txtControlType.Width = sizeType.Width;
                        dr["CONTROLTYPE"] = autoIdString;

                    }
                    children.Rows.Add(dr);
                }
                else
                {
                    if (_currentParentName == currentParentName)
                    {
                        DataRow dr = children.NewRow();
                        string currentParentType = currentParent.GetCurrentPropertyValue(AutomationElement.LocalizedControlTypeProperty, true) as string;
                        string currentParentID = currentParent.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty, true) as string;
                        dr["PARENTTYPE"] = currentParentType;
                        dr["PARENT"] = currentParentName;
                        if (currentParentID != null)
                        {
                            dr["PARENTID"] = currentParentID;
                        }
                        if (autoIdNoDefault == AutomationElement.NotSupported)
                        {
                            txtControlId.Text = "Not Supported";
                            dr["AUTOMATIONID"] = "";
                            // TODO Handle the case where you do not wish to proceed using the default value.
                        }
                        else
                        {
                            autoIdString = autoIdNoDefault as string;
                            txtControlId.Text = autoIdString;
                            Size sizeId = TextRenderer.MeasureText(txtControlId.Text, txtControlId.Font);
                            txtControlId.Width = sizeId.Width;
                            dr["AUTOMATIONID"] = autoIdString;

                        }
                        if (nameDefault == AutomationElement.NotSupported)
                        {
                            txtControlName.Text = "Not Supported";
                            dr["NAME"] = "";
                        }
                        else
                        {
                            autoIdString = nameDefault as string;
                            txtControlName.Text = autoIdString;
                            Size sizeName = TextRenderer.MeasureText(txtControlName.Text, txtControlName.Font);
                            txtControlName.Width = sizeName.Width;
                            dr["NAME"] = autoIdString;

                        }
                        if (controlTypeDefault == AutomationElement.NotSupported)
                        {
                            txtControlType.Text = "Not Supported";
                            dr["CONTROLTYPE"] = "";

                        }
                        else
                        {
                            autoIdString = controlTypeDefault as string;
                            txtControlType.Text = autoIdString;
                            Size sizeType = TextRenderer.MeasureText(txtControlType.Text, txtControlType.Font);
                            txtControlType.Width = sizeType.Width;
                            dr["CONTROLTYPE"] = autoIdString;

                        }
                        children.Rows.Add(dr);
                    }
                    else
                    {
                        if (MessageBox.Show("Selected element is not in the current parent scope. Do you want to add the element anyway", "Confirm?", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            _currentParentName = currentParentName;
                            DataRow dr = children.NewRow();
                            string currentParentType = currentParent.GetCurrentPropertyValue(AutomationElement.LocalizedControlTypeProperty, true) as string;
                            string currentParentID = currentParent.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty, true) as string;
                            dr["PARENTTYPE"] = currentParentType;
                            dr["PARENT"] = currentParentName;
                            if (currentParentID != null)
                            {
                                dr["PARENTID"] = currentParentID;
                            }
                            if (autoIdNoDefault == AutomationElement.NotSupported)
                            {
                                txtControlId.Text = "Not Supported";
                                dr["AUTOMATIONID"] = "";
                                // TODO Handle the case where you do not wish to proceed using the default value.
                            }
                            else
                            {
                                autoIdString = autoIdNoDefault as string;
                                txtControlId.Text = autoIdString;
                                Size sizeId = TextRenderer.MeasureText(txtControlId.Text, txtControlId.Font);
                                txtControlId.Width = sizeId.Width;
                                dr["AUTOMATIONID"] = autoIdString;

                            }
                            if (nameDefault == AutomationElement.NotSupported)
                            {
                                txtControlName.Text = "Not Supported";
                                dr["NAME"] = "";
                            }
                            else
                            {
                                autoIdString = nameDefault as string;
                                txtControlName.Text = autoIdString;
                                Size sizeName = TextRenderer.MeasureText(txtControlName.Text, txtControlName.Font);
                                txtControlName.Width = sizeName.Width;
                                dr["NAME"] = autoIdString;

                            }
                            if (controlTypeDefault == AutomationElement.NotSupported)
                            {
                                txtControlType.Text = "Not Supported";
                                dr["CONTROLTYPE"] = "";

                            }
                            else
                            {
                                autoIdString = controlTypeDefault as string;
                                txtControlType.Text = autoIdString;
                                Size sizeType = TextRenderer.MeasureText(txtControlType.Text, txtControlType.Font);
                                txtControlType.Width = sizeType.Width;
                                dr["CONTROLTYPE"] = autoIdString;

                            }
                            children.Rows.Add(dr);
                        }

                    }
                }


                dgvObjects.DataSource = children;
                if (dgvObjects.ColumnCount == 7)
                {
                    dgvObjects.Columns.Insert(6, actionColumn);
                }
                dgvObjects.AllowUserToOrderColumns = true;
                dgvObjects.Width = dgvObjects.Columns.Cast<DataGridViewColumn>().Sum(x => x.Width) + (dgvObjects.RowHeadersVisible ? dgvObjects.RowHeadersWidth : 0) + 50;
                dgvObjects.Height = dgvObjects.Rows.Cast<DataGridViewRow>().Sum(x => x.Height) + (dgvObjects.RowHeadersVisible ? dgvObjects.RowHeadersWidth : 0) + 7;
                dgvObjects.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);

        }
        private void txtBasefolder_TextChanged(object sender, EventArgs e)
        {
            ConfigurationManager.AppSettings.Set("baseFolder", txtBasefolder.Text);
            XmlDocument xml = new XmlDocument();
            XmlElement configuration = xml.CreateElement("configuration");
            xml.AppendChild(configuration);
            XmlAttribute xmlns = xml.CreateAttribute("xmlns");
            xmlns.InnerText = "automatedtests";
            configuration.Attributes.SetNamedItem(xmlns);
            XmlElement testBasePath = xml.CreateElement("testbasefolder");
            testBasePath.InnerText = txtBasefolder.Text;
            configuration.AppendChild(testBasePath);
            XmlElement testDataPath = xml.CreateElement("data");
            configuration.InsertAfter(testDataPath, testBasePath);
            XmlElement testReportsPath = xml.CreateElement("testreports");
            configuration.InsertAfter(testReportsPath, testDataPath);
            XmlElement testlogsPath = xml.CreateElement("testlogs");
            configuration.InsertAfter(testlogsPath, testReportsPath);
            string testData = System.IO.Path.Combine(txtBasefolder.Text, "Data");
            string testReports = System.IO.Path.Combine(txtBasefolder.Text, "Reports");
            string testLogs = System.IO.Path.Combine(txtBasefolder.Text, "Logs");
            if (!Directory.Exists(testData))
            {
                Directory.CreateDirectory(testData);

            }
            if (!Directory.Exists(testReports))
            {
                Directory.CreateDirectory(testReports);

            }
            if (!Directory.Exists(testLogs))
            {
                Directory.CreateDirectory(testLogs);

            }
            testDataPath.InnerText = testData;
            testReportsPath.InnerText = testReports;
            testlogsPath.InnerText = testLogs;
            xml.Save(txtBasefolder.Text + @"\NewProject.xml");
            string path = System.Windows.Forms.Application.CommonAppDataPath;
            File.Copy(path + @"\ProjectDb.sdf", txtBasefolder.Text + @"\Data\ProjectDb.sdf", true);
            this.Text = "NewProject";
            MessageBox.Show("New Project is created succesfully. Please configure new project");
            tabMain.TabPages.Add(tbSettings);
            tabMain.TabPages.Remove(tbNewProject);
        }

        private void btnTestApp_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Applications (*.exe)|*.exe";
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtTestApp.Text = dialog.FileName;
                _testAppPath = txtTestApp.Text;
                XmlDocument xml = new XmlDocument();
                string projectName = this.Text;
                xml.Load(_txtBaseFolder + @"\" + projectName + ".xml");
                XmlNode testApp = xml.CreateElement("TestApplication");
                testApp.InnerText = txtTestApp.Text;
                xml.DocumentElement.AppendChild(testApp);
                xml.Save(_txtBaseFolder + @"\" + projectName + ".xml");
            }

        }

        private void tsbtnSave_Click(object sender, EventArgs e)
        {
            string baseFolder = ConfigurationManager.AppSettings["baseFolder"];
            string projectName = null;
            if (File.Exists(baseFolder + @"\NewProject.xml") == true)
            {
                projectName = Interaction.InputBox("Enter project name", "Save", "");
                File.Move(baseFolder + @"\NewProject.xml", baseFolder + @"\" + projectName + ".xml");
            }

            this.Text = projectName;
        }

        private void txtTestAppTitle_TextChanged(object sender, EventArgs e)
        {
            XmlDocument xml = new XmlDocument();
            string projectName = this.Text;
            xml.Load(_txtBaseFolder + @"\" + projectName + ".xml");
            XmlNode testApp = xml.CreateElement("TestApplication");
            testApp.InnerText = txtTestAppTitle.Text;
            xml.DocumentElement.AppendChild(testApp);
            xml.Save(_txtBaseFolder + @"\" + projectName + ".xml");
        }

        private void dgvObjects_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            ContextMenu m = new ContextMenu();
            m.MenuItems.Add(new MenuItem("Insert"));
            m.MenuItems.Add(new MenuItem("Delete"));

            m.MenuItems.Add(new MenuItem("Paste"));

        }

        private void dgvObjects_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu mnu = new ContextMenu();
                MenuItem mnuInsert = new MenuItem("Insert", mnuInsert_Click);
                MenuItem mnuDelete = new MenuItem("Delete", mnuDelete_Click);
                mnu.MenuItems.AddRange(new MenuItem[] { mnuInsert, mnuDelete });
                int currentMouseOverRow = dgvObjects.HitTest(e.X, e.Y).RowIndex;
                _rowIndex = currentMouseOverRow;
                mnu.Show(dgvObjects, new Point(e.X, e.Y));

            }
        }
        private void mnuInsert_Click(object sender, EventArgs e)
        {
            DataRow toInsert = children.NewRow();
            children.Rows.InsertAt(toInsert, _rowIndex);
        }
        private void mnuDelete_Click(object sender, EventArgs e)
        {
            children.Rows.RemoveAt(_rowIndex);

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = new DataTable();
                dt = (DataTable)dgvObjects.DataSource;
                DataTable exportDb = new DataTable();
                exportDb.Columns.Add("InputData");
                exportDb.Columns.Add("Section");
                exportDb.Columns.Add("ParentType");
                exportDb.Columns.Add("ParentSearchBy");
                exportDb.Columns.Add("ParentSearchValue");
                exportDb.Columns.Add("ControlType");
                exportDb.Columns.Add("SearchBy");
                exportDb.Columns.Add("ControlName");
                exportDb.Columns.Add("FieldName");
                exportDb.Columns.Add("Action");
                exportDb.Columns.Add("Index");
                exportDb.Columns.Add("Data");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = exportDb.NewRow();
                    dr["InputData"] = "Y";
                    try
                    {
                        if (dt.Rows[i]["PARENTTYPE"].ToString() != "")
                        {
                            dr["ParentType"] = "uiautiomation" + dt.Rows[i]["PARENTTYPE"].ToString();
                        }
                        else
                        {
                            dr["ParentType"] = "";
                        }
                        if (dt.Rows[i]["PARENT"].ToString() != "")
                        {
                            dr["ParentSearchBy"] = "Title";
                            dr["ParentSearchValue"] = dt.Rows[i]["PARENT"].ToString();
                        }
                        else if (dt.Rows[i]["PARENTID"].ToString() != "")
                        {
                            dr["ParentSearchBy"] = "AutomationId";
                            dr["ParentSearchValue"] = dt.Rows[i]["PARENTID"].ToString();
                        }
                        if (dt.Rows[i]["CONTROLTYPE"].ToString() != "")
                        {
                            dr["ControlType"] = "uiautiomation" + dt.Rows[i]["CONTROLTYPE"].ToString();
                        }
                        else
                        {
                            dr["ControlType"] = "";
                        }
                        if (dt.Rows[i]["NAME"].ToString() != "")
                        {
                            dr["SearchBy"] = "Name";
                            dr["ControlName"] = dt.Rows[i]["NAME"].ToString();
                            dr["FieldName"] = dt.Rows[i]["NAME"].ToString();
                        }
                        else if (dt.Rows[i]["AUTOMATIONID"].ToString() != "")
                        {
                            dr["SearchBy"] = "AutomationId";
                            dr["ControlName"] = dt.Rows[i]["AUTOMATIONID"].ToString();
                            dr["FieldName"] = dt.Rows[i]["AUTOMATIONID"].ToString();
                        }
                        else
                        {
                            dr["SearchBy"] = "";
                            dr["ControlName"] = "";
                            dr["FieldName"] = "";
                        }
                        try
                        {
                            if (dgvObjects.Rows[i].Cells["ACTION"].Value.ToString() != null)
                            {
                                dr["Action"] = dgvObjects.Rows[i].Cells["ACTION"].Value.ToString();
                            }
                        }
                        catch (Exception ex)
                        {
                            dr["Action"] = "";
                        }
                        if (dt.Rows[i]["DATA"] != null)
                        {
                            int num;
                            bool isNum = int.TryParse(dt.Rows[i]["DATA"].ToString(), out num);
                            if (isNum)
                            {
                                dr["Data"] = "'" + dt.Rows[i]["DATA"].ToString();
                            }
                            else
                            {
                                dr["Data"] = dt.Rows[i]["DATA"].ToString();
                            }
                        }
                        exportDb.Rows.Add(dr);

                    }
                    catch (Exception ex)
                    {

                    }
                }
                string path = "";
                if (excelPath == "")
                {
                    path = Interaction.InputBox("Enter ExcelPath", "Create", "");
                    excelPath = path;

                    createExcel(exportDb, path);
                }
                else
                {
                    if (MessageBox.Show("Do you want to append lastly created excel sheet", "Confirm?", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        path = excelPath;
                        appendExcel(exportDb, path);
                    }
                    else
                    {
                        path = Interaction.InputBox("Enter ExcelPath", "Create", "");
                        excelPath = path;
                        createExcel(exportDb, path);
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }
        private void createExcel(DataTable dt, string path)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Workbooks.Add();
                Excel._Worksheet workSheet = (Excel._Worksheet)excelApp.ActiveSheet;
                workSheet.Name = "Structure";
                for (int i = 0; i < dt.Columns.Count - 1; i++)
                {
                    workSheet.Cells[1, (i + 1)] = dt.Columns[i].ColumnName;
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {

                    for (int j = 0; j < dt.Columns.Count - 1; j++)
                    {
                        if (dt.Rows[i]["Action"].ToString() != "")
                        {
                            int num;
                            bool isNum = int.TryParse(dt.Rows[i]["Data"].ToString(), out num);
                            if (isNum)
                            {
                                workSheet.Cells[(i + 2), (8)] = "'" + dt.Rows[i]["Data"].ToString();
                            }
                            else
                            {
                                workSheet.Cells[(i + 2), (8)] = dt.Rows[i]["Data"].ToString();
                            }

                            workSheet.Cells[(i + 2), (j + 1)] = dt.Rows[i][j];
                        }
                        else
                        {
                            workSheet.Cells[(i + 2), (j + 1)] = dt.Rows[i][j];
                        }
                    }
                }
                Excel.Worksheet dataSheet = (Excel.Worksheet)excelApp.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                dataSheet.Name = "Data";
                dataSheet.Cells[1, 1] = "InputData";
                dataSheet.Cells[1, 2] = "TestCase";
                dataSheet.Cells[2, 1] = "Y";
                int c = 1;
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["FieldName"].ToString() != "")
                    {
                        dataSheet.Cells[1, c + 2] = dt.Rows[i]["FieldName"].ToString();
                        dataSheet.Cells[2, c + 2] = dt.Rows[i]["Data"].ToString();
                        c = c + 1;
                    }
                }
                if (path != null && path != "")
                {
                    try
                    {
                        excelApp.ActiveWorkbook.SaveAs(path);
                        //workSheet.SaveAs(path);
                        excelApp.Quit();
                        MessageBox.Show("Excel file saved!");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                            + ex.Message);
                    }
                }
                else    // no filepath is given
                {
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {

            }
        }
        private void appendExcel(DataTable dt, string path)
        {
            Excel.Application excelApp = new Excel.Application();
            try
            {

                Excel.Workbook workBook = excelApp.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Excel.Worksheets workSheets = (Excel.Worksheets)workBook.Worksheets;
                Excel._Worksheet workSheet = (Excel._Worksheet)workBook.ActiveSheet;
                //Excel.Worksheet workSheet = (Excel.Worksheet)workSheets.get_Item("Structure");
                Excel.Range rng = workSheet.UsedRange;
                int colCount = rng.Columns.Count;
                int rowCount = rng.Rows.Count;
                int newRowCount = rowCount + dt.Rows.Count;
                for (int i = 1; i < dt.Rows.Count; i++)
                {

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        workSheet.Cells[(rowCount + i), (j + 1)] = dt.Rows[i][j];

                    }
                }
                if (path != null && path != "")
                {
                    try
                    {
                        workSheet.SaveAs(path);
                        excelApp.Quit();
                        MessageBox.Show("Excel file saved!");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n" + ex.Message);

                    }
                }
                else    // no filepath is given
                {
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel file not saved!");
                excelApp.Quit();
            }

        }
    }
}
