using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelAddIn1
{
    public partial class TaskPaneControl : UserControl
    {
        public TaskPaneControl()
        {
            InitializeComponent();
            tabControl1.TabPages.Remove(managerTabPage);
            tabControl1.TabPages.Remove(buysideTabPage);
            tabControl1.TabPages.Remove(sellsideTabPage);
        }
        private string PROTECTED_ERROR_MESSAGE = "Add-in has no permission to modify WorkBook's structure.";
        private string VISIBLE_SHEET_LESS_THAN_TWO_MESSAGE = "visible sheet should be at least more than one.";
        private string UserPasswordTable = "UserPasswordTable";
        private string UserPermissionTable = "UserPermissionTable";
        private string PermissionOperation = "Writable,ReadOnly,Invisible";
        private string Writable = "Writable";
        private string ReadOnly = "ReadOnly";
        private string Invisible = "Invisible";
        private const string key = "1234";
        private void deepHideWorkSheet(Worksheet theSheet)
        {
            if (theSheet == null) return;
            try
            {
                theSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(PROTECTED_ERROR_MESSAGE);
            }
        }
        private void unHideWorkSheet(Worksheet theSheet)
        {
            if (theSheet == null) return;
            try
            {
                theSheet.Visible = XlSheetVisibility.xlSheetVisible;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(PROTECTED_ERROR_MESSAGE);
            }
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Worksheet activeWorksheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            int visibleSheet = 0;

            foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                if (sheet.Visible == XlSheetVisibility.xlSheetVisible)
                {
                    visibleSheet += 1;
                }
            }
            if (visibleSheet > 1)
            {
                deepHideWorkSheet(activeWorksheet);
            }
            else
            {
                MessageBox.Show(VISIBLE_SHEET_LESS_THAN_TWO_MESSAGE);
            }
        }

        private void Button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                {
                    sheet.Visible = XlSheetVisibility.xlSheetVisible;
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(PROTECTED_ERROR_MESSAGE);
            }
        }

        private void managerbutton_Click(object sender, EventArgs e)
        {
            while (tabControl1.TabPages.Count > 1)
            {
                tabControl1.TabPages.RemoveAt(1);
            }
            tabControl1.TabPages.Add(managerTabPage);
            tabControl1.SelectedTab = managerTabPage;

        }

        //private void buysidebutton_Click(object sender, EventArgs e)
        //{
        //    while (tabControl1.TabPages.Count > 1)
        //    {
        //        tabControl1.TabPages.RemoveAt(1);
        //    }
        //    tabControl1.TabPages.Add(buysideTabPage);
        //    Worksheet theSheet = Globals.ThisAddIn.Application.Worksheets["ERP_User_Table"];
        //    deepHideWorkSheet(theSheet);
        //}

        //private void sellsidebutton_Click(object sender, EventArgs e)
        //{
        //    while (tabControl1.TabPages.Count > 1)
        //    {
        //        tabControl1.TabPages.RemoveAt(1);
        //    }
        //    tabControl1.TabPages.Add(sellsideTabPage);
        //    Worksheet theSheet = Globals.ThisAddIn.Application.Worksheets["ERP_User_Table"];
        //    deepHideWorkSheet(theSheet);
        //}

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void showUserButton_Click(object sender, EventArgs e)
        {
            var userTable = Globals.ThisAddIn.Application.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == UserPasswordTable);
            if (userTable == null)
            {
                Worksheet newsheet = Globals.ThisAddIn.Application.Worksheets.Add();
                newsheet.Name = UserPasswordTable;
                var rng = newsheet.Range[newsheet.Cells[1, 1], newsheet.Cells[2, 2]];
                string[,] values = new string[2, 2];
                values[0, 0] = "ID";
                values[0, 1] = "PASSWORD";
                values[1, 0] = "superuser";
                values[1, 1] = "su2018";
                rng.Value =values;
            }
            else
            {
                Worksheet theSheet = Globals.ThisAddIn.Application.Worksheets[UserPasswordTable];
                unHideWorkSheet(theSheet);
                try
                {
                    theSheet.Select();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    MessageBox.Show(PROTECTED_ERROR_MESSAGE);
                }
            }
            //Protect(getPasswordFromUser, missing, missing);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Worksheet worksheet = Globals.ThisAddIn.Application.Worksheets["ERP_User_Table"];
            worksheet.Protect(key);
            MessageBox.Show("锁定/解锁完成");
        }

        private void unlockbutton_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect(key);
            MessageBox.Show("解锁完成");
        }

        private Dictionary<string, string> getPermission(string userName)
        {
            Worksheet worksheet = Globals.ThisAddIn.Application.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == UserPermissionTable);
            Range currentFind;
            Range range = worksheet.Columns["A:A", Type.Missing];
            currentFind = range.Cells.Find(userName, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole,
                            XlSearchOrder.xlByRows, XlSearchDirection.xlNext, true, false, false);
            Range firstRow = worksheet.UsedRange.Rows[1];
            Range theRow = worksheet.UsedRange.Rows[currentFind.Row];
            //StringBuilder sb = new StringBuilder();
            Dictionary<string, string> permission = new Dictionary<string, string>();
            foreach(Range col in theRow.Columns)
            {
                //string v = firstRow.Cells[1, col.Column].value2;
                //sb.Append(v);
                //sb.Append(":"+ theRow.Cells[1,col.Column].value2+';');
                permission.Add(firstRow.Cells[1, col.Column].value2, theRow.Cells[1, col.Column].value2);
            }
            return permission;
            //MessageBox.Show(sb.ToString());
        }

        private void loginbutton_Click(object sender, EventArgs e)
        {
            string username, password = null;
            int entryNumber = -1;
            if(usernameBox.Text!=null && usernameBox.Text.Length > 0 && passwordBox.Text != null && passwordBox.Text.Length > 0)
            {
                username = usernameBox.Text;
                password = passwordBox.Text;
            }
            else
            {
                return;
            }
            //Worksheet worksheet = Globals.ThisAddIn.Application.Worksheets[UserPasswordTable];
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == UserPasswordTable);
            if (worksheet == null)
            {
                Console.WriteLine("ERROR: worksheet == null");
            }
            Range currentFind;
            Range range = worksheet.Columns["A:A", Type.Missing];
            currentFind = range.Cells.Find(username, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole,
                            XlSearchOrder.xlByRows, XlSearchDirection.xlNext, true, false, false);
            if (currentFind != null)
            {
                entryNumber = currentFind.Row;
                //MessageBox.Show(currentFind.Row.ToString());
            }
            else
            {
                MessageBox.Show("Username not found");
                return;
            }

            Range rangeRows = worksheet.Rows["1:1"];
            currentFind = rangeRows.Cells.Find("Password", Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole,
                XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false);
            if(currentFind != null)
            {
                //MessageBox.Show(currentFind.Row.ToString() + "," + currentFind.Column.ToString());
                Range rangeTarget = worksheet.Cells[entryNumber, currentFind.Column];
                if (string.Equals(rangeTarget.Value2, password))
                {
                    MessageBox.Show("Password match!");
                    //Globals.ThisAddIn.login(username);  
                    this.login(username);
                }
                else
                {
                    MessageBox.Show("Password doesn't match!");
                }
            }
            else
            {
                MessageBox.Show("Password column not found");
            }
        }

        private void usernameBox_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void passwordBox_TextChanged(object sender, EventArgs e)
        {
            passwordBox.UseSystemPasswordChar = true;
        }
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        private List<int> getRangeIndex(Range range, bool rowYN)
        {
            List<int> idx = new List<int>();

            if (rowYN)
            {
                foreach (Range v in range.Cells)
                {
                    idx.Add(v.Row);
                }
            }
            else
            {
                foreach (Range v in range.Cells)
                {
                    idx.Add(v.Column);
                }
            }
            return idx;
        }
        private void ManageButton_Click(object sender, EventArgs e)
        {

            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == UserPermissionTable);
            List<string> sheetNames = new List<string>();
            foreach (Worksheet ws in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                sheetNames.Add(ws.Name);
            }

            if (worksheet == null)
            {
                Worksheet theSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                theSheet.Name = UserPermissionTable;
                sheetNames.Add(UserPermissionTable);
                var numOfSheets = sheetNames.Count;
                var rng = theSheet.Range[theSheet.Cells[1, 1], theSheet.Cells[3, numOfSheets + 1]];
                string[,] values = new string[3, numOfSheets + 1];
                values[0, 0] = "ID";
                values[1, 0] = "superuser";
                values[2, 0] = "guest";
                for (var i = 0; i < sheetNames.Count; i++)
                {
                    values[0, i + 1] = sheetNames[i];
                    values[1, i + 1] = Writable;
                    values[2, i + 1] = ReadOnly;
                }
                rng.Value = values;
                theSheet.Range["B:"+ GetExcelColumnName(numOfSheets+1)].Validation.Add(XlDVType.xlValidateList, Type.Missing,XlFormatConditionOperator.xlBetween, PermissionOperation);
                theSheet.Range["1:1"].Validation.Delete();
                //theSheet.Range["A:A"].Validation.Delete();
            }
            else
            {
                Worksheet theSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[UserPermissionTable];
                unHideWorkSheet(theSheet);
                try
                {
                    theSheet.Select();
                    Range firstRow = theSheet.UsedRange.Rows[1];
                    System.Array myvalues = (System.Array)firstRow.Cells.Value;
                    List<string> lst = myvalues.OfType<object>().Select(o => o.ToString()).ToList();
                    List<string> toAdd = new List<string>();
                    foreach(string name in sheetNames)
                    {
                        if (!lst.Contains(name))
                        {
                            toAdd.Add(name);
                        }
                    }
                    //newSheetNames.AddRange(toAdd);
                    foreach(string newColumnName in toAdd)
                    {
                        Range rangeTarget = worksheet.Cells[1, theSheet.UsedRange.Columns.Count + 1];
                        rangeTarget.Value2 = newColumnName;
                    }
                    int totalColumns = theSheet.UsedRange.Columns.Count;
                    //MessageBox.Show(totalColumns.ToString());
                    theSheet.Range["B:" + GetExcelColumnName(totalColumns)].Validation.Delete();
                    theSheet.Range["B:" + GetExcelColumnName(totalColumns)].Validation.Add(XlDVType.xlValidateList, Type.Missing, XlFormatConditionOperator.xlBetween, PermissionOperation);
                    theSheet.Range["1:1"].Validation.Delete();
                    //Range firstColumn = theSheet.UsedRange.Columns[1];
                    //List<int> idx = getRangeIndex(firstColumn, true);
                    //MessageBox.Show(string.Join(",",idx));
                    //Range firstColumn = theSheet.UsedRange.Columns[1];
                    //myvalues = (System.Array)firstColumn.Cells.Value;
                    //lst = myvalues.OfType<object>().Select(o => o.ToString()).ToList();

                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    MessageBox.Show(PROTECTED_ERROR_MESSAGE);
                }
            }
        }

        private void logoutButton_Click(object sender, EventArgs e)
        {
            this.logout();
            //Globals.ThisAddIn.Application.ActiveWorkbook.Protect(key);
        }

        public void SetUserLabel(string text)
        {
            userLabel.Text = text;
        }
        public void login(string username)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect(key);
            ThisAddIn.applyPermission(ThisAddIn.getPermission(username));
            this.SetUserLabel(username);
        }
        public void logout()
        {
            ThisAddIn.applyPermission(ThisAddIn.getPermission("guest"));
            //Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect(key);
            Globals.ThisAddIn.Application.ActiveWorkbook.Protect(key, true);
            Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.Name + Globals.ThisAddIn.Application.ActiveWorkbook.ProtectStructure.ToString());
            this.SetUserLabel("guest");
        }
        //private void protect_Click(object sender, EventArgs e)
        //{
        //    Globals.ThisAddIn.Application.ActiveWorkbook.Protect(key, true);
        //}
    }
}
