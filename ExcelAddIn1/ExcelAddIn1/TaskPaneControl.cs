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

        private void buysidebutton_Click(object sender, EventArgs e)
        {
            while (tabControl1.TabPages.Count > 1)
            {
                tabControl1.TabPages.RemoveAt(1);
            }
            tabControl1.TabPages.Add(buysideTabPage);
            Worksheet theSheet = Globals.ThisAddIn.Application.Worksheets["ERP_User_Table"];
            deepHideWorkSheet(theSheet);
        }

        private void sellsidebutton_Click(object sender, EventArgs e)
        {
            while (tabControl1.TabPages.Count > 1)
            {
                tabControl1.TabPages.RemoveAt(1);
            }
            tabControl1.TabPages.Add(sellsideTabPage);
            Worksheet theSheet = Globals.ThisAddIn.Application.Worksheets["ERP_User_Table"];
            deepHideWorkSheet(theSheet);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void showUserButton_Click(object sender, EventArgs e)
        {
            List<string> names = new List<string>();
            foreach (Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                names.Add(worksheet.Name);
            }
            //Worksheet userTable = (Worksheet)Globals.ThisAddIn.Application.Sheets["ERP_User_Table"];
            var userTable = Globals.ThisAddIn.Application.Worksheets.Cast<Worksheet>()
                                   .SingleOrDefault(w => w.Name == "ERP_User_Table");
            if (userTable == null)
            {
                Worksheet newsheet = Globals.ThisAddIn.Application.Worksheets.Add();
                newsheet.Name = "ERP_User_Table";

                var rng = newsheet.Range[newsheet.Cells[1, 1], newsheet.Cells[3, 3]];
                
                rng.Value = new string[,] { { "Id", "password", "Sheet 1", "Sheet 2" }, { "zhang", "zhang123", "True", "True" }, { "Li", "Li123", "False", "True" } };
            }
            //foreach (Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            //{
            //    names.Add(worksheet.Name);
            //    //NamedRange1.Offset[index, 0].Value2 = displayWorksheet.Name;
            //    //index++;
            //}            //if (!names.Contains("ERP_User_Table"))
            //{
            //    Worksheet newsheet = Globals.ThisAddIn.Application.Worksheets.Add();
            //    newsheet.Name = "ERP_User_Table";
            //    var rng = newsheet.Range[newsheet.Cells[1, 1], newsheet.Cells[3, 3]];
            //    rng.Value = new string[,] { { "Id", "password","Sheet 1", "Sheet 2" }, { "zhang", "zhang123", "True", "True" }, { "Li", "Li123", "False", "True" } };
            //}

            else
            {
                Worksheet theSheet = Globals.ThisAddIn.Application.Worksheets["ERP_User_Table"];
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
            //Globals.ThisAddIn.Application.ActiveWorkbook.Protect("1111");
            Worksheet worksheet = Globals.ThisAddIn.Application.Worksheets["ERP_User_Table"];
            worksheet.Protect("1111");
            MessageBox.Show("锁定/解锁完成");
        }

        private void unlockbutton_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect("1111");
            MessageBox.Show("解锁完成");
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
            Worksheet worksheet = Globals.ThisAddIn.Application.Worksheets["ERP_User_Table"];
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
    }
}
