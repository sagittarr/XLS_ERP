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
                //NamedRange1.Offset[index, 0].Value2 = displayWorksheet.Name;
                //index++;
            }
            if (!names.Contains("ERP_User_Table"))
            {
                Worksheet newsheet = Globals.ThisAddIn.Application.Worksheets.Add();
                newsheet.Name = "ERP_User_Table";
                var rng = newsheet.Range[newsheet.Cells[1, 1], newsheet.Cells[3, 3]];
                rng.Value = new string[,] { { "用户Id", "Sheet 1", "Sheet 2" }, { "zhang", "True", "True" }, { "Li", "False", "True" } };
            }
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
            Globals.ThisAddIn.Application.ActiveWorkbook.Protect("1111");
            MessageBox.Show("锁定/解锁完成");
        }

        private void unlockbutton_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect("1111");
            MessageBox.Show("解锁完成");
        }
    }
}
