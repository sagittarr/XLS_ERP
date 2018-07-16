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
            tabControl1.TabPages.Remove(tabPage2);
            tabControl1.TabPages.Remove(tabPage3);
        }

        private void button1_Click(object sender, EventArgs e)
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
            //MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[2].Range["A1"].Value2.ToString());
            if (visibleSheet > 1)
            {
                try
                {
                    activeWorksheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    MessageBox.Show("Add-in has no permission to modify WorkBook's structure.");
                }
            }
            else
            {
                MessageBox.Show("visible sheet should >= one.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
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
                MessageBox.Show("Add-in has no permission to modify WorkBook's structure.");
            }

        }

        //private void button3_Click(object sender, EventArgs e)
        //{
        //    Globals.ThisAddIn.TaskPane.Visible = false;
        //    Globals.ThisAddIn.LogInPane.Visible = true;
        //}

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
            //MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[2].Range["A1"].Value2.ToString());
            if (visibleSheet > 1)
            {
                try
                {
                    activeWorksheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    MessageBox.Show("Add-in has no permission to modify WorkBook's structure.");
                }
            }
            else
            {
                MessageBox.Show("visible sheet should >= one.");
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
                MessageBox.Show("Add-in has no permission to modify WorkBook's structure.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Worksheet activeWorksheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            //List<string> names = new List<string>();
            //foreach (Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            //{
            //    names.Add(worksheet.Name);
            //    //NamedRange1.Offset[index, 0].Value2 = displayWorksheet.Name;
            //    //index++;
            //}
            //if (!names.Contains("ERP_User_Table"))
            //{
            //    Worksheet newsheet = Globals.ThisAddIn.Application.Worksheets.Add();
            //    newsheet.Name = "ERP_User_Table";
            //    var rng = newsheet.Range[newsheet.Cells[1, 1], newsheet.Cells[3, 3]];
            //    rng.Value = new string[,] { { "--", "Sheet 1", "Sheet 2" }, { "User A", "True", "True" }, { "User B", "False", "True" } };
            //}
            while (tabControl1.TabPages.Count > 1)
            {
                tabControl1.TabPages.RemoveAt(1);
            }
            //MessageBox.Show(names.Count.ToString());

            tabControl1.TabPages.Add(tabPage2);
            tabControl1.SelectedTab = tabPage2;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            while (tabControl1.TabPages.Count > 1)
            {
                tabControl1.TabPages.RemoveAt(1);
            }
            tabControl1.TabPages.Add(tabPage2);
            //tabControl1.TabPages.Remove(tabPage3);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            while (tabControl1.TabPages.Count > 1)
            {
                tabControl1.TabPages.RemoveAt(1);
            }
            tabControl1.TabPages.Add(tabPage3);
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
                foreach (Worksheet sheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    if(sheet.Name == "ERP_User_Table")
                    {
                        sheet.Select();
                        break;
                    }
                }
            }
        }
    }
}
