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

        private void button3_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.TaskPane.Visible = false;
            Globals.ThisAddIn.LogInPane.Visible = true;
        }
    }
}
