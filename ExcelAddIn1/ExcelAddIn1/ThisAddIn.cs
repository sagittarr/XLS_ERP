using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private TaskPaneControl taskPaneControl1;
        private Microsoft.Office.Tools.CustomTaskPane taskPaneValue;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            taskPaneControl1 = new TaskPaneControl();
            taskPaneValue = this.CustomTaskPanes.Add(
    taskPaneControl1, "ERP Excel 助手");
            taskPaneValue.VisibleChanged +=
                new EventHandler(taskPaneValue_VisibleChanged);
        }
        private void taskPaneValue_VisibleChanged(object sender, System.EventArgs e)
        {
            Globals.Ribbons.ManageTaskPaneRibbon.toggleButton1.Checked =
                taskPaneValue.Visible;
        }
        private void userloginPaneValue_VisibleChanged(object sender, System.EventArgs e)
        {
            //Globals.Ribbons.ManageTaskPaneRibbon.toggleButton1.Checked =
            //    taskPaneValue.Visible;
        }
        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return taskPaneValue;
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            MessageBox.Show("Exit.");
        }
        public static void deepHideWorkSheet(Excel.Worksheet theSheet)
        {
            if (theSheet == null) return;
            try
            {
                theSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(Constants.PROTECTED_ERROR_MESSAGE);
            }
        }
        public static void unHideWorkSheet(Excel.Worksheet theSheet)
        {
            if (theSheet == null) return;
            try
            {
                theSheet.Visible = XlSheetVisibility.xlSheetVisible;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(Constants.PROTECTED_ERROR_MESSAGE);
            }
        }
        public static Dictionary<string, string> getPermission(string userName)
        {
            //Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets
            if (userName == Constants.root) return null;
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Excel.Worksheet>().SingleOrDefault(w => w.Name == Constants.UserPermissionTable);
            if (worksheet == null) return null;
            Range currentFind;
            Range range = worksheet.Columns["A:A", Type.Missing];
            currentFind = range.Cells.Find(userName, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole,
                            XlSearchOrder.xlByRows, XlSearchDirection.xlNext, true, false, false);
            Range firstRow = worksheet.UsedRange.Rows[1];
            if(currentFind == null)
            {
                MessageBox.Show(userName + " is not found in permission table.");
                return null;
            }
            Range theRow = worksheet.UsedRange.Rows[currentFind.Row];
            //StringBuilder sb = new StringBuilder();
            Dictionary<string, string> permission = new Dictionary<string, string>();
            foreach (Range col in theRow.Columns)
            {
                permission.Add(firstRow.Cells[1, col.Column].value2, theRow.Cells[1, col.Column].value2);
            }
            return permission;
        }
        public static void applyPermission(Dictionary<string, string> permission, bool isRootUser = false)
        {
            try
            {
                if (isRootUser == true)
                {
                    Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect(Constants.key);
                    foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                    {
                       unHideWorkSheet(ws);
                       ws.Unprotect(Constants.key);
                    }
                    return;
                }
                if (permission == null) return;
                Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect(Constants.key);
                foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                {
                    if (permission.ContainsKey(ws.Name))
                    {
                        if (permission[ws.Name] == Constants.Invisible)
                        {
                            deepHideWorkSheet(ws);
                        }
                        else if (permission[ws.Name] == Constants.ReadOnly)
                        {
                            unHideWorkSheet(ws);
                            ws.Protect(Constants.key);
                        }
                        else
                        {
                            unHideWorkSheet(ws);
                            ws.Unprotect(Constants.key);
                        }
                    }
                }
                if (!permission.ContainsKey(Constants.structure) || permission[Constants.structure] != Constants.Mutable)
                {
                   Globals.ThisAddIn.Application.ActiveWorkbook.Protect(Constants.key, true);
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(Constants.PROTECTED_ERROR_MESSAGE);
            }
        }
        //private static object[,] ReadValues(Excel.Worksheet sheet, int lastRow, int lastColumn)
        //{
        //    object[,] cellValues;
        //    var firstCell = sheet.get_Range("A1", Type.Missing);
        //    var lastCell = (Excel.Range)sheet.Cells[lastRow, lastColumn];

        //    if (lastRow == 1 && lastColumn == 1)
        //    {
        //        cellValues = new object[2, 2];
        //        cellValues[1, 1] = firstCell.Value2;
        //    }
        //    else
        //    {
        //        Excel.Range worksheetCells = sheet.get_Range(firstCell, lastCell);
        //        cellValues = worksheetCells.Value2 as object[,];
        //    }

        //    return cellValues;
        //}
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
