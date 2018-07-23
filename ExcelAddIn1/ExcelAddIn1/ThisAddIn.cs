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
        public const string PROTECTED_ERROR_MESSAGE = "Add-in has no permission to modify WorkBook's structure.";
        public const string VISIBLE_SHEET_LESS_THAN_TWO_MESSAGE = "visible sheet should be at least more than one.";
        public const string Writable = "Writable";
        public const string ReadOnly = "ReadOnly";
        public const string Invisible = "Invisible";
        public const string UserPasswordTable = "UserPasswordTable";
        public const string UserPermissionTable = "UserPermissionTable";
        public const string key = "1234";
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
           
        }
        //////
        public void logout()
        {
            ThisAddIn.applyPermission(ThisAddIn.getPermission("guest"));
            //Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect(key);
            Globals.ThisAddIn.Application.ActiveWorkbook.Protect(key,true);
            Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.Name+Globals.ThisAddIn.Application.ActiveWorkbook.ProtectStructure.ToString());
            taskPaneControl1.SetUserLabel("guest");
        }
        public void login(string username)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect(key);
            ThisAddIn.applyPermission(ThisAddIn.getPermission(username));
            taskPaneControl1.SetUserLabel(username);
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
                MessageBox.Show(PROTECTED_ERROR_MESSAGE);
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
                MessageBox.Show(PROTECTED_ERROR_MESSAGE);
            }
        }
        public static Dictionary<string, string> getPermission(string userName)
        {
            //Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Excel.Worksheet>().SingleOrDefault(w => w.Name == UserPermissionTable);
            if (worksheet == null) return null;
            Range currentFind;
            Range range = worksheet.Columns["A:A", Type.Missing];
            currentFind = range.Cells.Find(userName, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole,
                            XlSearchOrder.xlByRows, XlSearchDirection.xlNext, true, false, false);
            Range firstRow = worksheet.UsedRange.Rows[1];
            Range theRow = worksheet.UsedRange.Rows[currentFind.Row];
            //StringBuilder sb = new StringBuilder();
            Dictionary<string, string> permission = new Dictionary<string, string>();
            foreach (Range col in theRow.Columns)
            {
                //string v = firstRow.Cells[1, col.Column].value2;
                //sb.Append(v);
                //sb.Append(":"+ theRow.Cells[1,col.Column].value2+';');
                permission.Add(firstRow.Cells[1, col.Column].value2, theRow.Cells[1, col.Column].value2);
            }
            return permission;
            //MessageBox.Show(sb.ToString());
        }
        public static void applyPermission(Dictionary<string, string> permission)
        {
            if (permission == null) return;
            foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                if (permission.ContainsKey(ws.Name))
                {
                    if (permission[ws.Name] == Invisible)
                    {
                        deepHideWorkSheet(ws);
                    }
                    else if (permission[ws.Name] == ReadOnly)
                    {
                        unHideWorkSheet(ws);
                        ws.Protect(key);
                    }
                    else
                    {
                        unHideWorkSheet(ws);
                        ws.Unprotect(key);
                    }
                }
            }
            
        }
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
