﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private TaskPaneControl taskPaneControl1;
        private UserLogInControl userloginControl;
        private Microsoft.Office.Tools.CustomTaskPane taskPaneValue;
        private Microsoft.Office.Tools.CustomTaskPane userloginPaneValue;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            taskPaneControl1 = new TaskPaneControl();
            taskPaneValue = this.CustomTaskPanes.Add(
    taskPaneControl1, "MyCustomTaskPane");
            taskPaneValue.VisibleChanged +=
                new EventHandler(taskPaneValue_VisibleChanged);

            userloginControl = new UserLogInControl();
            userloginPaneValue = this.CustomTaskPanes.Add(
    taskPaneControl1, "MyCustomTaskPane");
            userloginPaneValue.VisibleChanged +=
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
        public Microsoft.Office.Tools.CustomTaskPane LogInPane 
        {
            get
            {
                return userloginPaneValue;
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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