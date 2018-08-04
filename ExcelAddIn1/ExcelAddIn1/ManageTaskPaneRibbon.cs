using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1
{
    public partial class ManageTaskPaneRibbon
    {
        private void ManageTaskPaneRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            var taskpane = TaskPaneManager.GetTaskPane("A", "ERP Excel 助手", () => new TaskPaneControl());
            taskpane.Visible = ((RibbonToggleButton)sender).Checked;
            var tpc = (TaskPaneControl)taskpane.Control;
            Dictionary<string, string> p = ThisAddIn.getPermission(Constants.guest);
            if (p != null)
            {
                tpc.login(Constants.guest);
                tpc.SetUserLabel(Constants.guest);
            }
        }
    }
}
