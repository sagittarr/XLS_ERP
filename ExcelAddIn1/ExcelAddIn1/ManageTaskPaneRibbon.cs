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
            Globals.ThisAddIn.TaskPane.Visible = ((RibbonToggleButton)sender).Checked;
            Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect(ThisAddIn.key);
            Dictionary<string, string> p = ThisAddIn.getPermission("guest");
            if (p != null)
                Globals.ThisAddIn.login("guest");
        }
    }
}
