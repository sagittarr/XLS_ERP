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
//using System.Data;

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
        private void deepHideWorkSheet(Worksheet theSheet)
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
        private void unHideWorkSheet(Worksheet theSheet)
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
                MessageBox.Show(Constants.VISIBLE_SHEET_LESS_THAN_TWO_MESSAGE);
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
                MessageBox.Show(Constants.PROTECTED_ERROR_MESSAGE);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void showUserButton_Click(object sender, EventArgs e)
        {
            try
            {
                var userTable = Globals.ThisAddIn.Application.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == Constants.UserPasswordTable);
                if (userTable == null)
                {
                    Worksheet newsheet = Globals.ThisAddIn.Application.Worksheets.Add();
                    newsheet.Name = Constants.UserPasswordTable;
                    var rng = newsheet.Range[newsheet.Cells[1, 1], newsheet.Cells[2, 2]];
                    string[,] values = new string[2, 2];
                    values[0, 0] = "ID";
                    values[0, 1] = "PASSWORD";
                    values[1, 0] = "admin";
                    values[1, 1] = "su2018";
                    rng.Value = values;
                }
                else
                {
                    Worksheet theSheet = Globals.ThisAddIn.Application.Worksheets[Constants.UserPasswordTable];
                    unHideWorkSheet(theSheet);

                    theSheet.Select();

                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(Constants.PROTECTED_ERROR_MESSAGE);
            }
        }

        private Dictionary<string, string> getPermission(string userName)
        {
            Worksheet worksheet = Globals.ThisAddIn.Application.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == Constants.UserPermissionTable);
            Range currentFind;
            Range range = worksheet.Columns["A:A", Type.Missing];
            currentFind = range.Cells.Find(userName, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole,
                            XlSearchOrder.xlByRows, XlSearchDirection.xlNext, true, false, false);
            Range firstRow = worksheet.UsedRange.Rows[1];
            Range theRow = worksheet.UsedRange.Rows[currentFind.Row];
            Dictionary<string, string> permission = new Dictionary<string, string>();
            foreach (Range col in theRow.Columns)
            {
                permission.Add(firstRow.Cells[1, col.Column].value2, theRow.Cells[1, col.Column].value2);
            }
            return permission;
        }

        private void loginbutton_Click(object sender, EventArgs e)
        {
            try
            {
                string username, password = null;
                int entryNumber = -1;
                if (usernameBox.Text != null && usernameBox.Text.Length > 0 && passwordBox.Text != null && passwordBox.Text.Length > 0)
                {
                    username = usernameBox.Text;
                    password = passwordBox.Text;
                }
                else
                {
                    return;
                }
                if (username == Constants.root && password == "root2018")
                {
                    MessageBox.Show(username + " 登录成功");
                    //Globals.ThisAddIn.login(username);  
                    this.login(username);
                    return;
                }
                //Worksheet worksheet = Globals.ThisAddIn.Application.Worksheets[UserPasswordTable];
                Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == Constants.UserPasswordTable);
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
                if (currentFind != null)
                {
                    //MessageBox.Show(currentFind.Row.ToString() + "," + currentFind.Column.ToString());
                    Range rangeTarget = worksheet.Cells[entryNumber, currentFind.Column];
                    if (string.Equals(rangeTarget.Value2, password))
                    {
                        MessageBox.Show(username + " 登录成功");
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
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(Constants.PROTECTED_ERROR_MESSAGE);
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
            try
            {
                Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == Constants.UserPermissionTable);
                List<string> sheetNames = new List<string>();
                foreach (Worksheet ws in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                {
                    sheetNames.Add(ws.Name);
                }

                if (worksheet == null)
                {
                    Worksheet theSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                    theSheet.Name = Constants.UserPermissionTable;
                    sheetNames.Add(Constants.UserPermissionTable);
                    int columnOffset = 5;
                    var numOfSheets = sheetNames.Count;
                    var rng = theSheet.Range[theSheet.Cells[1, 1], theSheet.Cells[3, numOfSheets + columnOffset]];
                    string[,] values = new string[3, numOfSheets + columnOffset];
                    values[0, 0] = "ID";
                    values[1, 0] = "admin";
                    values[2, 0] = Constants.guest;

                    values[0, 1] = Constants.structure;
                    values[1, 1] = Constants.Mutable;
                    values[2, 1] = Constants.InMutable;

                    values[0, 2] = Constants.managementTab;
                    values[1, 2] = Constants.Visible;
                    values[2, 2] = Constants.Invisible;

                    values[0, 3] = Constants.buysideTab;
                    values[1, 3] = Constants.Visible;
                    values[2, 3] = Constants.Invisible;

                    values[0, 4] = Constants.sellsideTab;
                    values[1, 4] = Constants.Visible;
                    values[2, 4] = Constants.Invisible;

                    for (var i = 0; i < sheetNames.Count; i++)
                    {
                        values[0, i + columnOffset] = sheetNames[i];
                        values[1, i + columnOffset] = Constants.Writable;
                        values[2, i + columnOffset] = Constants.ReadOnly;
                    }
                    rng.Value = values;
                    theSheet.Range["B:B"].Validation.Add(XlDVType.xlValidateList, Type.Missing, XlFormatConditionOperator.xlBetween, Constants.Mutable + "," + Constants.InMutable);
                    theSheet.Range["B:B"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
                    theSheet.Range["C:E"].Validation.Add(XlDVType.xlValidateList, Type.Missing, XlFormatConditionOperator.xlBetween, Constants.Visible + "," + Constants.Invisible);
                    theSheet.Range["C:E"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Wheat);
                    theSheet.Range["F:" + GetExcelColumnName(numOfSheets + columnOffset)].Validation.Add(XlDVType.xlValidateList, Type.Missing, XlFormatConditionOperator.xlBetween, Constants.PermissionOperation);
                    theSheet.Range["F:" + GetExcelColumnName(numOfSheets + columnOffset)].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    theSheet.Range["1:1"].Validation.Delete();
                    //theSheet.Range["A:A"].Validation.Delete();
                }
                else
                {
                    Worksheet theSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[Constants.UserPermissionTable];
                    unHideWorkSheet(theSheet);

                    theSheet.Select();
                    Range firstRow = theSheet.UsedRange.Rows[1];
                    System.Array myvalues = (System.Array)firstRow.Cells.Value;
                    List<string> lst = myvalues.OfType<object>().Select(o => o.ToString()).ToList();
                    List<string> toAdd = new List<string>();
                    foreach (string name in sheetNames)
                    {
                        if (!lst.Contains(name))
                        {
                            toAdd.Add(name);
                        }
                    }
                    //newSheetNames.AddRange(toAdd);
                    foreach (string newColumnName in toAdd)
                    {
                        Range rangeTarget = worksheet.Cells[1, theSheet.UsedRange.Columns.Count + 1];
                        rangeTarget.Value2 = newColumnName;
                    }
                    int totalColumns = sheetNames.Count + 5;
                    //MessageBox.Show(totalColumns.ToString());
                    theSheet.Range["B:" + GetExcelColumnName(totalColumns)].Validation.Delete();
                    theSheet.Range["B:B"].Validation.Add(XlDVType.xlValidateList, Type.Missing, XlFormatConditionOperator.xlBetween, Constants.Mutable + "," + Constants.InMutable);
                    theSheet.Range["B:B"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
                    theSheet.Range["C:E"].Validation.Add(XlDVType.xlValidateList, Type.Missing, XlFormatConditionOperator.xlBetween, Constants.Visible + "," + Constants.Invisible);
                    theSheet.Range["C:E"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Wheat);
                    theSheet.Range["F:" + GetExcelColumnName(totalColumns)].Validation.Add(XlDVType.xlValidateList, Type.Missing, XlFormatConditionOperator.xlBetween, Constants.PermissionOperation);
                    theSheet.Range["F:" + GetExcelColumnName(totalColumns)].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    theSheet.Range["1:1"].Validation.Delete();

                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(Constants.PROTECTED_ERROR_MESSAGE);
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
            var permission = ThisAddIn.getPermission(username);
            ThisAddIn.applyPermission(permission, username == Constants.root);
            this.SetUserLabel(username);
            showTabForUser(username, permission);
            CheckSheets();
            //checkedListBox1.Items
        }
        public void logout()
        {
            ThisAddIn.applyPermission(ThisAddIn.getPermission(Constants.guest));
            //Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect(key);
            Globals.ThisAddIn.Application.ActiveWorkbook.Protect(Constants.key, true);
            Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            if (Globals.ThisAddIn.Application.ActiveWorkbook.ProtectStructure)
            {
                MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.Name + " 退出登录成功");
            }
            this.SetUserLabel(Constants.guest);
            removeDupTabs();
        }
        private void removeDupTabs()
        {
            while (tabControl1.TabPages.Count > 1)
            {
                tabControl1.TabPages.RemoveAt(1);
            }
        }
        //private void showManagerTab()
        //{
        //    removeDupTabs();
        //    tabControl1.TabPages.Add(managerTabPage);
        //    tabControl1.SelectedTab = managerTabPage;
        //}

        private void showTabForUser(string username, Dictionary<string, string> permission = null)
        {
            if (username == Constants.root)
            {
                removeDupTabs();
                tabControl1.TabPages.Add(managerTabPage);
                tabControl1.TabPages.Add(sellsideTabPage);
                tabControl1.TabPages.Add(buysideTabPage);
                tabControl1.SelectedTab = managerTabPage;
                return;
            }
            removeDupTabs();
            if (permission != null && permission.ContainsKey(Constants.managementTab) && permission[Constants.managementTab] == Constants.Visible)
            {
                tabControl1.TabPages.Add(managerTabPage);
            }
            if (permission != null && permission.ContainsKey(Constants.sellsideTab) && permission[Constants.sellsideTab] == Constants.Visible)
            {
                tabControl1.TabPages.Add(sellsideTabPage);
            }
            if (permission != null && permission.ContainsKey(Constants.buysideTab) && permission[Constants.buysideTab] == Constants.Visible)
            {
                tabControl1.TabPages.Add(buysideTabPage);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
 
        }
        
        string ConvertObjectToString(object obj)
        {
            return obj?.ToString() ?? string.Empty;
        }

        public System.Data.DataTable toDataTable(List<List<string>> lists)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            foreach (string colStr in lists[0])
            {
                var colStrVal = colStr;
                DataColumn dataColumn = new DataColumn(colStrVal);
                dataColumn.DataType = System.Type.GetType("System.String");
                table.Columns.Add(dataColumn);
            }

            foreach (var rowInput in lists.Skip(1))
            {
                DataRow row = table.NewRow();
                for (var i = 0; i < lists[0].Count; i++)
                {
                    if (table.Columns.Contains(lists[0][i]))
                    {
                        row[lists[0][i]] = rowInput[i];
                    }
                    else
                    {
                        MessageBox.Show("'" + lists[0][i] + "' is not found");
                    }
                }
                table.Rows.Add(row);
            }
            return table;
        }

        public Dictionary<String, Dictionary<string, string>> ToDict(List<List<string>> table, string keyColumn)
        {
            Dictionary<String, Dictionary<string, string>> rangeDict = new Dictionary<String, Dictionary<string, string>>();
            if (table.Count > 1)
            {
                List<String> headers = table[0];
                int indexOfKey = headers.IndexOf(keyColumn);
                if (indexOfKey == -1)
                {
                    MessageBox.Show("'" + keyColumn + "' is not found");
                    return null;
                }
                for (var i = 1; i < table.Count; i++)
                {
                    string key = table[i][indexOfKey];
                    Dictionary<string, string> rowDict = new Dictionary<string, string>();
                    for (var j = 0; j < headers.Count; j++)
                    {
                        rowDict[headers[j]] = table[i][j];
                    };
                    //MessageBox.Show(string.Join(";", rowDict.Select(x => x.Key + "=" + x.Value).ToArray()));
                    rangeDict[key] = rowDict;
                }
            }
            return rangeDict;

        }
        public List<List<string>> GetFormula(Range range, int columns)
        {
            object[,] values = range.Formula as object[,];

            //firstColumn.t
            List<List<string>> table = new List<List<string>>();
            List<string> list = new List<string>();
            foreach (object o in values)
            {
                if (list.Count == columns)
                {
                    table.Add(list);
                    list = new List<string>();
                }
                if (o == null)
                {
                    list.Add("");
                }
                else
                {
                    list.Add(o.ToString());
                }
            }
            if (list.Count != 0)
            {
                table.Add(list);
            }
            return table;
        }
        public List<List<string>> ToValues(Range range, int columns)
        {
            object[,] values = range.Value2 as object[,];

            //firstColumn.t
            List<List<string>> table = new List<List<string>>();
            List<string> list = new List<string>();
            foreach (object o in values)
            {
                if (list.Count == columns)
                {
                    table.Add(list);
                    list = new List<string>();
                }
                if (o == null)
                {
                    list.Add("");
                }
                else
                {
                    list.Add(o.ToString());
                }
            }
            if (list.Count != 0)
            {
                table.Add(list);
            }
            return table;
        }

        private void InputButton_Click(object sender, EventArgs e)
        {
            Worksheet orderInput = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == "JUN2");
            //Range range1 = orderInput.UsedRange.Rows;
            //List<List<string>> table1 = ToValues(range1, orderInput.UsedRange.Columns.Count);
            //orderInput.get_Range("C:C").EntireColumn.Hidden = true;
            string columnsToHide = "C,E,G,H,I,K,L,N,O,Q,T";
            string[] cols = columnsToHide.Split(',');
            foreach (string col in cols)
            {

                orderInput.UsedRange.Columns[col + ":" + col, Type.Missing].Hidden = true;
            }
            //orderInput.UsedRange.Columns["C:C", Type.Missing].Hidden = true;
            //orderInput.UsedRange.Columns["E:E", Type.Missing].Hidden = true;
            //orderInput.get_Range("E:G").Columns.Hidden = true;
            //MessageBox.Show(orderInput.UsedRange.Columns.Count.ToString());


        }

        private void button3_Click_1(object sender, EventArgs e)
        {

        }

        private void showColumnsButton_Click(object sender, EventArgs e)
        {
            Worksheet orderInput = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == "JUN2");
            orderInput.UsedRange.Columns.Hidden = false;
        }

        private void showColumnPermission_Click(object sender, EventArgs e)
        {

        }

        private void aggregationButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!CheckSheets())
                {
                    return;
                }
                Worksheet orderInputSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == "订单输入");
                Range range1 = orderInputSheet.UsedRange.Rows;
                List<List<string>> orderInputLists = ToValues(range1, orderInputSheet.UsedRange.Columns.Count);
                if (orderInputLists!=null && orderInputLists.Count < 1)
                {
                    MessageBox.Show("请确认<订单输入>表单至少有一行数据");
                    return;
                }
                if(orderInputLists[0]==null || !orderInputLists[0].Contains("OrderNo"))
                {
                    MessageBox.Show("请确认<订单输入>表单有<OrderNo>列");
                    return;
                }
                Dictionary<String, Dictionary<string, string>> orderInputDict = ToDict(orderInputLists, "OrderNo");


                Worksheet receiptSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == "收款输入");
                Range range2 = receiptSheet.UsedRange.Rows;
                List<List<string>> receiptLists = ToValues(range2, receiptSheet.UsedRange.Columns.Count);
                if (receiptLists != null && receiptLists.Count < 1)
                {
                    MessageBox.Show("请确认<收款输入>表单至少有一行数据");
                    return;
                }
                if (receiptLists[0] == null || !receiptLists[0].Contains("OrderNo"))
                {
                    MessageBox.Show("请确认<收款输入>表单有<OrderNo>列");
                    return;
                }
                Worksheet outputSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == "汇总结果");
                Range range3 = outputSheet.UsedRange.Rows;
                int width = outputSheet.UsedRange.Columns.Count;
                List<List<string>> outputLists = ToValues(outputSheet.UsedRange.Columns, outputSheet.UsedRange.Columns.Count);
                if (outputLists.Count < 4)
                {
                    MessageBox.Show("请确认当前表单至少有四行数据");
                    return;
                }
                string[] sources = outputLists[0].ToArray();
                string[] headers = outputLists[1].ToArray();
                List<string> formats = outputLists[2];
                if (sources.Length != headers.Length)
                {
                    MessageBox.Show("请确认第一行与第二行的列数相等。");
                    return;
                }
                outputSheet.Range["5:" + outputLists.Count.ToString()].Delete();
                System.Data.DataTable receiptTable = toDataTable(receiptLists);
                foreach (string orderNumber in orderInputDict.Keys)
                {
                    var orderInputRow = orderInputDict[orderNumber];
                    List<object> outputRow = new List<object>();
                    string expression;
                    expression = "OrderNo =" + "'" + orderNumber + "'";
                    DataRow[] foundRows;
                    foundRows = receiptTable.Select(expression);
                    foreach (DataRow r in foundRows)
                    {
                        for (var i = 0; i < sources.Length; i++)
                        {
                            if (sources[i] == receiptSheet.Name)
                            {
                                if (receiptTable.Columns.Contains(headers[i]))
                                {
                                    outputRow.Add(r[headers[i]]);
                                    continue;
                                }
                            }
                            outputRow.Add("");
                        }
                        var newRow = outputSheet.UsedRange.Rows.Count + 1;
                        var rng = outputSheet.Range[outputSheet.Cells[newRow, 1], outputSheet.Cells[newRow, width]];
                        //string[] result = Array.ConvertAll<object, string>(outputRow.ToArray(), ConvertObjectToString);
                        //rng.Value2 = result;
                        rng.Value2 = outputRow.ToArray();
                        outputRow.Clear();
                    }
                    for (var i = 0; i < sources.Length; i++)
                    {
                        if (sources[i] == orderInputSheet.Name)
                        {
                            if (orderInputRow.ContainsKey(headers[i]))
                            {
                                outputRow.Add(orderInputRow[headers[i]]);
                                continue;
                            }
                        }
                        outputRow.Add("");
                    }
                    var newRow1 = outputSheet.UsedRange.Rows.Count + 1;
                    Range rng1 = outputSheet.Range[outputSheet.Cells[newRow1, 1], outputSheet.Cells[newRow1, width]];
                    rng1.Value2 = outputRow.ToArray();

                }
                //for (var i = 0; i < formats.Count; i++)
                //{
                //    if (string.Equals(formats[i], "TextToColumn", StringComparison.OrdinalIgnoreCase))
                //    {
                //        outputSheet.UsedRange.Columns[i + 1].TextToColumns();
                //    }
                //}

                for (var j = 5; j < outputSheet.UsedRange.Rows.Count + 1; j++)
                {
                    Range row = outputSheet.UsedRange.Rows[j];
                    for (var i = 1; i < row.Columns.Count; i++)
                    {
                        if (outputSheet.UsedRange.Rows[4].Columns[i].FormulaR1C1.ToString() != "")
                        {
                            //MessageBox.Show(outputSheet.UsedRange.Rows[3].Columns[i].FormulaR1C1.ToString());
                            row.Columns[i].FormulaR1C1 = outputSheet.UsedRange.Rows[4].Columns[i].FormulaR1C1;
                            row.Columns[i].Calculate();
                        }
                        //
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(Constants.PROTECTED_ERROR_MESSAGE);
            }
        }

        private bool CheckSheets()
        {
            for(var i = 0; i < checkedListBox1.Items.Count; i++)
            {
                object item = checkedListBox1.Items[i];
                Worksheet ws = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == item.ToString());
                if (ws == null)
                {
                    checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
                    MessageBox.Show("Worksheet "+ item.ToString()+ " is missing.");
                    return false;
                }
                else
                {
                    checkedListBox1.SetItemChecked(i, true);
                }
            }
            return true;
        }


        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void NaviToSheet(string sheetName)
        {
            try
            {
                Worksheet theSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[sheetName];
                if (theSheet != null)
                {
                    theSheet.Select();
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(Constants.PROTECTED_ERROR_MESSAGE);
            }
        }
        private void NaviToAggregationButtion_Click(object sender, EventArgs e)
        {
            NaviToSheet("汇总结果");
        }
        private void ShowAggregationAsSalesButton_Click(object sender, EventArgs e)
        {
            NaviToSheet("汇总结果");
        }
        private void ShowAggregationAsBuysideButton_Click(object sender, EventArgs e)
        {
            NaviToSheet("汇总结果");
        }

        private void NaviToOrderInputAsSalesButton_Click(object sender, EventArgs e)
        {
            NaviToSheet("订单输入");
        }
        private void NaviToOrderInputAsBuysideButton_Click(object sender, EventArgs e)
        {
            NaviToSheet("订单输入");
        }
        private void NaviToReceiptAsSalesButton_Click(object sender, EventArgs e)
        {
            NaviToSheet("收款输入");
        }
        private void NaviToReceiptAsBuysideButton_Click(object sender, EventArgs e)
        {
            NaviToSheet("收款输入");
        }

        private void unhideRange_Click(object sender, EventArgs e)
        {
            Worksheet theSheet = Globals.ThisAddIn.Application.ActiveSheet;
            if (theSheet != null)
            {
                theSheet.UsedRange.Columns.Hidden = false;
            }
        }
    }
}
