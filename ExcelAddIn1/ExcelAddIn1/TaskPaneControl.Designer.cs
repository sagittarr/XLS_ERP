namespace ExcelAddIn1
{
    partial class TaskPaneControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.sellsidebutton = new System.Windows.Forms.Button();
            this.buysidebutton = new System.Windows.Forms.Button();
            this.managerButton = new System.Windows.Forms.Button();
            this.managerTabPage = new System.Windows.Forms.TabPage();
            this.showUserButton = new System.Windows.Forms.Button();
            this.unhidebutton = new System.Windows.Forms.Button();
            this.deephidebutton = new System.Windows.Forms.Button();
            this.buysideTabPage = new System.Windows.Forms.TabPage();
            this.sellsideTabPage = new System.Windows.Forms.TabPage();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.eventLog1 = new System.Diagnostics.EventLog();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.managerTabPage.SuspendLayout();
            this.buysideTabPage.SuspendLayout();
            this.sellsideTabPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.eventLog1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.managerTabPage);
            this.tabControl1.Controls.Add(this.buysideTabPage);
            this.tabControl1.Controls.Add(this.sellsideTabPage);
            this.tabControl1.Location = new System.Drawing.Point(0, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(346, 497);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.textBox2);
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.sellsidebutton);
            this.tabPage1.Controls.Add(this.buysidebutton);
            this.tabPage1.Controls.Add(this.managerButton);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(338, 464);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "身份选择";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(101, 69);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(142, 20);
            this.label3.TabIndex = 7;
            this.label3.Text = "ERP 系统用户登录";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 182);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 20);
            this.label2.TabIndex = 6;
            this.label2.Text = "密码";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 134);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 20);
            this.label1.TabIndex = 5;
            this.label1.Text = "用户名";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(86, 131);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(189, 26);
            this.textBox2.TabIndex = 4;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(86, 179);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(189, 26);
            this.textBox1.TabIndex = 3;
            // 
            // sellsidebutton
            // 
            this.sellsidebutton.Location = new System.Drawing.Point(86, 337);
            this.sellsidebutton.Name = "sellsidebutton";
            this.sellsidebutton.Size = new System.Drawing.Size(189, 37);
            this.sellsidebutton.TabIndex = 2;
            this.sellsidebutton.Text = "销售人员";
            this.sellsidebutton.UseVisualStyleBackColor = true;
            this.sellsidebutton.Click += new System.EventHandler(this.sellsidebutton_Click);
            // 
            // buysidebutton
            // 
            this.buysidebutton.Location = new System.Drawing.Point(86, 280);
            this.buysidebutton.Name = "buysidebutton";
            this.buysidebutton.Size = new System.Drawing.Size(189, 37);
            this.buysidebutton.TabIndex = 1;
            this.buysidebutton.Text = "采购人员";
            this.buysidebutton.UseVisualStyleBackColor = true;
            this.buysidebutton.Click += new System.EventHandler(this.buysidebutton_Click);
            // 
            // managerButton
            // 
            this.managerButton.Location = new System.Drawing.Point(86, 225);
            this.managerButton.Name = "managerButton";
            this.managerButton.Size = new System.Drawing.Size(189, 33);
            this.managerButton.TabIndex = 0;
            this.managerButton.Text = "管理员";
            this.managerButton.UseVisualStyleBackColor = true;
            this.managerButton.Click += new System.EventHandler(this.managerbutton_Click);
            // 
            // managerTabPage
            // 
            this.managerTabPage.Controls.Add(this.showUserButton);
            this.managerTabPage.Controls.Add(this.unhidebutton);
            this.managerTabPage.Controls.Add(this.deephidebutton);
            this.managerTabPage.Location = new System.Drawing.Point(4, 29);
            this.managerTabPage.Name = "managerTabPage";
            this.managerTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.managerTabPage.Size = new System.Drawing.Size(338, 464);
            this.managerTabPage.TabIndex = 1;
            this.managerTabPage.Text = "管理员控制台";
            this.managerTabPage.UseVisualStyleBackColor = true;
            // 
            // showUserButton
            // 
            this.showUserButton.Location = new System.Drawing.Point(110, 108);
            this.showUserButton.Name = "showUserButton";
            this.showUserButton.Size = new System.Drawing.Size(121, 31);
            this.showUserButton.TabIndex = 10;
            this.showUserButton.Text = "显示用户权限";
            this.showUserButton.UseVisualStyleBackColor = true;
            this.showUserButton.Click += new System.EventHandler(this.showUserButton_Click);
            // 
            // unhidebutton
            // 
            this.unhidebutton.Location = new System.Drawing.Point(110, 363);
            this.unhidebutton.Name = "unhidebutton";
            this.unhidebutton.Size = new System.Drawing.Size(121, 30);
            this.unhidebutton.TabIndex = 9;
            this.unhidebutton.Text = "显示所有表单";
            this.unhidebutton.UseVisualStyleBackColor = true;
            this.unhidebutton.Click += new System.EventHandler(this.Button2_Click_1);
            // 
            // deephidebutton
            // 
            this.deephidebutton.Location = new System.Drawing.Point(110, 299);
            this.deephidebutton.Name = "deephidebutton";
            this.deephidebutton.Size = new System.Drawing.Size(121, 30);
            this.deephidebutton.TabIndex = 8;
            this.deephidebutton.Text = "深度隐藏表单";
            this.deephidebutton.UseVisualStyleBackColor = true;
            this.deephidebutton.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // buysideTabPage
            // 
            this.buysideTabPage.Controls.Add(this.button2);
            this.buysideTabPage.Controls.Add(this.button1);
            this.buysideTabPage.Controls.Add(this.listBox1);
            this.buysideTabPage.Location = new System.Drawing.Point(4, 29);
            this.buysideTabPage.Name = "buysideTabPage";
            this.buysideTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.buysideTabPage.Size = new System.Drawing.Size(338, 464);
            this.buysideTabPage.TabIndex = 2;
            this.buysideTabPage.Text = "采购模块";
            this.buysideTabPage.UseVisualStyleBackColor = true;
            this.buysideTabPage.Click += new System.EventHandler(this.tabPage3_Click);
            // 
            // sellsideTabPage
            // 
            this.sellsideTabPage.Controls.Add(this.comboBox1);
            this.sellsideTabPage.Location = new System.Drawing.Point(4, 29);
            this.sellsideTabPage.Name = "sellsideTabPage";
            this.sellsideTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.sellsideTabPage.Size = new System.Drawing.Size(338, 464);
            this.sellsideTabPage.TabIndex = 3;
            this.sellsideTabPage.Text = "销售模块";
            this.sellsideTabPage.UseVisualStyleBackColor = true;
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 20;
            this.listBox1.Items.AddRange(new object[] {
            "待办事项1",
            "待办事项2"});
            this.listBox1.Location = new System.Drawing.Point(50, 44);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(248, 64);
            this.listBox1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(50, 139);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(248, 29);
            this.button1.TabIndex = 2;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(50, 193);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(248, 29);
            this.button2.TabIndex = 3;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // printDocument1
            // 
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
            // 
            // eventLog1
            // 
            this.eventLog1.SynchronizingObject = this;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "客户A",
            "客户B"});
            this.comboBox1.Location = new System.Drawing.Point(77, 61);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(206, 28);
            this.comboBox1.TabIndex = 0;
            // 
            // TaskPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Name = "TaskPaneControl";
            this.Size = new System.Drawing.Size(349, 503);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.managerTabPage.ResumeLayout(false);
            this.buysideTabPage.ResumeLayout(false);
            this.sellsideTabPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.eventLog1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage managerTabPage;
        private System.Windows.Forms.TabPage buysideTabPage;
        private System.Windows.Forms.Button unhidebutton;
        private System.Windows.Forms.Button deephidebutton;
        private System.Windows.Forms.Button buysidebutton;
        private System.Windows.Forms.Button managerButton;
        private System.Windows.Forms.Button sellsidebutton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TabPage sellsideTabPage;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button showUserButton;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Diagnostics.EventLog eventLog1;
    }
}
