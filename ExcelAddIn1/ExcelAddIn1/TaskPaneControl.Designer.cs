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
            this.showColumnPermissionButton = new System.Windows.Forms.Button();
            this.showColumnsButton = new System.Windows.Forms.Button();
            this.InputButton = new System.Windows.Forms.Button();
            this.aggregationButton = new System.Windows.Forms.Button();
            this.userLabel = new System.Windows.Forms.Label();
            this.logoutButton = new System.Windows.Forms.Button();
            this.loginbutton = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.usernameBox = new System.Windows.Forms.TextBox();
            this.passwordBox = new System.Windows.Forms.TextBox();
            this.managerTabPage = new System.Windows.Forms.TabPage();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.ManageButton = new System.Windows.Forms.Button();
            this.showUserButton = new System.Windows.Forms.Button();
            this.buysideTabPage = new System.Windows.Forms.TabPage();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.sellsideTabPage = new System.Windows.Forms.TabPage();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.eventLog1 = new System.Diagnostics.EventLog();
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
            this.tabControl1.Location = new System.Drawing.Point(0, 2);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(300, 400);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.showColumnPermissionButton);
            this.tabPage1.Controls.Add(this.showColumnsButton);
            this.tabPage1.Controls.Add(this.InputButton);
            this.tabPage1.Controls.Add(this.aggregationButton);
            this.tabPage1.Controls.Add(this.userLabel);
            this.tabPage1.Controls.Add(this.logoutButton);
            this.tabPage1.Controls.Add(this.loginbutton);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.usernameBox);
            this.tabPage1.Controls.Add(this.passwordBox);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage1.Size = new System.Drawing.Size(292, 374);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "身份选择";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // showColumnPermissionButton
            // 
            this.showColumnPermissionButton.Location = new System.Drawing.Point(86, 311);
            this.showColumnPermissionButton.Name = "showColumnPermissionButton";
            this.showColumnPermissionButton.Size = new System.Drawing.Size(126, 34);
            this.showColumnPermissionButton.TabIndex = 16;
            this.showColumnPermissionButton.Text = "列权限";
            this.showColumnPermissionButton.UseVisualStyleBackColor = true;
            this.showColumnPermissionButton.Click += new System.EventHandler(this.showColumnPermission_Click);
            // 
            // showColumnsButton
            // 
            this.showColumnsButton.Location = new System.Drawing.Point(86, 271);
            this.showColumnsButton.Name = "showColumnsButton";
            this.showColumnsButton.Size = new System.Drawing.Size(126, 34);
            this.showColumnsButton.TabIndex = 15;
            this.showColumnsButton.Text = "显示所有列";
            this.showColumnsButton.UseVisualStyleBackColor = true;
            this.showColumnsButton.Click += new System.EventHandler(this.showColumnsButton_Click);
            // 
            // InputButton
            // 
            this.InputButton.Location = new System.Drawing.Point(86, 230);
            this.InputButton.Name = "InputButton";
            this.InputButton.Size = new System.Drawing.Size(126, 34);
            this.InputButton.TabIndex = 14;
            this.InputButton.Text = "销售订单输入";
            this.InputButton.UseVisualStyleBackColor = true;
            this.InputButton.Click += new System.EventHandler(this.InputButton_Click);
            // 
            // aggregationButton
            // 
            this.aggregationButton.Location = new System.Drawing.Point(86, 189);
            this.aggregationButton.Name = "aggregationButton";
            this.aggregationButton.Size = new System.Drawing.Size(43, 35);
            this.aggregationButton.TabIndex = 13;
            this.aggregationButton.Text = "聚合";
            this.aggregationButton.UseVisualStyleBackColor = true;
            this.aggregationButton.Click += new System.EventHandler(this.button3_Click);
            // 
            // userLabel
            // 
            this.userLabel.AutoSize = true;
            this.userLabel.Location = new System.Drawing.Point(4, 2);
            this.userLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.userLabel.Name = "userLabel";
            this.userLabel.Size = new System.Drawing.Size(13, 13);
            this.userLabel.TabIndex = 12;
            this.userLabel.Text = "--";
            // 
            // logoutButton
            // 
            this.logoutButton.Location = new System.Drawing.Point(134, 189);
            this.logoutButton.Margin = new System.Windows.Forms.Padding(2);
            this.logoutButton.Name = "logoutButton";
            this.logoutButton.Size = new System.Drawing.Size(37, 35);
            this.logoutButton.TabIndex = 11;
            this.logoutButton.Text = "退出登录并保护工作簿";
            this.logoutButton.UseVisualStyleBackColor = true;
            this.logoutButton.Click += new System.EventHandler(this.logoutButton_Click);
            // 
            // loginbutton
            // 
            this.loginbutton.Location = new System.Drawing.Point(85, 150);
            this.loginbutton.Margin = new System.Windows.Forms.Padding(2);
            this.loginbutton.Name = "loginbutton";
            this.loginbutton.Size = new System.Drawing.Size(51, 34);
            this.loginbutton.TabIndex = 8;
            this.loginbutton.Text = "登录";
            this.loginbutton.UseVisualStyleBackColor = true;
            this.loginbutton.Click += new System.EventHandler(this.loginbutton_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(67, 45);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "ERP 系统用户登录";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 118);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "密码";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 87);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "用户名";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // usernameBox
            // 
            this.usernameBox.Location = new System.Drawing.Point(85, 84);
            this.usernameBox.Margin = new System.Windows.Forms.Padding(2);
            this.usernameBox.Name = "usernameBox";
            this.usernameBox.Size = new System.Drawing.Size(127, 20);
            this.usernameBox.TabIndex = 4;
            this.usernameBox.TextChanged += new System.EventHandler(this.usernameBox_TextChanged);
            // 
            // passwordBox
            // 
            this.passwordBox.Location = new System.Drawing.Point(85, 115);
            this.passwordBox.Margin = new System.Windows.Forms.Padding(2);
            this.passwordBox.Name = "passwordBox";
            this.passwordBox.Size = new System.Drawing.Size(127, 20);
            this.passwordBox.TabIndex = 3;
            this.passwordBox.TextChanged += new System.EventHandler(this.passwordBox_TextChanged);
            // 
            // managerTabPage
            // 
            this.managerTabPage.Controls.Add(this.button4);
            this.managerTabPage.Controls.Add(this.button3);
            this.managerTabPage.Controls.Add(this.ManageButton);
            this.managerTabPage.Controls.Add(this.showUserButton);
            this.managerTabPage.Location = new System.Drawing.Point(4, 22);
            this.managerTabPage.Margin = new System.Windows.Forms.Padding(2);
            this.managerTabPage.Name = "managerTabPage";
            this.managerTabPage.Padding = new System.Windows.Forms.Padding(2);
            this.managerTabPage.Size = new System.Drawing.Size(292, 374);
            this.managerTabPage.TabIndex = 1;
            this.managerTabPage.Text = "管理员控制台";
            this.managerTabPage.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(73, 204);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(119, 34);
            this.button4.TabIndex = 14;
            this.button4.Text = "单证输入权限";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(73, 164);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(119, 34);
            this.button3.TabIndex = 13;
            this.button3.Text = "销售输入权限";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // ManageButton
            // 
            this.ManageButton.Location = new System.Drawing.Point(73, 115);
            this.ManageButton.Name = "ManageButton";
            this.ManageButton.Size = new System.Drawing.Size(119, 34);
            this.ManageButton.TabIndex = 12;
            this.ManageButton.Text = "Sheet权限管理";
            this.ManageButton.UseVisualStyleBackColor = true;
            this.ManageButton.Click += new System.EventHandler(this.ManageButton_Click);
            // 
            // showUserButton
            // 
            this.showUserButton.Location = new System.Drawing.Point(73, 69);
            this.showUserButton.Margin = new System.Windows.Forms.Padding(2);
            this.showUserButton.Name = "showUserButton";
            this.showUserButton.Size = new System.Drawing.Size(119, 33);
            this.showUserButton.TabIndex = 10;
            this.showUserButton.Text = "密码管理";
            this.showUserButton.UseVisualStyleBackColor = true;
            this.showUserButton.Click += new System.EventHandler(this.showUserButton_Click);
            // 
            // buysideTabPage
            // 
            this.buysideTabPage.Controls.Add(this.button2);
            this.buysideTabPage.Controls.Add(this.button1);
            this.buysideTabPage.Controls.Add(this.listBox1);
            this.buysideTabPage.Location = new System.Drawing.Point(4, 22);
            this.buysideTabPage.Margin = new System.Windows.Forms.Padding(2);
            this.buysideTabPage.Name = "buysideTabPage";
            this.buysideTabPage.Padding = new System.Windows.Forms.Padding(2);
            this.buysideTabPage.Size = new System.Drawing.Size(292, 374);
            this.buysideTabPage.TabIndex = 2;
            this.buysideTabPage.Text = "采购模块";
            this.buysideTabPage.UseVisualStyleBackColor = true;
            this.buysideTabPage.Click += new System.EventHandler(this.tabPage3_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(33, 125);
            this.button2.Margin = new System.Windows.Forms.Padding(2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(165, 19);
            this.button2.TabIndex = 3;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(33, 90);
            this.button1.Margin = new System.Windows.Forms.Padding(2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(165, 19);
            this.button1.TabIndex = 2;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Items.AddRange(new object[] {
            "待办事项1",
            "待办事项2"});
            this.listBox1.Location = new System.Drawing.Point(33, 29);
            this.listBox1.Margin = new System.Windows.Forms.Padding(2);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(167, 43);
            this.listBox1.TabIndex = 1;
            // 
            // sellsideTabPage
            // 
            this.sellsideTabPage.Controls.Add(this.comboBox1);
            this.sellsideTabPage.Location = new System.Drawing.Point(4, 22);
            this.sellsideTabPage.Margin = new System.Windows.Forms.Padding(2);
            this.sellsideTabPage.Name = "sellsideTabPage";
            this.sellsideTabPage.Padding = new System.Windows.Forms.Padding(2);
            this.sellsideTabPage.Size = new System.Drawing.Size(292, 374);
            this.sellsideTabPage.TabIndex = 3;
            this.sellsideTabPage.Text = "销售模块";
            this.sellsideTabPage.UseVisualStyleBackColor = true;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "客户A",
            "客户B"});
            this.comboBox1.Location = new System.Drawing.Point(51, 40);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(2);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(139, 21);
            this.comboBox1.TabIndex = 0;
            // 
            // eventLog1
            // 
            this.eventLog1.SynchronizingObject = this;
            // 
            // TaskPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "TaskPaneControl";
            this.Size = new System.Drawing.Size(300, 400);
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
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox usernameBox;
        private System.Windows.Forms.TextBox passwordBox;
        private System.Windows.Forms.TabPage sellsideTabPage;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button showUserButton;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Diagnostics.EventLog eventLog1;
        private System.Windows.Forms.Button loginbutton;
        private System.Windows.Forms.Button ManageButton;
        private System.Windows.Forms.Label userLabel;
        private System.Windows.Forms.Button logoutButton;
        private System.Windows.Forms.Button aggregationButton;
        private System.Windows.Forms.Button InputButton;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button showColumnsButton;
        private System.Windows.Forms.Button showColumnPermissionButton;
    }
}
