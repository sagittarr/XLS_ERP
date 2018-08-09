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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TaskPaneControl));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.userLabel = new System.Windows.Forms.Label();
            this.logoutButton = new System.Windows.Forms.Button();
            this.loginbutton = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.usernameBox = new System.Windows.Forms.TextBox();
            this.passwordBox = new System.Windows.Forms.TextBox();
            this.managerTabPage = new System.Windows.Forms.TabPage();
            this.aggregationLabel = new System.Windows.Forms.Label();
            this.unhideRange = new System.Windows.Forms.Button();
            this.NaviToAggregationButtion = new System.Windows.Forms.Button();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.aggregationButton = new System.Windows.Forms.Button();
            this.ManageButton = new System.Windows.Forms.Button();
            this.showUserButton = new System.Windows.Forms.Button();
            this.buysideTabPage = new System.Windows.Forms.TabPage();
            this.ShowAggregationAsBuysideButton = new System.Windows.Forms.Button();
            this.NaviToReceiptAsBuysideButton = new System.Windows.Forms.Button();
            this.NaviToOrderInputAsBuysideButton = new System.Windows.Forms.Button();
            this.sellsideTabPage = new System.Windows.Forms.TabPage();
            this.ShowAggregationAsSalesButton = new System.Windows.Forms.Button();
            this.NaviToReceiptAsSalesButton = new System.Windows.Forms.Button();
            this.NaviToOrderInputAsSalesButton = new System.Windows.Forms.Button();
            this.helpTab = new System.Windows.Forms.TabPage();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.managerTabPage.SuspendLayout();
            this.buysideTabPage.SuspendLayout();
            this.sellsideTabPage.SuspendLayout();
            this.helpTab.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.managerTabPage);
            this.tabControl1.Controls.Add(this.buysideTabPage);
            this.tabControl1.Controls.Add(this.sellsideTabPage);
            this.tabControl1.Controls.Add(this.helpTab);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(300, 400);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.pictureBox1);
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
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox1.Location = new System.Drawing.Point(85, 20);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(127, 81);
            this.pictureBox1.TabIndex = 13;
            this.pictureBox1.TabStop = false;
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
            this.logoutButton.Location = new System.Drawing.Point(85, 291);
            this.logoutButton.Margin = new System.Windows.Forms.Padding(2);
            this.logoutButton.Name = "logoutButton";
            this.logoutButton.Size = new System.Drawing.Size(127, 35);
            this.logoutButton.TabIndex = 11;
            this.logoutButton.Text = "退出登录";
            this.logoutButton.UseVisualStyleBackColor = true;
            this.logoutButton.Click += new System.EventHandler(this.logoutButton_Click);
            // 
            // loginbutton
            // 
            this.loginbutton.Location = new System.Drawing.Point(85, 244);
            this.loginbutton.Margin = new System.Windows.Forms.Padding(2);
            this.loginbutton.Name = "loginbutton";
            this.loginbutton.Size = new System.Drawing.Size(127, 34);
            this.loginbutton.TabIndex = 8;
            this.loginbutton.Text = "登录";
            this.loginbutton.UseVisualStyleBackColor = true;
            this.loginbutton.Click += new System.EventHandler(this.loginbutton_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(94, 121);
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
            this.label2.Location = new System.Drawing.Point(27, 201);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "密码";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 162);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "用户名";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // usernameBox
            // 
            this.usernameBox.Location = new System.Drawing.Point(85, 159);
            this.usernameBox.Margin = new System.Windows.Forms.Padding(2);
            this.usernameBox.Name = "usernameBox";
            this.usernameBox.Size = new System.Drawing.Size(127, 20);
            this.usernameBox.TabIndex = 4;
            this.usernameBox.TextChanged += new System.EventHandler(this.usernameBox_TextChanged);
            // 
            // passwordBox
            // 
            this.passwordBox.Location = new System.Drawing.Point(85, 201);
            this.passwordBox.Margin = new System.Windows.Forms.Padding(2);
            this.passwordBox.Name = "passwordBox";
            this.passwordBox.Size = new System.Drawing.Size(127, 20);
            this.passwordBox.TabIndex = 3;
            this.passwordBox.TextChanged += new System.EventHandler(this.passwordBox_TextChanged);
            // 
            // managerTabPage
            // 
            this.managerTabPage.Controls.Add(this.aggregationLabel);
            this.managerTabPage.Controls.Add(this.unhideRange);
            this.managerTabPage.Controls.Add(this.NaviToAggregationButtion);
            this.managerTabPage.Controls.Add(this.checkedListBox1);
            this.managerTabPage.Controls.Add(this.aggregationButton);
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
            this.managerTabPage.Click += new System.EventHandler(this.managerTabPage_Click);
            // 
            // aggregationLabel
            // 
            this.aggregationLabel.AutoSize = true;
            this.aggregationLabel.Location = new System.Drawing.Point(116, 171);
            this.aggregationLabel.Name = "aggregationLabel";
            this.aggregationLabel.Size = new System.Drawing.Size(55, 13);
            this.aggregationLabel.TabIndex = 20;
            this.aggregationLabel.Text = "汇总表单";
            // 
            // unhideRange
            // 
            this.unhideRange.Location = new System.Drawing.Point(83, 297);
            this.unhideRange.Name = "unhideRange";
            this.unhideRange.Size = new System.Drawing.Size(119, 34);
            this.unhideRange.TabIndex = 19;
            this.unhideRange.Text = "显示所有行列";
            this.unhideRange.UseVisualStyleBackColor = true;
            this.unhideRange.Click += new System.EventHandler(this.unhideRange_Click);
            // 
            // NaviToAggregationButtion
            // 
            this.NaviToAggregationButtion.Location = new System.Drawing.Point(83, 119);
            this.NaviToAggregationButtion.Name = "NaviToAggregationButtion";
            this.NaviToAggregationButtion.Size = new System.Drawing.Size(119, 34);
            this.NaviToAggregationButtion.TabIndex = 18;
            this.NaviToAggregationButtion.Text = "查看汇总结果";
            this.NaviToAggregationButtion.UseVisualStyleBackColor = true;
            this.NaviToAggregationButtion.Click += new System.EventHandler(this.NaviToAggregationButtion_Click);
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.Cursor = System.Windows.Forms.Cursors.Default;
            this.checkedListBox1.Enabled = false;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Items.AddRange(new object[] {
            "订单输入",
            "收款输入",
            "汇总结果"});
            this.checkedListBox1.Location = new System.Drawing.Point(83, 187);
            this.checkedListBox1.MultiColumn = true;
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(119, 64);
            this.checkedListBox1.TabIndex = 17;
            // 
            // aggregationButton
            // 
            this.aggregationButton.Location = new System.Drawing.Point(83, 257);
            this.aggregationButton.Name = "aggregationButton";
            this.aggregationButton.Size = new System.Drawing.Size(119, 34);
            this.aggregationButton.TabIndex = 15;
            this.aggregationButton.Text = "一键汇总";
            this.aggregationButton.UseVisualStyleBackColor = true;
            this.aggregationButton.Click += new System.EventHandler(this.aggregationButton_Click);
            // 
            // ManageButton
            // 
            this.ManageButton.Location = new System.Drawing.Point(83, 79);
            this.ManageButton.Name = "ManageButton";
            this.ManageButton.Size = new System.Drawing.Size(119, 34);
            this.ManageButton.TabIndex = 12;
            this.ManageButton.Text = "Sheet权限管理";
            this.ManageButton.UseVisualStyleBackColor = true;
            this.ManageButton.Click += new System.EventHandler(this.ManageButton_Click);
            // 
            // showUserButton
            // 
            this.showUserButton.Location = new System.Drawing.Point(83, 41);
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
            this.buysideTabPage.Controls.Add(this.ShowAggregationAsBuysideButton);
            this.buysideTabPage.Controls.Add(this.NaviToReceiptAsBuysideButton);
            this.buysideTabPage.Controls.Add(this.NaviToOrderInputAsBuysideButton);
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
            // ShowAggregationAsBuysideButton
            // 
            this.ShowAggregationAsBuysideButton.Location = new System.Drawing.Point(87, 177);
            this.ShowAggregationAsBuysideButton.Name = "ShowAggregationAsBuysideButton";
            this.ShowAggregationAsBuysideButton.Size = new System.Drawing.Size(119, 34);
            this.ShowAggregationAsBuysideButton.TabIndex = 19;
            this.ShowAggregationAsBuysideButton.Text = "查看汇总结果";
            this.ShowAggregationAsBuysideButton.UseVisualStyleBackColor = true;
            this.ShowAggregationAsBuysideButton.Click += new System.EventHandler(this.ShowAggregationAsBuysideButton_Click);
            // 
            // NaviToReceiptAsBuysideButton
            // 
            this.NaviToReceiptAsBuysideButton.Location = new System.Drawing.Point(87, 227);
            this.NaviToReceiptAsBuysideButton.Margin = new System.Windows.Forms.Padding(2);
            this.NaviToReceiptAsBuysideButton.Name = "NaviToReceiptAsBuysideButton";
            this.NaviToReceiptAsBuysideButton.Size = new System.Drawing.Size(119, 33);
            this.NaviToReceiptAsBuysideButton.TabIndex = 13;
            this.NaviToReceiptAsBuysideButton.Text = "收款记录";
            this.NaviToReceiptAsBuysideButton.UseVisualStyleBackColor = true;
            this.NaviToReceiptAsBuysideButton.Click += new System.EventHandler(this.NaviToReceiptAsBuysideButton_Click);
            // 
            // NaviToOrderInputAsBuysideButton
            // 
            this.NaviToOrderInputAsBuysideButton.Location = new System.Drawing.Point(87, 127);
            this.NaviToOrderInputAsBuysideButton.Margin = new System.Windows.Forms.Padding(2);
            this.NaviToOrderInputAsBuysideButton.Name = "NaviToOrderInputAsBuysideButton";
            this.NaviToOrderInputAsBuysideButton.Size = new System.Drawing.Size(119, 33);
            this.NaviToOrderInputAsBuysideButton.TabIndex = 12;
            this.NaviToOrderInputAsBuysideButton.Text = "销售订单";
            this.NaviToOrderInputAsBuysideButton.UseVisualStyleBackColor = true;
            this.NaviToOrderInputAsBuysideButton.Click += new System.EventHandler(this.NaviToOrderInputAsBuysideButton_Click);
            // 
            // sellsideTabPage
            // 
            this.sellsideTabPage.Controls.Add(this.ShowAggregationAsSalesButton);
            this.sellsideTabPage.Controls.Add(this.NaviToReceiptAsSalesButton);
            this.sellsideTabPage.Controls.Add(this.NaviToOrderInputAsSalesButton);
            this.sellsideTabPage.Location = new System.Drawing.Point(4, 22);
            this.sellsideTabPage.Margin = new System.Windows.Forms.Padding(2);
            this.sellsideTabPage.Name = "sellsideTabPage";
            this.sellsideTabPage.Padding = new System.Windows.Forms.Padding(2);
            this.sellsideTabPage.Size = new System.Drawing.Size(292, 374);
            this.sellsideTabPage.TabIndex = 3;
            this.sellsideTabPage.Text = "销售模块";
            this.sellsideTabPage.UseVisualStyleBackColor = true;
            // 
            // ShowAggregationAsSalesButton
            // 
            this.ShowAggregationAsSalesButton.Location = new System.Drawing.Point(87, 170);
            this.ShowAggregationAsSalesButton.Name = "ShowAggregationAsSalesButton";
            this.ShowAggregationAsSalesButton.Size = new System.Drawing.Size(119, 34);
            this.ShowAggregationAsSalesButton.TabIndex = 19;
            this.ShowAggregationAsSalesButton.Text = "查看汇总结果";
            this.ShowAggregationAsSalesButton.UseVisualStyleBackColor = true;
            this.ShowAggregationAsSalesButton.Click += new System.EventHandler(this.ShowAggregationAsSalesButton_Click);
            // 
            // NaviToReceiptAsSalesButton
            // 
            this.NaviToReceiptAsSalesButton.Location = new System.Drawing.Point(87, 212);
            this.NaviToReceiptAsSalesButton.Margin = new System.Windows.Forms.Padding(2);
            this.NaviToReceiptAsSalesButton.Name = "NaviToReceiptAsSalesButton";
            this.NaviToReceiptAsSalesButton.Size = new System.Drawing.Size(119, 33);
            this.NaviToReceiptAsSalesButton.TabIndex = 12;
            this.NaviToReceiptAsSalesButton.Text = "收款记录";
            this.NaviToReceiptAsSalesButton.UseVisualStyleBackColor = true;
            this.NaviToReceiptAsSalesButton.Click += new System.EventHandler(this.NaviToReceiptAsSalesButton_Click);
            // 
            // NaviToOrderInputAsSalesButton
            // 
            this.NaviToOrderInputAsSalesButton.Location = new System.Drawing.Point(87, 126);
            this.NaviToOrderInputAsSalesButton.Margin = new System.Windows.Forms.Padding(2);
            this.NaviToOrderInputAsSalesButton.Name = "NaviToOrderInputAsSalesButton";
            this.NaviToOrderInputAsSalesButton.Size = new System.Drawing.Size(119, 33);
            this.NaviToOrderInputAsSalesButton.TabIndex = 11;
            this.NaviToOrderInputAsSalesButton.Text = "销售订单";
            this.NaviToOrderInputAsSalesButton.UseVisualStyleBackColor = true;
            this.NaviToOrderInputAsSalesButton.Click += new System.EventHandler(this.NaviToOrderInputAsSalesButton_Click);
            // 
            // helpTab
            // 
            this.helpTab.Controls.Add(this.richTextBox1);
            this.helpTab.Location = new System.Drawing.Point(4, 22);
            this.helpTab.Name = "helpTab";
            this.helpTab.Padding = new System.Windows.Forms.Padding(3);
            this.helpTab.Size = new System.Drawing.Size(292, 374);
            this.helpTab.TabIndex = 4;
            this.helpTab.Text = "说明";
            this.helpTab.UseVisualStyleBackColor = true;
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(6, 6);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(280, 362);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "\n说明文档\n\n1.用户名与密码由管理员提供\n2.管理员请勿使用以下保留名作为新Sheet的名字\n\n";
            this.richTextBox1.TextChanged += new System.EventHandler(this.richTextBox1_TextChanged);
            // 
            // toolTip1
            // 
            this.toolTip1.Popup += new System.Windows.Forms.PopupEventHandler(this.toolTip1_Popup_1);
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
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.managerTabPage.ResumeLayout(false);
            this.managerTabPage.PerformLayout();
            this.buysideTabPage.ResumeLayout(false);
            this.sellsideTabPage.ResumeLayout(false);
            this.helpTab.ResumeLayout(false);
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
        private System.Windows.Forms.Button loginbutton;
        private System.Windows.Forms.Button ManageButton;
        private System.Windows.Forms.Label userLabel;
        private System.Windows.Forms.Button logoutButton;
        private System.Windows.Forms.Button aggregationButton;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label aggregationLabel;
        private System.Windows.Forms.Button unhideRange;
        private System.Windows.Forms.Button NaviToAggregationButtion;
        private System.Windows.Forms.Button NaviToReceiptAsBuysideButton;
        private System.Windows.Forms.Button NaviToOrderInputAsBuysideButton;
        private System.Windows.Forms.Button NaviToReceiptAsSalesButton;
        private System.Windows.Forms.Button NaviToOrderInputAsSalesButton;
        private System.Windows.Forms.Button ShowAggregationAsBuysideButton;
        private System.Windows.Forms.Button ShowAggregationAsSalesButton;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TabPage helpTab;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}
