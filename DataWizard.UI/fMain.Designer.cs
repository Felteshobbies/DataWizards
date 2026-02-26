namespace DataWizard.UI
{
    partial class fMain
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.mainTabControl = new System.Windows.Forms.TabControl();
            this.convertTab = new System.Windows.Forms.TabPage();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.convertButton = new System.Windows.Forms.Button();
            this.dragDropPanel = new System.Windows.Forms.Panel();
            this.settingsTab = new System.Windows.Forms.TabPage();
            this.button1 = new System.Windows.Forms.Button();
            this.tbOutputfolder = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cbOverwrite = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbAllSheets = new System.Windows.Forms.CheckBox();
            this.cbQuoteAllText = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cbXlsSeparator = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cbXlsEncoding = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cbHeader = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cbCsvSeparator = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cbCsvEncoding = new System.Windows.Forms.ComboBox();
            this.watchTab = new System.Windows.Forms.TabPage();
            this.logTab = new System.Windows.Forms.TabPage();
            this.buttonPanel = new System.Windows.Forms.Panel();
            this.btClearLog = new System.Windows.Forms.Button();
            this.logTextBox = new System.Windows.Forms.RichTextBox();
            this.mainTabControl.SuspendLayout();
            this.convertTab.SuspendLayout();
            this.settingsTab.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.logTab.SuspendLayout();
            this.buttonPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainTabControl
            // 
            this.mainTabControl.Controls.Add(this.convertTab);
            this.mainTabControl.Controls.Add(this.settingsTab);
            this.mainTabControl.Controls.Add(this.watchTab);
            this.mainTabControl.Controls.Add(this.logTab);
            this.mainTabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainTabControl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mainTabControl.Location = new System.Drawing.Point(0, 0);
            this.mainTabControl.Name = "mainTabControl";
            this.mainTabControl.SelectedIndex = 0;
            this.mainTabControl.Size = new System.Drawing.Size(784, 311);
            this.mainTabControl.TabIndex = 0;
            // 
            // convertTab
            // 
            this.convertTab.BackColor = System.Drawing.Color.White;
            this.convertTab.Controls.Add(this.progressBar);
            this.convertTab.Controls.Add(this.convertButton);
            this.convertTab.Controls.Add(this.dragDropPanel);
            this.convertTab.Location = new System.Drawing.Point(4, 29);
            this.convertTab.Name = "convertTab";
            this.convertTab.Padding = new System.Windows.Forms.Padding(3);
            this.convertTab.Size = new System.Drawing.Size(776, 278);
            this.convertTab.TabIndex = 0;
            this.convertTab.Text = "tabPage1";
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.BackColor = System.Drawing.Color.White;
            this.progressBar.Location = new System.Drawing.Point(8, 198);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(760, 23);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 2;
            // 
            // convertButton
            // 
            this.convertButton.BackColor = System.Drawing.Color.LightGreen;
            this.convertButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.convertButton.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.convertButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.convertButton.ForeColor = System.Drawing.Color.Black;
            this.convertButton.Location = new System.Drawing.Point(3, 225);
            this.convertButton.Name = "convertButton";
            this.convertButton.Size = new System.Drawing.Size(770, 50);
            this.convertButton.TabIndex = 1;
            this.convertButton.Text = "Convert";
            this.convertButton.UseVisualStyleBackColor = false;
            this.convertButton.Click += new System.EventHandler(this.convertButton_Click);
            // 
            // dragDropPanel
            // 
            this.dragDropPanel.AllowDrop = true;
            this.dragDropPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dragDropPanel.BackColor = System.Drawing.Color.WhiteSmoke;
            this.dragDropPanel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.dragDropPanel.Location = new System.Drawing.Point(8, 6);
            this.dragDropPanel.Name = "dragDropPanel";
            this.dragDropPanel.Size = new System.Drawing.Size(760, 186);
            this.dragDropPanel.TabIndex = 0;
            this.dragDropPanel.Click += new System.EventHandler(this.dragDropPanel_Click);
            // 
            // settingsTab
            // 
            this.settingsTab.Controls.Add(this.button1);
            this.settingsTab.Controls.Add(this.tbOutputfolder);
            this.settingsTab.Controls.Add(this.label5);
            this.settingsTab.Controls.Add(this.cbOverwrite);
            this.settingsTab.Controls.Add(this.groupBox2);
            this.settingsTab.Controls.Add(this.groupBox1);
            this.settingsTab.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.settingsTab.Location = new System.Drawing.Point(4, 29);
            this.settingsTab.Name = "settingsTab";
            this.settingsTab.Padding = new System.Windows.Forms.Padding(3);
            this.settingsTab.Size = new System.Drawing.Size(776, 278);
            this.settingsTab.TabIndex = 1;
            this.settingsTab.Text = "tabPage2";
            this.settingsTab.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(201, 44);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(47, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "Pick";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // tbOutputfolder
            // 
            this.tbOutputfolder.Location = new System.Drawing.Point(254, 44);
            this.tbOutputfolder.Name = "tbOutputfolder";
            this.tbOutputfolder.Size = new System.Drawing.Size(511, 23);
            this.tbOutputfolder.TabIndex = 5;
            this.tbOutputfolder.TextChanged += new System.EventHandler(this.tbOutputfolder_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(198, 24);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(266, 17);
            this.label5.TabIndex = 4;
            this.label5.Text = "Outputfolder (Empty for sourcefile folder)";
            // 
            // cbOverwrite
            // 
            this.cbOverwrite.AutoSize = true;
            this.cbOverwrite.Location = new System.Drawing.Point(17, 24);
            this.cbOverwrite.Name = "cbOverwrite";
            this.cbOverwrite.Size = new System.Drawing.Size(116, 21);
            this.cbOverwrite.TabIndex = 2;
            this.cbOverwrite.Text = "Overwrite files";
            this.cbOverwrite.UseVisualStyleBackColor = true;
            this.cbOverwrite.CheckedChanged += new System.EventHandler(this.cbOverwrite_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cbAllSheets);
            this.groupBox2.Controls.Add(this.cbQuoteAllText);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.cbXlsSeparator);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.cbXlsEncoding);
            this.groupBox2.Location = new System.Drawing.Point(401, 101);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(370, 165);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Excel Convert Settings";
            // 
            // cbAllSheets
            // 
            this.cbAllSheets.AutoSize = true;
            this.cbAllSheets.Location = new System.Drawing.Point(136, 120);
            this.cbAllSheets.Name = "cbAllSheets";
            this.cbAllSheets.Size = new System.Drawing.Size(88, 21);
            this.cbAllSheets.TabIndex = 8;
            this.cbAllSheets.Text = "All sheets";
            this.cbAllSheets.UseVisualStyleBackColor = true;
            this.cbAllSheets.CheckedChanged += new System.EventHandler(this.cbAllSheets_CheckedChanged);
            // 
            // cbQuoteAllText
            // 
            this.cbQuoteAllText.AutoSize = true;
            this.cbQuoteAllText.CheckState = System.Windows.Forms.CheckState.Indeterminate;
            this.cbQuoteAllText.ThreeState = true;
            this.cbQuoteAllText.Location = new System.Drawing.Point(136, 93);
            this.cbQuoteAllText.Name = "cbQuoteAllText";
            this.cbQuoteAllText.Size = new System.Drawing.Size(140, 21);
            this.cbQuoteAllText.TabIndex = 7;
            this.cbQuoteAllText.Text = "Force text quote";
            this.cbQuoteAllText.UseVisualStyleBackColor = true;
            this.cbQuoteAllText.CheckedChanged += new System.EventHandler(this.cbQuoteAllText_CheckedChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 63);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 17);
            this.label4.TabIndex = 5;
            this.label4.Text = "Separator";
            // 
            // cbXlsSeparator
            // 
            this.cbXlsSeparator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbXlsSeparator.FormattingEnabled = true;
            this.cbXlsSeparator.Location = new System.Drawing.Point(136, 60);
            this.cbXlsSeparator.Name = "cbXlsSeparator";
            this.cbXlsSeparator.Size = new System.Drawing.Size(228, 24);
            this.cbXlsSeparator.TabIndex = 4;
            this.cbXlsSeparator.SelectedIndexChanged += new System.EventHandler(this.cbXlsSeparator_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 17);
            this.label2.TabIndex = 3;
            this.label2.Text = "Encoding";
            // 
            // cbXlsEncoding
            // 
            this.cbXlsEncoding.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbXlsEncoding.FormattingEnabled = true;
            this.cbXlsEncoding.Location = new System.Drawing.Point(136, 30);
            this.cbXlsEncoding.Name = "cbXlsEncoding";
            this.cbXlsEncoding.Size = new System.Drawing.Size(228, 24);
            this.cbXlsEncoding.TabIndex = 2;
            this.cbXlsEncoding.SelectedIndexChanged += new System.EventHandler(this.cbXlsEncoding_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.cbHeader);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.cbCsvSeparator);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cbCsvEncoding);
            this.groupBox1.Location = new System.Drawing.Point(8, 101);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(370, 169);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "CSV Convert Settings";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 93);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(55, 17);
            this.label6.TabIndex = 5;
            this.label6.Text = "Header";
            // 
            // cbHeader
            // 
            this.cbHeader.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbHeader.FormattingEnabled = true;
            this.cbHeader.Location = new System.Drawing.Point(136, 90);
            this.cbHeader.Name = "cbHeader";
            this.cbHeader.Size = new System.Drawing.Size(228, 24);
            this.cbHeader.TabIndex = 4;
            this.cbHeader.SelectedIndexChanged += new System.EventHandler(this.cbHeader_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 63);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 17);
            this.label3.TabIndex = 3;
            this.label3.Text = "Separator";
            // 
            // cbCsvSeparator
            // 
            this.cbCsvSeparator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbCsvSeparator.FormattingEnabled = true;
            this.cbCsvSeparator.Location = new System.Drawing.Point(136, 60);
            this.cbCsvSeparator.Name = "cbCsvSeparator";
            this.cbCsvSeparator.Size = new System.Drawing.Size(228, 24);
            this.cbCsvSeparator.TabIndex = 2;
            this.cbCsvSeparator.SelectedIndexChanged += new System.EventHandler(this.cbCsvSeparator_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Encoding";
            // 
            // cbCsvEncoding
            // 
            this.cbCsvEncoding.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbCsvEncoding.FormattingEnabled = true;
            this.cbCsvEncoding.Location = new System.Drawing.Point(136, 30);
            this.cbCsvEncoding.Name = "cbCsvEncoding";
            this.cbCsvEncoding.Size = new System.Drawing.Size(228, 24);
            this.cbCsvEncoding.TabIndex = 0;
            this.cbCsvEncoding.SelectedIndexChanged += new System.EventHandler(this.cbCsvEncoding_SelectedIndexChanged);
            // 
            // watchTab
            // 
            this.watchTab.Location = new System.Drawing.Point(4, 29);
            this.watchTab.Name = "watchTab";
            this.watchTab.Padding = new System.Windows.Forms.Padding(3);
            this.watchTab.Size = new System.Drawing.Size(776, 278);
            this.watchTab.TabIndex = 2;
            this.watchTab.Text = "tabPage3";
            this.watchTab.UseVisualStyleBackColor = true;
            // 
            // logTab
            // 
            this.logTab.Controls.Add(this.buttonPanel);
            this.logTab.Controls.Add(this.logTextBox);
            this.logTab.Location = new System.Drawing.Point(4, 29);
            this.logTab.Name = "logTab";
            this.logTab.Padding = new System.Windows.Forms.Padding(3);
            this.logTab.Size = new System.Drawing.Size(776, 278);
            this.logTab.TabIndex = 3;
            this.logTab.Text = "tabPage4";
            this.logTab.UseVisualStyleBackColor = true;
            // 
            // buttonPanel
            // 
            this.buttonPanel.Controls.Add(this.btClearLog);
            this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.buttonPanel.Location = new System.Drawing.Point(3, 225);
            this.buttonPanel.Name = "buttonPanel";
            this.buttonPanel.Size = new System.Drawing.Size(770, 50);
            this.buttonPanel.TabIndex = 1;
            // 
            // btClearLog
            // 
            this.btClearLog.Location = new System.Drawing.Point(10, 10);
            this.btClearLog.Name = "btClearLog";
            this.btClearLog.Size = new System.Drawing.Size(120, 30);
            this.btClearLog.TabIndex = 0;
            this.btClearLog.Text = "Clear Log";
            this.btClearLog.UseVisualStyleBackColor = true;
            this.btClearLog.Click += new System.EventHandler(this.btClearLog_Click);
            // 
            // logTextBox
            // 
            this.logTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.logTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logTextBox.Location = new System.Drawing.Point(3, 3);
            this.logTextBox.Name = "logTextBox";
            this.logTextBox.Size = new System.Drawing.Size(770, 216);
            this.logTextBox.TabIndex = 0;
            this.logTextBox.Text = "";
            this.logTextBox.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.logTextBox_LinkClicked);
            // 
            // fMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 311);
            this.Controls.Add(this.mainTabControl);
            this.Name = "fMain";
            this.Text = "DataWizard: CSV <> XLSX";
            this.mainTabControl.ResumeLayout(false);
            this.convertTab.ResumeLayout(false);
            this.settingsTab.ResumeLayout(false);
            this.settingsTab.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.logTab.ResumeLayout(false);
            this.buttonPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl mainTabControl;
        private System.Windows.Forms.TabPage convertTab;
        private System.Windows.Forms.TabPage settingsTab;
        private System.Windows.Forms.TabPage watchTab;
        private System.Windows.Forms.TabPage logTab;
        private System.Windows.Forms.RichTextBox logTextBox;
        private System.Windows.Forms.Panel buttonPanel;
        private System.Windows.Forms.Button btClearLog;
        private System.Windows.Forms.Panel dragDropPanel;
        private System.Windows.Forms.Button convertButton;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox cbOverwrite;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbCsvEncoding;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbXlsEncoding;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbCsvSeparator;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbXlsSeparator;
        private System.Windows.Forms.TextBox tbOutputfolder;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cbHeader;
        private System.Windows.Forms.CheckBox cbAllSheets;
        private System.Windows.Forms.CheckBox cbQuoteAllText;
    }
}