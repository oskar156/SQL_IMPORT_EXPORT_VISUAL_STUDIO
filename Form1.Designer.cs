namespace SQL_SERVER_IMPORT_EXPORT
{
    partial class Form1
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
            this.ImportPath = new System.Windows.Forms.Label();
            this.Extension = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.ImportPathTextBox = new System.Windows.Forms.TextBox();
            this.Server = new System.Windows.Forms.Label();
            this.Database = new System.Windows.Forms.Label();
            this.ExtensionListBoxImport = new System.Windows.Forms.ListBox();
            this.TablesToExportCommaList = new System.Windows.Forms.TextBox();
            this.TablesToExportRegex = new System.Windows.Forms.TextBox();
            this.ImportPathNotes = new System.Windows.Forms.Label();
            this.OutputTypeListBox = new System.Windows.Forms.ListBox();
            this.ExportPathTextBox = new System.Windows.Forms.TextBox();
            this.ExportPath = new System.Windows.Forms.Label();
            this.ExportPathNotes = new System.Windows.Forms.Label();
            this.IncludeHeadersCheckBox = new System.Windows.Forms.CheckBox();
            this.ImportDelimeterListBox = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.ExportDelimeterListBox = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.ExportQualifierListBox = new System.Windows.Forms.ListBox();
            this.label6 = new System.Windows.Forms.Label();
            this.ExportButton = new System.Windows.Forms.Button();
            this.ImportButton = new System.Windows.Forms.Button();
            this.ImportPanel = new System.Windows.Forms.Panel();
            this.ImportDelimeterLastSelected = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.ImportToSingleTableTextBox = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.FixedWidthColumnFilePathLabel = new System.Windows.Forms.Label();
            this.FixedWidthColumnFilePathTextBox = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.ExportPanel = new System.Windows.Forms.Panel();
            this.ExportDelimeterLastSelected = new System.Windows.Forms.Label();
            this.ExportQualifierLastSelected = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.FixedWidthColumnLengthMethodListBox = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.LoadSqlTables = new System.Windows.Forms.Button();
            this.TablesToExportListFromSql = new System.Windows.Forms.ListBox();
            this.label8 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.TablePickerRadioButton = new System.Windows.Forms.RadioButton();
            this.RegexPatternTableSearchRadioButton = new System.Windows.Forms.RadioButton();
            this.CommaSeperatedListTableSearchRadioButton = new System.Windows.Forms.RadioButton();
            this.label13 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.ServerComboBox = new System.Windows.Forms.ComboBox();
            this.DatabaseComboBox = new System.Windows.Forms.ComboBox();
            this.ImportPanel.SuspendLayout();
            this.ExportPanel.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ImportPath
            // 
            this.ImportPath.AutoSize = true;
            this.ImportPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ImportPath.Location = new System.Drawing.Point(15, 477);
            this.ImportPath.Name = "ImportPath";
            this.ImportPath.Size = new System.Drawing.Size(95, 16);
            this.ImportPath.TabIndex = 1;
            this.ImportPath.Text = "Import Path *";
            // 
            // Extension
            // 
            this.Extension.AutoSize = true;
            this.Extension.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Extension.Location = new System.Drawing.Point(10, 32);
            this.Extension.Name = "Extension";
            this.Extension.Size = new System.Drawing.Size(130, 16);
            this.Extension.TabIndex = 3;
            this.Extension.Text = "Import File Type *";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(8, 272);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(131, 16);
            this.label4.TabIndex = 5;
            this.label4.Text = "Export File Type *";
            // 
            // ImportPathTextBox
            // 
            this.ImportPathTextBox.Location = new System.Drawing.Point(9, 551);
            this.ImportPathTextBox.Name = "ImportPathTextBox";
            this.ImportPathTextBox.Size = new System.Drawing.Size(291, 20);
            this.ImportPathTextBox.TabIndex = 7;
            this.ImportPathTextBox.Text = "C:\\Users\\oscar\\Desktop\\e";
            // 
            // Server
            // 
            this.Server.AutoSize = true;
            this.Server.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Server.Location = new System.Drawing.Point(8, 8);
            this.Server.Name = "Server";
            this.Server.Size = new System.Drawing.Size(78, 20);
            this.Server.TabIndex = 10;
            this.Server.Text = "Server: *";
            // 
            // Database
            // 
            this.Database.AutoSize = true;
            this.Database.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Database.Location = new System.Drawing.Point(294, 8);
            this.Database.Name = "Database";
            this.Database.Size = new System.Drawing.Size(104, 20);
            this.Database.TabIndex = 11;
            this.Database.Text = "Database: *";
            // 
            // ExtensionListBoxImport
            // 
            this.ExtensionListBoxImport.FormattingEnabled = true;
            this.ExtensionListBoxImport.Items.AddRange(new object[] {
            "csv",
            "txt",
            "xls*"});
            this.ExtensionListBoxImport.Location = new System.Drawing.Point(13, 48);
            this.ExtensionListBoxImport.Name = "ExtensionListBoxImport";
            this.ExtensionListBoxImport.Size = new System.Drawing.Size(127, 69);
            this.ExtensionListBoxImport.TabIndex = 3;
            this.ExtensionListBoxImport.SelectedIndexChanged += new System.EventHandler(this.ExtensionListBoxImport_SelectedIndexChanged);
            // 
            // TablesToExportCommaList
            // 
            this.TablesToExportCommaList.Location = new System.Drawing.Point(151, 48);
            this.TablesToExportCommaList.Multiline = true;
            this.TablesToExportCommaList.Name = "TablesToExportCommaList";
            this.TablesToExportCommaList.Size = new System.Drawing.Size(302, 36);
            this.TablesToExportCommaList.TabIndex = 9;
            // 
            // TablesToExportRegex
            // 
            this.TablesToExportRegex.Location = new System.Drawing.Point(151, 91);
            this.TablesToExportRegex.Name = "TablesToExportRegex";
            this.TablesToExportRegex.Size = new System.Drawing.Size(302, 20);
            this.TablesToExportRegex.TabIndex = 10;
            // 
            // ImportPathNotes
            // 
            this.ImportPathNotes.AutoSize = true;
            this.ImportPathNotes.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.ImportPathNotes.Location = new System.Drawing.Point(11, 497);
            this.ImportPathNotes.Name = "ImportPathNotes";
            this.ImportPathNotes.Size = new System.Drawing.Size(339, 17);
            this.ImportPathNotes.TabIndex = 18;
            this.ImportPathNotes.Text = "  -Enter file path (w/extension) to import only that file.";
            // 
            // OutputTypeListBox
            // 
            this.OutputTypeListBox.FormattingEnabled = true;
            this.OutputTypeListBox.Items.AddRange(new object[] {
            "csv",
            "txt",
            "xlsx"});
            this.OutputTypeListBox.Location = new System.Drawing.Point(12, 298);
            this.OutputTypeListBox.Name = "OutputTypeListBox";
            this.OutputTypeListBox.Size = new System.Drawing.Size(127, 69);
            this.OutputTypeListBox.TabIndex = 12;
            this.OutputTypeListBox.SelectedIndexChanged += new System.EventHandler(this.OutputTypeListBox_SelectedIndexChanged);
            // 
            // ExportPathTextBox
            // 
            this.ExportPathTextBox.Location = new System.Drawing.Point(11, 551);
            this.ExportPathTextBox.Name = "ExportPathTextBox";
            this.ExportPathTextBox.Size = new System.Drawing.Size(291, 20);
            this.ExportPathTextBox.TabIndex = 17;
            this.ExportPathTextBox.Text = "C:\\Users\\oscar\\Desktop\\e";
            // 
            // ExportPath
            // 
            this.ExportPath.AutoSize = true;
            this.ExportPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ExportPath.Location = new System.Drawing.Point(8, 522);
            this.ExportPath.Name = "ExportPath";
            this.ExportPath.Size = new System.Drawing.Size(96, 16);
            this.ExportPath.TabIndex = 28;
            this.ExportPath.Text = "Export Path *";
            // 
            // ExportPathNotes
            // 
            this.ExportPathNotes.AutoSize = true;
            this.ExportPathNotes.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.ExportPathNotes.Location = new System.Drawing.Point(8, 535);
            this.ExportPathNotes.Name = "ExportPathNotes";
            this.ExportPathNotes.Size = new System.Drawing.Size(217, 17);
            this.ExportPathNotes.TabIndex = 30;
            this.ExportPathNotes.Text = "Enter a folder path (no file name)";
            // 
            // IncludeHeadersCheckBox
            // 
            this.IncludeHeadersCheckBox.AutoSize = true;
            this.IncludeHeadersCheckBox.Checked = true;
            this.IncludeHeadersCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.IncludeHeadersCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.IncludeHeadersCheckBox.Location = new System.Drawing.Point(9, 373);
            this.IncludeHeadersCheckBox.Name = "IncludeHeadersCheckBox";
            this.IncludeHeadersCheckBox.Size = new System.Drawing.Size(130, 21);
            this.IncludeHeadersCheckBox.TabIndex = 16;
            this.IncludeHeadersCheckBox.Text = "Include Headers";
            this.IncludeHeadersCheckBox.UseVisualStyleBackColor = true;
            // 
            // ImportDelimeterListBox
            // 
            this.ImportDelimeterListBox.FormattingEnabled = true;
            this.ImportDelimeterListBox.Items.AddRange(new object[] {
            "COMMA",
            "PIPE",
            "TAB",
            "FIXED WIDTH"});
            this.ImportDelimeterListBox.Location = new System.Drawing.Point(169, 48);
            this.ImportDelimeterListBox.Name = "ImportDelimeterListBox";
            this.ImportDelimeterListBox.Size = new System.Drawing.Size(127, 69);
            this.ImportDelimeterListBox.TabIndex = 4;
            this.ImportDelimeterListBox.SelectedIndexChanged += new System.EventHandler(this.ImportDelimeterListBox_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(166, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 16);
            this.label2.TabIndex = 34;
            this.label2.Text = "Import Delimeter *";
            // 
            // ExportDelimeterListBox
            // 
            this.ExportDelimeterListBox.FormattingEnabled = true;
            this.ExportDelimeterListBox.Items.AddRange(new object[] {
            "COMMA",
            "PIPE",
            "TAB",
            "FIXED WIDTH"});
            this.ExportDelimeterListBox.Location = new System.Drawing.Point(160, 298);
            this.ExportDelimeterListBox.Name = "ExportDelimeterListBox";
            this.ExportDelimeterListBox.Size = new System.Drawing.Size(127, 69);
            this.ExportDelimeterListBox.TabIndex = 13;
            this.ExportDelimeterListBox.SelectedIndexChanged += new System.EventHandler(this.ExportDelimeterListBox_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(157, 272);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(127, 16);
            this.label3.TabIndex = 36;
            this.label3.Text = "Export Delimiter *";
            // 
            // ExportQualifierListBox
            // 
            this.ExportQualifierListBox.FormattingEnabled = true;
            this.ExportQualifierListBox.Items.AddRange(new object[] {
            "\"",
            "\'",
            "<NO QUALIFIER>"});
            this.ExportQualifierListBox.Location = new System.Drawing.Point(299, 298);
            this.ExportQualifierListBox.Name = "ExportQualifierListBox";
            this.ExportQualifierListBox.Size = new System.Drawing.Size(127, 69);
            this.ExportQualifierListBox.TabIndex = 14;
            this.ExportQualifierListBox.SelectedIndexChanged += new System.EventHandler(this.ExportQualifierListBox_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.5F);
            this.label6.Location = new System.Drawing.Point(306, 272);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(97, 16);
            this.label6.TabIndex = 41;
            this.label6.Text = "Export Qualifier";
            // 
            // ExportButton
            // 
            this.ExportButton.Location = new System.Drawing.Point(11, 577);
            this.ExportButton.Name = "ExportButton";
            this.ExportButton.Size = new System.Drawing.Size(175, 42);
            this.ExportButton.TabIndex = 18;
            this.ExportButton.Text = "EXPORT";
            this.ExportButton.UseVisualStyleBackColor = true;
            this.ExportButton.Click += new System.EventHandler(this.ExportButton_Click);
            // 
            // ImportButton
            // 
            this.ImportButton.Location = new System.Drawing.Point(18, 577);
            this.ImportButton.Name = "ImportButton";
            this.ImportButton.Size = new System.Drawing.Size(175, 42);
            this.ImportButton.TabIndex = 8;
            this.ImportButton.Text = "IMPORT";
            this.ImportButton.UseVisualStyleBackColor = true;
            this.ImportButton.Click += new System.EventHandler(this.ImportButton_Click);
            // 
            // ImportPanel
            // 
            this.ImportPanel.Controls.Add(this.ImportDelimeterLastSelected);
            this.ImportPanel.Controls.Add(this.label10);
            this.ImportPanel.Controls.Add(this.ImportToSingleTableTextBox);
            this.ImportPanel.Controls.Add(this.label12);
            this.ImportPanel.Controls.Add(this.FixedWidthColumnFilePathLabel);
            this.ImportPanel.Controls.Add(this.FixedWidthColumnFilePathTextBox);
            this.ImportPanel.Controls.Add(this.label9);
            this.ImportPanel.Controls.Add(this.label7);
            this.ImportPanel.Controls.Add(this.ImportButton);
            this.ImportPanel.Controls.Add(this.ImportDelimeterListBox);
            this.ImportPanel.Controls.Add(this.ExtensionListBoxImport);
            this.ImportPanel.Controls.Add(this.label2);
            this.ImportPanel.Controls.Add(this.Extension);
            this.ImportPanel.Controls.Add(this.ImportPath);
            this.ImportPanel.Controls.Add(this.ImportPathNotes);
            this.ImportPanel.Controls.Add(this.ImportPathTextBox);
            this.ImportPanel.Location = new System.Drawing.Point(12, 35);
            this.ImportPanel.Name = "ImportPanel";
            this.ImportPanel.Size = new System.Drawing.Size(386, 622);
            this.ImportPanel.TabIndex = 50;
            // 
            // ImportDelimeterLastSelected
            // 
            this.ImportDelimeterLastSelected.AutoSize = true;
            this.ImportDelimeterLastSelected.Location = new System.Drawing.Point(166, 120);
            this.ImportDelimeterLastSelected.Name = "ImportDelimeterLastSelected";
            this.ImportDelimeterLastSelected.Size = new System.Drawing.Size(47, 13);
            this.ImportDelimeterLastSelected.TabIndex = 57;
            this.ImportDelimeterLastSelected.Text = "COMMA";
            this.ImportDelimeterLastSelected.Visible = false;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label10.Location = new System.Drawing.Point(6, 272);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(345, 34);
            this.label10.TabIndex = 56;
            this.label10.Text = "Enter a Table Name to Import all files to a single table\r\n(must have identical he" +
    "aders)";
            // 
            // ImportToSingleTableTextBox
            // 
            this.ImportToSingleTableTextBox.Location = new System.Drawing.Point(9, 309);
            this.ImportToSingleTableTextBox.Name = "ImportToSingleTableTextBox";
            this.ImportToSingleTableTextBox.Size = new System.Drawing.Size(287, 20);
            this.ImportToSingleTableTextBox.TabIndex = 5;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(11, 382);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(134, 52);
            this.label12.TabIndex = 53;
            this.label12.Text = "Format example:\r\nCOLUMN-NAME LENGTH\r\nFIRST 25\r\nCREDIT SCORE 3 etc...\r\n";
            // 
            // FixedWidthColumnFilePathLabel
            // 
            this.FixedWidthColumnFilePathLabel.AutoSize = true;
            this.FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.FixedWidthColumnFilePathLabel.Location = new System.Drawing.Point(6, 348);
            this.FixedWidthColumnFilePathLabel.Name = "FixedWidthColumnFilePathLabel";
            this.FixedWidthColumnFilePathLabel.Size = new System.Drawing.Size(309, 34);
            this.FixedWidthColumnFilePathLabel.TabIndex = 52;
            this.FixedWidthColumnFilePathLabel.Text = "Fixed Width Column File Path\r\n (^required with FIXED WIDTH Import Delimeter)";
            // 
            // FixedWidthColumnFilePathTextBox
            // 
            this.FixedWidthColumnFilePathTextBox.Location = new System.Drawing.Point(9, 437);
            this.FixedWidthColumnFilePathTextBox.Name = "FixedWidthColumnFilePathTextBox";
            this.FixedWidthColumnFilePathTextBox.Size = new System.Drawing.Size(291, 20);
            this.FixedWidthColumnFilePathTextBox.TabIndex = 6;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label9.Location = new System.Drawing.Point(11, 514);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(360, 34);
            this.label9.TabIndex = 50;
            this.label9.Text = "  -Enter folder path (no extension) to import every file in \r\n   that folder that" +
    " matches Import File Type.";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(4, 4);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 24);
            this.label7.TabIndex = 49;
            this.label7.Text = "Import";
            // 
            // ExportPanel
            // 
            this.ExportPanel.Controls.Add(this.ExportDelimeterLastSelected);
            this.ExportPanel.Controls.Add(this.ExportQualifierLastSelected);
            this.ExportPanel.Controls.Add(this.label5);
            this.ExportPanel.Controls.Add(this.FixedWidthColumnLengthMethodListBox);
            this.ExportPanel.Controls.Add(this.label1);
            this.ExportPanel.Controls.Add(this.label14);
            this.ExportPanel.Controls.Add(this.LoadSqlTables);
            this.ExportPanel.Controls.Add(this.TablesToExportListFromSql);
            this.ExportPanel.Controls.Add(this.label8);
            this.ExportPanel.Controls.Add(this.ExportButton);
            this.ExportPanel.Controls.Add(this.TablesToExportCommaList);
            this.ExportPanel.Controls.Add(this.ExportQualifierListBox);
            this.ExportPanel.Controls.Add(this.label4);
            this.ExportPanel.Controls.Add(this.label6);
            this.ExportPanel.Controls.Add(this.TablesToExportRegex);
            this.ExportPanel.Controls.Add(this.ExportDelimeterListBox);
            this.ExportPanel.Controls.Add(this.OutputTypeListBox);
            this.ExportPanel.Controls.Add(this.label3);
            this.ExportPanel.Controls.Add(this.ExportPath);
            this.ExportPanel.Controls.Add(this.ExportPathTextBox);
            this.ExportPanel.Controls.Add(this.IncludeHeadersCheckBox);
            this.ExportPanel.Controls.Add(this.ExportPathNotes);
            this.ExportPanel.Controls.Add(this.panel1);
            this.ExportPanel.Location = new System.Drawing.Point(404, 35);
            this.ExportPanel.Name = "ExportPanel";
            this.ExportPanel.Size = new System.Drawing.Size(472, 622);
            this.ExportPanel.TabIndex = 51;
            // 
            // ExportDelimeterLastSelected
            // 
            this.ExportDelimeterLastSelected.AutoSize = true;
            this.ExportDelimeterLastSelected.Location = new System.Drawing.Point(157, 369);
            this.ExportDelimeterLastSelected.Name = "ExportDelimeterLastSelected";
            this.ExportDelimeterLastSelected.Size = new System.Drawing.Size(47, 13);
            this.ExportDelimeterLastSelected.TabIndex = 60;
            this.ExportDelimeterLastSelected.Text = "COMMA";
            this.ExportDelimeterLastSelected.Visible = false;
            // 
            // ExportQualifierLastSelected
            // 
            this.ExportQualifierLastSelected.AutoSize = true;
            this.ExportQualifierLastSelected.Location = new System.Drawing.Point(296, 369);
            this.ExportQualifierLastSelected.Name = "ExportQualifierLastSelected";
            this.ExportQualifierLastSelected.Size = new System.Drawing.Size(12, 13);
            this.ExportQualifierLastSelected.TabIndex = 58;
            this.ExportQualifierLastSelected.Text = "\"";
            this.ExportQualifierLastSelected.Visible = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(190, 467);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(263, 52);
            this.label5.TabIndex = 58;
            this.label5.Text = "MAX LEN = Maximum Length of Values in Column\r\nCOL_LENGTH = Length of Column set i" +
    "n it\'s Definition\r\n\r\n(unfinished - COL_LENGTH is inconsistent)";
            this.label5.Visible = false;
            // 
            // FixedWidthColumnLengthMethodListBox
            // 
            this.FixedWidthColumnLengthMethodListBox.Enabled = false;
            this.FixedWidthColumnLengthMethodListBox.FormattingEnabled = true;
            this.FixedWidthColumnLengthMethodListBox.Items.AddRange(new object[] {
            "MAX LEN",
            "COL_LENGTH"});
            this.FixedWidthColumnLengthMethodListBox.Location = new System.Drawing.Point(11, 464);
            this.FixedWidthColumnLengthMethodListBox.Name = "FixedWidthColumnLengthMethodListBox";
            this.FixedWidthColumnLengthMethodListBox.Size = new System.Drawing.Size(175, 43);
            this.FixedWidthColumnLengthMethodListBox.TabIndex = 59;
            this.FixedWidthColumnLengthMethodListBox.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label1.Location = new System.Drawing.Point(8, 426);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(306, 34);
            this.label1.TabIndex = 58;
            this.label1.Text = "How to Define Fixed Width Column Lengths?\r\n(^required with FIXED WIDTH Export Del" +
    "imeter)";
            this.label1.Visible = false;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(3, 32);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(222, 16);
            this.label14.TabIndex = 54;
            this.label14.Text = "Choose Table Export Method: *";
            // 
            // LoadSqlTables
            // 
            this.LoadSqlTables.Location = new System.Drawing.Point(151, 117);
            this.LoadSqlTables.Name = "LoadSqlTables";
            this.LoadSqlTables.Size = new System.Drawing.Size(152, 19);
            this.LoadSqlTables.TabIndex = 53;
            this.LoadSqlTables.Text = "LOAD TABLES BELOW";
            this.LoadSqlTables.UseVisualStyleBackColor = true;
            this.LoadSqlTables.Click += new System.EventHandler(this.LoadSqlTables_Click);
            // 
            // TablesToExportListFromSql
            // 
            this.TablesToExportListFromSql.FormattingEnabled = true;
            this.TablesToExportListFromSql.Location = new System.Drawing.Point(15, 142);
            this.TablesToExportListFromSql.Name = "TablesToExportListFromSql";
            this.TablesToExportListFromSql.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.TablesToExportListFromSql.Size = new System.Drawing.Size(454, 121);
            this.TablesToExportListFromSql.TabIndex = 11;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(2, 4);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(71, 24);
            this.label8.TabIndex = 50;
            this.label8.Text = "Export";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.TablePickerRadioButton);
            this.panel1.Controls.Add(this.RegexPatternTableSearchRadioButton);
            this.panel1.Controls.Add(this.CommaSeperatedListTableSearchRadioButton);
            this.panel1.Location = new System.Drawing.Point(12, 49);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(201, 100);
            this.panel1.TabIndex = 55;
            // 
            // TablePickerRadioButton
            // 
            this.TablePickerRadioButton.AutoSize = true;
            this.TablePickerRadioButton.Location = new System.Drawing.Point(13, 68);
            this.TablePickerRadioButton.Name = "TablePickerRadioButton";
            this.TablePickerRadioButton.Size = new System.Drawing.Size(87, 17);
            this.TablePickerRadioButton.TabIndex = 2;
            this.TablePickerRadioButton.TabStop = true;
            this.TablePickerRadioButton.Text = "Table picker:";
            this.TablePickerRadioButton.UseVisualStyleBackColor = true;
            // 
            // RegexPatternTableSearchRadioButton
            // 
            this.RegexPatternTableSearchRadioButton.AutoSize = true;
            this.RegexPatternTableSearchRadioButton.Location = new System.Drawing.Point(13, 45);
            this.RegexPatternTableSearchRadioButton.Name = "RegexPatternTableSearchRadioButton";
            this.RegexPatternTableSearchRadioButton.Size = new System.Drawing.Size(95, 17);
            this.RegexPatternTableSearchRadioButton.TabIndex = 1;
            this.RegexPatternTableSearchRadioButton.TabStop = true;
            this.RegexPatternTableSearchRadioButton.Text = "Regex pattern:";
            this.RegexPatternTableSearchRadioButton.UseVisualStyleBackColor = true;
            // 
            // CommaSeperatedListTableSearchRadioButton
            // 
            this.CommaSeperatedListTableSearchRadioButton.AutoSize = true;
            this.CommaSeperatedListTableSearchRadioButton.Location = new System.Drawing.Point(13, 3);
            this.CommaSeperatedListTableSearchRadioButton.Name = "CommaSeperatedListTableSearchRadioButton";
            this.CommaSeperatedListTableSearchRadioButton.Size = new System.Drawing.Size(128, 17);
            this.CommaSeperatedListTableSearchRadioButton.TabIndex = 0;
            this.CommaSeperatedListTableSearchRadioButton.TabStop = true;
            this.CommaSeperatedListTableSearchRadioButton.Text = "Comma seperated list:";
            this.CommaSeperatedListTableSearchRadioButton.UseVisualStyleBackColor = true;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label13.Location = new System.Drawing.Point(615, 7);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(70, 17);
            this.label13.TabIndex = 52;
            this.label13.Text = "* required";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label15.Location = new System.Drawing.Point(710, 7);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(154, 17);
            this.label15.TabIndex = 53;
            this.label15.Text = "^ conditionally required";
            // 
            // ServerComboBox
            // 
            this.ServerComboBox.FormattingEnabled = true;
            this.ServerComboBox.Items.AddRange(new object[] {
            "SQL04",
            "SQL05"});
            this.ServerComboBox.Location = new System.Drawing.Point(92, 7);
            this.ServerComboBox.Name = "ServerComboBox";
            this.ServerComboBox.Size = new System.Drawing.Size(169, 21);
            this.ServerComboBox.TabIndex = 54;
            this.ServerComboBox.Text = "SQL04";
            // 
            // DatabaseComboBox
            // 
            this.DatabaseComboBox.FormattingEnabled = true;
            this.DatabaseComboBox.Items.AddRange(new object[] {
            "BMW",
            "TEMP_BMW",
            "TEMP_EA",
            "TEMP_J",
            "TEMP_JC",
            "TEMP_LG",
            "TEMP_NS",
            "TEMP_OK"});
            this.DatabaseComboBox.Location = new System.Drawing.Point(404, 7);
            this.DatabaseComboBox.Name = "DatabaseComboBox";
            this.DatabaseComboBox.Size = new System.Drawing.Size(169, 21);
            this.DatabaseComboBox.TabIndex = 55;
            this.DatabaseComboBox.Text = "TEMP_OK";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(882, 660);
            this.Controls.Add(this.DatabaseComboBox);
            this.Controls.Add(this.ServerComboBox);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.Database);
            this.Controls.Add(this.Server);
            this.Controls.Add(this.ImportPanel);
            this.Controls.Add(this.ExportPanel);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ImportPanel.ResumeLayout(false);
            this.ImportPanel.PerformLayout();
            this.ExportPanel.ResumeLayout(false);
            this.ExportPanel.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label ImportPath;
        private System.Windows.Forms.Label Extension;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox ImportPathTextBox;
        private System.Windows.Forms.Label Server;
        private System.Windows.Forms.Label Database;
        private System.Windows.Forms.ListBox ExtensionListBoxImport;
        private System.Windows.Forms.TextBox TablesToExportCommaList;
        private System.Windows.Forms.TextBox TablesToExportRegex;
        private System.Windows.Forms.Label ImportPathNotes;
        private System.Windows.Forms.ListBox OutputTypeListBox;
        private System.Windows.Forms.TextBox ExportPathTextBox;
        private System.Windows.Forms.Label ExportPath;
        private System.Windows.Forms.Label ExportPathNotes;
        private System.Windows.Forms.CheckBox IncludeHeadersCheckBox;
        private System.Windows.Forms.ListBox ImportDelimeterListBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox ExportDelimeterListBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox ExportQualifierListBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button ExportButton;
        private System.Windows.Forms.Button ImportButton;
        private System.Windows.Forms.Panel ImportPanel;
        private System.Windows.Forms.Panel ExportPanel;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ListBox TablesToExportListFromSql;
        private System.Windows.Forms.Label FixedWidthColumnFilePathLabel;
        private System.Windows.Forms.TextBox FixedWidthColumnFilePathTextBox;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox ImportToSingleTableTextBox;
        private System.Windows.Forms.Button LoadSqlTables;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton TablePickerRadioButton;
        private System.Windows.Forms.RadioButton RegexPatternTableSearchRadioButton;
        private System.Windows.Forms.RadioButton CommaSeperatedListTableSearchRadioButton;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.ComboBox ServerComboBox;
        private System.Windows.Forms.ComboBox DatabaseComboBox;
        private System.Windows.Forms.Label ImportDelimeterLastSelected;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox FixedWidthColumnLengthMethodListBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label ExportQualifierLastSelected;
        private System.Windows.Forms.Label ExportDelimeterLastSelected;
    }
}

