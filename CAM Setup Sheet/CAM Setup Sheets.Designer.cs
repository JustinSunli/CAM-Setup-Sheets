namespace CAM_Setup_Sheets
{
    partial class SOLIDWORKS_CAM_Setup_Sheets
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SOLIDWORKS_CAM_Setup_Sheets));
            this.NCFileBrowseButton = new System.Windows.Forms.Button();
            this.LoadingInstructionsButton = new System.Windows.Forms.Button();
            this.SetupInstructionsTextBox = new System.Windows.Forms.TextBox();
            this.SetupInstructionsFile_label = new System.Windows.Forms.Label();
            this.CancelButton = new System.Windows.Forms.Button();
            this.OKButton = new System.Windows.Forms.Button();
            this.SelectTemplateFileButton = new System.Windows.Forms.Button();
            this.SetupSheetTemplatelabel = new System.Windows.Forms.Label();
            this.SelectSetupSheetFileToSaveToButton = new System.Windows.Forms.Button();
            this.SetupSheetFileTextBox = new System.Windows.Forms.TextBox();
            this.SaveSetupSheetTo_label = new System.Windows.Forms.Label();
            this.NewProgramCheckBox = new System.Windows.Forms.CheckBox();
            this.NCProgramtToExtractNBlocksFromabel = new System.Windows.Forms.Label();
            this.OutputType_groupBox = new System.Windows.Forms.GroupBox();
            this.SOLIDWORKSDrawing_radioButton = new System.Windows.Forms.RadioButton();
            this.ExcelFile_radioButton = new System.Windows.Forms.RadioButton();
            this.EditSetupSheetItemsbutton = new System.Windows.Forms.Button();
            this.CreateToolModelsButton = new System.Windows.Forms.Button();
            this.AddCreatedToolToSetupSheetCheckBox = new System.Windows.Forms.CheckBox();
            this.checkBoxExcelScreenUpdating = new System.Windows.Forms.CheckBox();
            this.NBlocksCheckBox = new System.Windows.Forms.CheckBox();
            this.NCFileName_For_NBlocks_textbox = new System.Windows.Forms.TextBox();
            this.checkBoxOutputAllToolsInCrib = new System.Windows.Forms.CheckBox();
            this.SetupSheetTemplateTextBox = new System.Windows.Forms.TextBox();
            this.ExcelTemplateWizard_button = new System.Windows.Forms.Button();
            this.OutputType_groupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // NCFileBrowseButton
            // 
            this.NCFileBrowseButton.Location = new System.Drawing.Point(733, 351);
            this.NCFileBrowseButton.Name = "NCFileBrowseButton";
            this.NCFileBrowseButton.Size = new System.Drawing.Size(30, 23);
            this.NCFileBrowseButton.TabIndex = 50;
            this.NCFileBrowseButton.Text = "...";
            this.NCFileBrowseButton.UseVisualStyleBackColor = true;
            this.NCFileBrowseButton.Click += new System.EventHandler(this.NCFileBrowseButton_Click);
            // 
            // LoadingInstructionsButton
            // 
            this.LoadingInstructionsButton.Location = new System.Drawing.Point(733, 221);
            this.LoadingInstructionsButton.Name = "LoadingInstructionsButton";
            this.LoadingInstructionsButton.Size = new System.Drawing.Size(30, 23);
            this.LoadingInstructionsButton.TabIndex = 44;
            this.LoadingInstructionsButton.Text = "...";
            this.LoadingInstructionsButton.UseVisualStyleBackColor = true;
            this.LoadingInstructionsButton.Click += new System.EventHandler(this.LoadingInstructionsButton_Click);
            // 
            // SetupInstructionsTextBox
            // 
            this.SetupInstructionsTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SetupInstructionsTextBox.ForeColor = System.Drawing.Color.Lime;
            this.SetupInstructionsTextBox.Location = new System.Drawing.Point(3, 224);
            this.SetupInstructionsTextBox.Name = "SetupInstructionsTextBox";
            this.SetupInstructionsTextBox.ReadOnly = true;
            this.SetupInstructionsTextBox.Size = new System.Drawing.Size(724, 22);
            this.SetupInstructionsTextBox.TabIndex = 43;
            this.SetupInstructionsTextBox.Tag = "Box";
            // 
            // SetupInstructionsFile_label
            // 
            this.SetupInstructionsFile_label.AutoSize = true;
            this.SetupInstructionsFile_label.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SetupInstructionsFile_label.ForeColor = System.Drawing.Color.Blue;
            this.SetupInstructionsFile_label.Location = new System.Drawing.Point(3, 203);
            this.SetupInstructionsFile_label.Name = "SetupInstructionsFile_label";
            this.SetupInstructionsFile_label.Size = new System.Drawing.Size(247, 16);
            this.SetupInstructionsFile_label.TabIndex = 42;
            this.SetupInstructionsFile_label.Text = "Setup Instructions File (.SLDDRW):";
            // 
            // CancelButton
            // 
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.Location = new System.Drawing.Point(687, 391);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 40;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            // 
            // OKButton
            // 
            this.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OKButton.Enabled = false;
            this.OKButton.Location = new System.Drawing.Point(606, 391);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new System.Drawing.Size(75, 23);
            this.OKButton.TabIndex = 39;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
            // 
            // SelectTemplateFileButton
            // 
            this.SelectTemplateFileButton.Location = new System.Drawing.Point(733, 129);
            this.SelectTemplateFileButton.Name = "SelectTemplateFileButton";
            this.SelectTemplateFileButton.Size = new System.Drawing.Size(30, 23);
            this.SelectTemplateFileButton.TabIndex = 38;
            this.SelectTemplateFileButton.Text = "...";
            this.SelectTemplateFileButton.UseVisualStyleBackColor = true;
            this.SelectTemplateFileButton.Click += new System.EventHandler(this.SelectTemplateFileButton_Click);
            // 
            // SetupSheetTemplatelabel
            // 
            this.SetupSheetTemplatelabel.AutoSize = true;
            this.SetupSheetTemplatelabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SetupSheetTemplatelabel.ForeColor = System.Drawing.Color.DodgerBlue;
            this.SetupSheetTemplatelabel.Location = new System.Drawing.Point(3, 111);
            this.SetupSheetTemplatelabel.Name = "SetupSheetTemplatelabel";
            this.SetupSheetTemplatelabel.Size = new System.Drawing.Size(204, 16);
            this.SetupSheetTemplatelabel.TabIndex = 36;
            this.SetupSheetTemplatelabel.Text = "Excel Setup Sheet Template";
            // 
            // SelectSetupSheetFileToSaveToButton
            // 
            this.SelectSetupSheetFileToSaveToButton.Location = new System.Drawing.Point(733, 173);
            this.SelectSetupSheetFileToSaveToButton.Name = "SelectSetupSheetFileToSaveToButton";
            this.SelectSetupSheetFileToSaveToButton.Size = new System.Drawing.Size(30, 23);
            this.SelectSetupSheetFileToSaveToButton.TabIndex = 35;
            this.SelectSetupSheetFileToSaveToButton.Text = "...";
            this.SelectSetupSheetFileToSaveToButton.UseVisualStyleBackColor = true;
            this.SelectSetupSheetFileToSaveToButton.Click += new System.EventHandler(this.SelectSetupSheetFileToSaveToButton_Click);
            // 
            // SetupSheetFileTextBox
            // 
            this.SetupSheetFileTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SetupSheetFileTextBox.ForeColor = System.Drawing.Color.Lime;
            this.SetupSheetFileTextBox.Location = new System.Drawing.Point(3, 176);
            this.SetupSheetFileTextBox.Name = "SetupSheetFileTextBox";
            this.SetupSheetFileTextBox.ReadOnly = true;
            this.SetupSheetFileTextBox.Size = new System.Drawing.Size(724, 22);
            this.SetupSheetFileTextBox.TabIndex = 34;
            this.SetupSheetFileTextBox.Tag = "Box";
            // 
            // SaveSetupSheetTo_label
            // 
            this.SaveSetupSheetTo_label.AutoSize = true;
            this.SaveSetupSheetTo_label.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SaveSetupSheetTo_label.ForeColor = System.Drawing.Color.Blue;
            this.SaveSetupSheetTo_label.Location = new System.Drawing.Point(3, 155);
            this.SaveSetupSheetTo_label.Name = "SaveSetupSheetTo_label";
            this.SaveSetupSheetTo_label.Size = new System.Drawing.Size(159, 16);
            this.SaveSetupSheetTo_label.TabIndex = 33;
            this.SaveSetupSheetTo_label.Text = "Save Setup Sheet To:";
            // 
            // NewProgramCheckBox
            // 
            this.NewProgramCheckBox.AutoSize = true;
            this.NewProgramCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NewProgramCheckBox.ForeColor = System.Drawing.Color.Red;
            this.NewProgramCheckBox.Location = new System.Drawing.Point(12, 64);
            this.NewProgramCheckBox.Name = "NewProgramCheckBox";
            this.NewProgramCheckBox.Size = new System.Drawing.Size(205, 20);
            this.NewProgramCheckBox.TabIndex = 45;
            this.NewProgramCheckBox.Text = "This is a NEW PROGRAM";
            this.NewProgramCheckBox.UseVisualStyleBackColor = true;
            this.NewProgramCheckBox.CheckedChanged += new System.EventHandler(this.NewProgramCheckBox_CheckedChanged);
            // 
            // NCProgramtToExtractNBlocksFromabel
            // 
            this.NCProgramtToExtractNBlocksFromabel.AutoSize = true;
            this.NCProgramtToExtractNBlocksFromabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NCProgramtToExtractNBlocksFromabel.ForeColor = System.Drawing.Color.Blue;
            this.NCProgramtToExtractNBlocksFromabel.Location = new System.Drawing.Point(3, 330);
            this.NCProgramtToExtractNBlocksFromabel.Name = "NCProgramtToExtractNBlocksFromabel";
            this.NCProgramtToExtractNBlocksFromabel.Size = new System.Drawing.Size(264, 16);
            this.NCProgramtToExtractNBlocksFromabel.TabIndex = 52;
            this.NCProgramtToExtractNBlocksFromabel.Text = "NC Program to extract N-Blocks from:";
            // 
            // OutputType_groupBox
            // 
            this.OutputType_groupBox.Controls.Add(this.SOLIDWORKSDrawing_radioButton);
            this.OutputType_groupBox.Controls.Add(this.ExcelFile_radioButton);
            this.OutputType_groupBox.Location = new System.Drawing.Point(12, 12);
            this.OutputType_groupBox.Name = "OutputType_groupBox";
            this.OutputType_groupBox.Size = new System.Drawing.Size(268, 46);
            this.OutputType_groupBox.TabIndex = 55;
            this.OutputType_groupBox.TabStop = false;
            this.OutputType_groupBox.Text = "Output Type";
            // 
            // SOLIDWORKSDrawing_radioButton
            // 
            this.SOLIDWORKSDrawing_radioButton.AutoSize = true;
            this.SOLIDWORKSDrawing_radioButton.Location = new System.Drawing.Point(110, 19);
            this.SOLIDWORKSDrawing_radioButton.Name = "SOLIDWORKSDrawing_radioButton";
            this.SOLIDWORKSDrawing_radioButton.Size = new System.Drawing.Size(140, 17);
            this.SOLIDWORKSDrawing_radioButton.TabIndex = 54;
            this.SOLIDWORKSDrawing_radioButton.Text = "SOLIDWORKS Drawing";
            this.SOLIDWORKSDrawing_radioButton.UseVisualStyleBackColor = true;
            this.SOLIDWORKSDrawing_radioButton.CheckedChanged += new System.EventHandler(this.SOLIDWORKSDrawing_radioButton_CheckedChanged);
            // 
            // ExcelFile_radioButton
            // 
            this.ExcelFile_radioButton.AutoSize = true;
            this.ExcelFile_radioButton.Checked = true;
            this.ExcelFile_radioButton.Location = new System.Drawing.Point(6, 19);
            this.ExcelFile_radioButton.Name = "ExcelFile_radioButton";
            this.ExcelFile_radioButton.Size = new System.Drawing.Size(70, 17);
            this.ExcelFile_radioButton.TabIndex = 54;
            this.ExcelFile_radioButton.TabStop = true;
            this.ExcelFile_radioButton.Text = "Excel File";
            this.ExcelFile_radioButton.UseVisualStyleBackColor = true;
            this.ExcelFile_radioButton.CheckedChanged += new System.EventHandler(this.ExcelFile_radioButton_CheckedChanged);
            // 
            // EditSetupSheetItemsbutton
            // 
            this.EditSetupSheetItemsbutton.Location = new System.Drawing.Point(542, 31);
            this.EditSetupSheetItemsbutton.Name = "EditSetupSheetItemsbutton";
            this.EditSetupSheetItemsbutton.Size = new System.Drawing.Size(161, 23);
            this.EditSetupSheetItemsbutton.TabIndex = 56;
            this.EditSetupSheetItemsbutton.Text = "Edit Setup Sheet Items";
            this.EditSetupSheetItemsbutton.UseVisualStyleBackColor = true;
            this.EditSetupSheetItemsbutton.Click += new System.EventHandler(this.EditSetupSheetItemsbutton_Click);
            // 
            // CreateToolModelsButton
            // 
            this.CreateToolModelsButton.Location = new System.Drawing.Point(542, 60);
            this.CreateToolModelsButton.Name = "CreateToolModelsButton";
            this.CreateToolModelsButton.Size = new System.Drawing.Size(115, 23);
            this.CreateToolModelsButton.TabIndex = 57;
            this.CreateToolModelsButton.Text = "Create Tool Models";
            this.CreateToolModelsButton.UseVisualStyleBackColor = true;
            this.CreateToolModelsButton.Click += new System.EventHandler(this.CreateToolModelsButton_Click);
            // 
            // AddCreatedToolToSetupSheetCheckBox
            // 
            this.AddCreatedToolToSetupSheetCheckBox.AutoSize = true;
            this.AddCreatedToolToSetupSheetCheckBox.Checked = global::CAM_Setup_Sheets.Properties.Settings.Default.AddCreatedToolsToSetupSheet;
            this.AddCreatedToolToSetupSheetCheckBox.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::CAM_Setup_Sheets.Properties.Settings.Default, "AddCreatedToolsToSetupSheet", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.AddCreatedToolToSetupSheetCheckBox.Location = new System.Drawing.Point(542, 89);
            this.AddCreatedToolToSetupSheetCheckBox.Name = "AddCreatedToolToSetupSheetCheckBox";
            this.AddCreatedToolToSetupSheetCheckBox.Size = new System.Drawing.Size(220, 17);
            this.AddCreatedToolToSetupSheetCheckBox.TabIndex = 58;
            this.AddCreatedToolToSetupSheetCheckBox.Text = "Add Created Tool Models to Setup Sheet";
            this.AddCreatedToolToSetupSheetCheckBox.UseVisualStyleBackColor = true;
            // 
            // checkBoxExcelScreenUpdating
            // 
            this.checkBoxExcelScreenUpdating.AutoSize = true;
            this.checkBoxExcelScreenUpdating.Checked = global::CAM_Setup_Sheets.Properties.Settings.Default.ExcelScreenUpdating;
            this.checkBoxExcelScreenUpdating.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxExcelScreenUpdating.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::CAM_Setup_Sheets.Properties.Settings.Default, "ExcelScreenUpdating", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.checkBoxExcelScreenUpdating.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxExcelScreenUpdating.ForeColor = System.Drawing.Color.Blue;
            this.checkBoxExcelScreenUpdating.Location = new System.Drawing.Point(3, 270);
            this.checkBoxExcelScreenUpdating.Name = "checkBoxExcelScreenUpdating";
            this.checkBoxExcelScreenUpdating.Size = new System.Drawing.Size(185, 20);
            this.checkBoxExcelScreenUpdating.TabIndex = 53;
            this.checkBoxExcelScreenUpdating.Text = "Excel Screen Updating";
            this.checkBoxExcelScreenUpdating.UseVisualStyleBackColor = true;
            this.checkBoxExcelScreenUpdating.CheckedChanged += new System.EventHandler(this.checkBoxExcelScreenUpdating_CheckedChanged);
            // 
            // NBlocksCheckBox
            // 
            this.NBlocksCheckBox.AutoSize = true;
            this.NBlocksCheckBox.Checked = global::CAM_Setup_Sheets.Properties.Settings.Default.IncludeNBlocksOnSetupSheet;
            this.NBlocksCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.NBlocksCheckBox.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::CAM_Setup_Sheets.Properties.Settings.Default, "IncludeNBlocksOnSetupSheet", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.NBlocksCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NBlocksCheckBox.ForeColor = System.Drawing.Color.Blue;
            this.NBlocksCheckBox.Location = new System.Drawing.Point(3, 296);
            this.NBlocksCheckBox.Name = "NBlocksCheckBox";
            this.NBlocksCheckBox.Size = new System.Drawing.Size(311, 20);
            this.NBlocksCheckBox.TabIndex = 51;
            this.NBlocksCheckBox.Text = "Include N-Block Numbers on Setup Sheet";
            this.NBlocksCheckBox.UseVisualStyleBackColor = true;
            this.NBlocksCheckBox.CheckedChanged += new System.EventHandler(this.NBlocksCheckBox_CheckedChanged);
            // 
            // NCFileName_For_NBlocks_textbox
            // 
            this.NCFileName_For_NBlocks_textbox.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::CAM_Setup_Sheets.Properties.Settings.Default, "NCFileForNBlocks", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.NCFileName_For_NBlocks_textbox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NCFileName_For_NBlocks_textbox.ForeColor = System.Drawing.Color.Lime;
            this.NCFileName_For_NBlocks_textbox.Location = new System.Drawing.Point(3, 351);
            this.NCFileName_For_NBlocks_textbox.Name = "NCFileName_For_NBlocks_textbox";
            this.NCFileName_For_NBlocks_textbox.ReadOnly = true;
            this.NCFileName_For_NBlocks_textbox.Size = new System.Drawing.Size(724, 22);
            this.NCFileName_For_NBlocks_textbox.TabIndex = 49;
            this.NCFileName_For_NBlocks_textbox.Tag = "Box";
            this.NCFileName_For_NBlocks_textbox.Text = global::CAM_Setup_Sheets.Properties.Settings.Default.NCFileForNBlocks;
            // 
            // checkBoxOutputAllToolsInCrib
            // 
            this.checkBoxOutputAllToolsInCrib.AutoSize = true;
            this.checkBoxOutputAllToolsInCrib.Checked = global::CAM_Setup_Sheets.Properties.Settings.Default.OutputAllTools;
            this.checkBoxOutputAllToolsInCrib.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::CAM_Setup_Sheets.Properties.Settings.Default, "OutputAllTools", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.checkBoxOutputAllToolsInCrib.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxOutputAllToolsInCrib.ForeColor = System.Drawing.Color.Blue;
            this.checkBoxOutputAllToolsInCrib.Location = new System.Drawing.Point(223, 64);
            this.checkBoxOutputAllToolsInCrib.Name = "checkBoxOutputAllToolsInCrib";
            this.checkBoxOutputAllToolsInCrib.Size = new System.Drawing.Size(184, 20);
            this.checkBoxOutputAllToolsInCrib.TabIndex = 48;
            this.checkBoxOutputAllToolsInCrib.Text = "Output all Tools in Crib";
            this.checkBoxOutputAllToolsInCrib.UseVisualStyleBackColor = true;
            this.checkBoxOutputAllToolsInCrib.CheckedChanged += new System.EventHandler(this.checkBoxOutputAllToolsInCrib_CheckedChanged);
            // 
            // SetupSheetTemplateTextBox
            // 
            this.SetupSheetTemplateTextBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.SetupSheetTemplateTextBox.Location = new System.Drawing.Point(3, 132);
            this.SetupSheetTemplateTextBox.Name = "SetupSheetTemplateTextBox";
            this.SetupSheetTemplateTextBox.ReadOnly = true;
            this.SetupSheetTemplateTextBox.Size = new System.Drawing.Size(724, 20);
            this.SetupSheetTemplateTextBox.TabIndex = 37;
            this.SetupSheetTemplateTextBox.Tag = "Box";
            this.SetupSheetTemplateTextBox.Text = global::CAM_Setup_Sheets.Properties.Settings.Default.ExcelDefaultTemplateFileName;
            // 
            // ExcelTemplateWizard_button
            // 
            this.ExcelTemplateWizard_button.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.ExcelTemplateWizard_button.Location = new System.Drawing.Point(359, 31);
            this.ExcelTemplateWizard_button.Name = "ExcelTemplateWizard_button";
            this.ExcelTemplateWizard_button.Size = new System.Drawing.Size(161, 23);
            this.ExcelTemplateWizard_button.TabIndex = 59;
            this.ExcelTemplateWizard_button.Text = "Excel Template Wizard";
            this.ExcelTemplateWizard_button.UseVisualStyleBackColor = false;
            this.ExcelTemplateWizard_button.Click += new System.EventHandler(this.ExcelTemplateWizard_button_Click);
            // 
            // SOLIDWORKS_CAM_Setup_Sheets
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(770, 425);
            this.Controls.Add(this.ExcelTemplateWizard_button);
            this.Controls.Add(this.AddCreatedToolToSetupSheetCheckBox);
            this.Controls.Add(this.CreateToolModelsButton);
            this.Controls.Add(this.EditSetupSheetItemsbutton);
            this.Controls.Add(this.OutputType_groupBox);
            this.Controls.Add(this.checkBoxExcelScreenUpdating);
            this.Controls.Add(this.NCProgramtToExtractNBlocksFromabel);
            this.Controls.Add(this.NBlocksCheckBox);
            this.Controls.Add(this.NCFileBrowseButton);
            this.Controls.Add(this.NCFileName_For_NBlocks_textbox);
            this.Controls.Add(this.checkBoxOutputAllToolsInCrib);
            this.Controls.Add(this.NewProgramCheckBox);
            this.Controls.Add(this.LoadingInstructionsButton);
            this.Controls.Add(this.SetupInstructionsTextBox);
            this.Controls.Add(this.SetupInstructionsFile_label);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.SelectTemplateFileButton);
            this.Controls.Add(this.SetupSheetTemplateTextBox);
            this.Controls.Add(this.SetupSheetTemplatelabel);
            this.Controls.Add(this.SelectSetupSheetFileToSaveToButton);
            this.Controls.Add(this.SetupSheetFileTextBox);
            this.Controls.Add(this.SaveSetupSheetTo_label);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SOLIDWORKS_CAM_Setup_Sheets";
            this.Text = "SOLIDWORKS CAM Setup Sheets";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SOLIDWORKS_CAM_Setup_Sheets_FormClosing);
            this.Load += new System.EventHandler(this.SOLIDWORKS_CAM_Setup_Sheets_Load);
            this.OutputType_groupBox.ResumeLayout(false);
            this.OutputType_groupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox NBlocksCheckBox;
        private System.Windows.Forms.Button NCFileBrowseButton;
        private System.Windows.Forms.TextBox NCFileName_For_NBlocks_textbox;
        public System.Windows.Forms.CheckBox checkBoxOutputAllToolsInCrib;
        private System.Windows.Forms.Button LoadingInstructionsButton;
        private System.Windows.Forms.TextBox SetupInstructionsTextBox;
        private System.Windows.Forms.Label SetupInstructionsFile_label;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.Button OKButton;
        private System.Windows.Forms.Button SelectTemplateFileButton;
        private System.Windows.Forms.TextBox SetupSheetTemplateTextBox;
        private System.Windows.Forms.Label SetupSheetTemplatelabel;
        private System.Windows.Forms.Button SelectSetupSheetFileToSaveToButton;
        private System.Windows.Forms.TextBox SetupSheetFileTextBox;
        private System.Windows.Forms.Label SaveSetupSheetTo_label;
        private System.Windows.Forms.CheckBox NewProgramCheckBox;
        private System.Windows.Forms.Label NCProgramtToExtractNBlocksFromabel;
        private System.Windows.Forms.CheckBox checkBoxExcelScreenUpdating;
        private System.Windows.Forms.RadioButton ExcelFile_radioButton;
        private System.Windows.Forms.RadioButton SOLIDWORKSDrawing_radioButton;
        private System.Windows.Forms.GroupBox OutputType_groupBox;
        private System.Windows.Forms.Button EditSetupSheetItemsbutton;
        private System.Windows.Forms.Button CreateToolModelsButton;
        private System.Windows.Forms.CheckBox AddCreatedToolToSetupSheetCheckBox;
        private System.Windows.Forms.Button ExcelTemplateWizard_button;
    }
}