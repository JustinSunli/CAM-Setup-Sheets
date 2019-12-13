namespace CAM_Setup_Sheets
{
    partial class CWToolCribToolListSelectionForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CWToolCribToolListSelectionForm));
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridViewCWToolCribTools = new System.Windows.Forms.DataGridView();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.FolderPathTextBox = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.buttonSelectFolder = new System.Windows.Forms.Button();
            this.checkBoxDisableScreenUpdating = new System.Windows.Forms.CheckBox();
            this.checkBoxCreatAssyOfTools = new System.Windows.Forms.CheckBox();
            this.numericUpDownXDistanceBetweenTools = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.pictureBoxCWToolCuttingPortionColor = new System.Windows.Forms.PictureBox();
            this.pictureBoxCWToolHolderColor = new System.Windows.Forms.PictureBox();
            this.pictureBoxCWToolShankColor = new System.Windows.Forms.PictureBox();
            this.CreateToolAsSTL_checkBox = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCWToolCribTools)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownXDistanceBetweenTools)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxCWToolCuttingPortionColor)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxCWToolHolderColor)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxCWToolShankColor)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 121);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(160, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Tool(s) to Create from List";
            // 
            // dataGridViewCWToolCribTools
            // 
            this.dataGridViewCWToolCribTools.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewCWToolCribTools.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridViewCWToolCribTools.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewCWToolCribTools.Location = new System.Drawing.Point(12, 137);
            this.dataGridViewCWToolCribTools.Name = "dataGridViewCWToolCribTools";
            this.dataGridViewCWToolCribTools.ReadOnly = true;
            this.dataGridViewCWToolCribTools.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewCWToolCribTools.Size = new System.Drawing.Size(692, 387);
            this.dataGridViewCWToolCribTools.TabIndex = 1;
            this.dataGridViewCWToolCribTools.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewCWToolCribTools_CellClick);
            this.dataGridViewCWToolCribTools.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridViewCWToolCribTools_CellMouseDoubleClick);
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOK.Location = new System.Drawing.Point(540, 530);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 2;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(630, 530);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 3;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(88, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(127, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Tool Cutting Portion Color";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(88, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Tool Shank Color";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(88, 67);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Holder Color";
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.MyComputer;
            // 
            // FolderPathTextBox
            // 
            this.FolderPathTextBox.Location = new System.Drawing.Point(258, 30);
            this.FolderPathTextBox.Name = "FolderPathTextBox";
            this.FolderPathTextBox.ReadOnly = true;
            this.FolderPathTextBox.Size = new System.Drawing.Size(401, 20);
            this.FolderPathTextBox.TabIndex = 19;
            this.FolderPathTextBox.TextChanged += new System.EventHandler(this.FolderPathTextBox_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(255, 12);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(286, 13);
            this.label5.TabIndex = 20;
            this.label5.Text = "Folder for created models (Will be created  if it doesn\'t exist)";
            // 
            // buttonSelectFolder
            // 
            this.buttonSelectFolder.Location = new System.Drawing.Point(665, 28);
            this.buttonSelectFolder.Name = "buttonSelectFolder";
            this.buttonSelectFolder.Size = new System.Drawing.Size(26, 23);
            this.buttonSelectFolder.TabIndex = 21;
            this.buttonSelectFolder.Text = "...";
            this.buttonSelectFolder.UseVisualStyleBackColor = true;
            this.buttonSelectFolder.Click += new System.EventHandler(this.buttonSelectFolder_Click);
            // 
            // checkBoxDisableScreenUpdating
            // 
            this.checkBoxDisableScreenUpdating.AutoSize = true;
            this.checkBoxDisableScreenUpdating.Location = new System.Drawing.Point(258, 56);
            this.checkBoxDisableScreenUpdating.Name = "checkBoxDisableScreenUpdating";
            this.checkBoxDisableScreenUpdating.Size = new System.Drawing.Size(213, 17);
            this.checkBoxDisableScreenUpdating.TabIndex = 23;
            this.checkBoxDisableScreenUpdating.Text = "Disable screen updating during creation";
            this.checkBoxDisableScreenUpdating.UseVisualStyleBackColor = true;
            this.checkBoxDisableScreenUpdating.CheckedChanged += new System.EventHandler(this.checkBoxDisableScreenUpdating_CheckedChanged);
            // 
            // checkBoxCreatAssyOfTools
            // 
            this.checkBoxCreatAssyOfTools.AutoSize = true;
            this.checkBoxCreatAssyOfTools.Location = new System.Drawing.Point(258, 79);
            this.checkBoxCreatAssyOfTools.Name = "checkBoxCreatAssyOfTools";
            this.checkBoxCreatAssyOfTools.Size = new System.Drawing.Size(190, 17);
            this.checkBoxCreatAssyOfTools.TabIndex = 23;
            this.checkBoxCreatAssyOfTools.Text = "Create Assembly of Selected Tools";
            this.checkBoxCreatAssyOfTools.UseVisualStyleBackColor = true;
            this.checkBoxCreatAssyOfTools.CheckedChanged += new System.EventHandler(this.checkBoxCreatAssyOfTools_CheckedChanged);
            // 
            // numericUpDownXDistanceBetweenTools
            // 
            this.numericUpDownXDistanceBetweenTools.DecimalPlaces = 3;
            this.numericUpDownXDistanceBetweenTools.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numericUpDownXDistanceBetweenTools.Location = new System.Drawing.Point(442, 102);
            this.numericUpDownXDistanceBetweenTools.Name = "numericUpDownXDistanceBetweenTools";
            this.numericUpDownXDistanceBetweenTools.Size = new System.Drawing.Size(120, 20);
            this.numericUpDownXDistanceBetweenTools.TabIndex = 24;
            this.numericUpDownXDistanceBetweenTools.ValueChanged += new System.EventHandler(this.numericUpDownXDistanceBetweenTools_ValueChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(255, 104);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(181, 13);
            this.label6.TabIndex = 25;
            this.label6.Text = "Distance Between Tools in Assembly";
            // 
            // pictureBoxCWToolCuttingPortionColor
            // 
            this.pictureBoxCWToolCuttingPortionColor.BackColor = global::CAM_Setup_Sheets.Properties.Settings.Default.CWToolCuttingPortionColor;
            this.pictureBoxCWToolCuttingPortionColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBoxCWToolCuttingPortionColor.DataBindings.Add(new System.Windows.Forms.Binding("BackColor", global::CAM_Setup_Sheets.Properties.Settings.Default, "CWToolCuttingPortionColor", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.pictureBoxCWToolCuttingPortionColor.Location = new System.Drawing.Point(15, 5);
            this.pictureBoxCWToolCuttingPortionColor.Name = "pictureBoxCWToolCuttingPortionColor";
            this.pictureBoxCWToolCuttingPortionColor.Size = new System.Drawing.Size(68, 23);
            this.pictureBoxCWToolCuttingPortionColor.TabIndex = 16;
            this.pictureBoxCWToolCuttingPortionColor.TabStop = false;
            this.pictureBoxCWToolCuttingPortionColor.Click += new System.EventHandler(this.pictureBoxCWToolCuttingPortionColor_Click);
            // 
            // pictureBoxCWToolHolderColor
            // 
            this.pictureBoxCWToolHolderColor.BackColor = global::CAM_Setup_Sheets.Properties.Settings.Default.CWHolderColor;
            this.pictureBoxCWToolHolderColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBoxCWToolHolderColor.DataBindings.Add(new System.Windows.Forms.Binding("BackColor", global::CAM_Setup_Sheets.Properties.Settings.Default, "CWHolderColor", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.pictureBoxCWToolHolderColor.Location = new System.Drawing.Point(15, 62);
            this.pictureBoxCWToolHolderColor.Name = "pictureBoxCWToolHolderColor";
            this.pictureBoxCWToolHolderColor.Size = new System.Drawing.Size(68, 23);
            this.pictureBoxCWToolHolderColor.TabIndex = 14;
            this.pictureBoxCWToolHolderColor.TabStop = false;
            this.pictureBoxCWToolHolderColor.Click += new System.EventHandler(this.pictureBoxCWToolHolderColor_Click);
            // 
            // pictureBoxCWToolShankColor
            // 
            this.pictureBoxCWToolShankColor.BackColor = global::CAM_Setup_Sheets.Properties.Settings.Default.CWToolShankColor;
            this.pictureBoxCWToolShankColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBoxCWToolShankColor.DataBindings.Add(new System.Windows.Forms.Binding("BackColor", global::CAM_Setup_Sheets.Properties.Settings.Default, "CWToolShankColor", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.pictureBoxCWToolShankColor.Location = new System.Drawing.Point(15, 34);
            this.pictureBoxCWToolShankColor.Name = "pictureBoxCWToolShankColor";
            this.pictureBoxCWToolShankColor.Size = new System.Drawing.Size(68, 23);
            this.pictureBoxCWToolShankColor.TabIndex = 12;
            this.pictureBoxCWToolShankColor.TabStop = false;
            this.pictureBoxCWToolShankColor.Click += new System.EventHandler(this.pictureBoxCWToolShankColor_Click);
            // 
            // CreateToolAsSTL_checkBox
            // 
            this.CreateToolAsSTL_checkBox.AutoSize = true;
            this.CreateToolAsSTL_checkBox.Location = new System.Drawing.Point(477, 57);
            this.CreateToolAsSTL_checkBox.Name = "CreateToolAsSTL_checkBox";
            this.CreateToolAsSTL_checkBox.Size = new System.Drawing.Size(118, 17);
            this.CreateToolAsSTL_checkBox.TabIndex = 26;
            this.CreateToolAsSTL_checkBox.Text = "Create Tool as STL";
            this.CreateToolAsSTL_checkBox.UseVisualStyleBackColor = true;
            this.CreateToolAsSTL_checkBox.CheckedChanged += new System.EventHandler(this.CreateToolAsSTL_checkBox_CheckedChanged);
            // 
            // CWToolCribToolListSelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(717, 565);
            this.Controls.Add(this.CreateToolAsSTL_checkBox);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.numericUpDownXDistanceBetweenTools);
            this.Controls.Add(this.checkBoxCreatAssyOfTools);
            this.Controls.Add(this.checkBoxDisableScreenUpdating);
            this.Controls.Add(this.buttonSelectFolder);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.FolderPathTextBox);
            this.Controls.Add(this.pictureBoxCWToolCuttingPortionColor);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pictureBoxCWToolHolderColor);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.pictureBoxCWToolShankColor);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.dataGridViewCWToolCribTools);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CWToolCribToolListSelectionForm";
            this.Text = "CAMWorks ToolCrib Tools - Select Tools to create new Solid Tool Model";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CWToolCribToolListSelectionForm_FormClosing);
            this.Load += new System.EventHandler(this.CWToolCribToolListSelectionForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCWToolCribTools)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownXDistanceBetweenTools)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxCWToolCuttingPortionColor)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxCWToolHolderColor)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxCWToolShankColor)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridViewCWToolCribTools;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.PictureBox pictureBoxCWToolShankColor;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.PictureBox pictureBoxCWToolHolderColor;
        private System.Windows.Forms.PictureBox pictureBoxCWToolCuttingPortionColor;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox FolderPathTextBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button buttonSelectFolder;
        private System.Windows.Forms.CheckBox checkBoxDisableScreenUpdating;
        private System.Windows.Forms.CheckBox checkBoxCreatAssyOfTools;
        private System.Windows.Forms.NumericUpDown numericUpDownXDistanceBetweenTools;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox CreateToolAsSTL_checkBox;
    }
}