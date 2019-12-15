using SwConst;
using System;
using System.IO;
using System.Windows.Forms;

namespace CAM_Setup_Sheets
{
    public partial class SOLIDWORKS_CAM_Setup_Sheets : Form
    {
        public SOLIDWORKS_CAM_Setup_Sheets()
        {
            InitializeComponent();
            // Uncomment to reset properties
            //Properties.Settings.Default.Reset();

            SOLIDWORKSDrawing_radioButton.Checked = Properties.Settings.Default.OutputTypeSWDrawing;
            ExcelFile_radioButton.Checked = Properties.Settings.Default.OutputTypeExcel;
            //if(SOLIDWORKSDrawing_radioButton.Checked && !Properties.Settings.Default.IncludeNBlocksOnSetupSheet)
            //{
            //    OKButton.Enabled = true;
            //}
        }

        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        private void NBlocksCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            //if (NBlocksCheckBox.Checked)
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2)
            //        && (NCFileName_For_NBlocks_textbox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}

            //else
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}
            Properties.Settings.Default.IncludeNBlocksOnSetupSheet = NBlocksCheckBox.Checked;
            CheckForOKButtonEnable();
        }

        private void NewProgramCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (NewProgramCheckBox.Checked)
                CAM_Setup_Sheets_Addin.bNewProgram = true;

            if (!NewProgramCheckBox.Checked)
                CAM_Setup_Sheets_Addin.bNewProgram = false;
        }

        private void checkBoxOutputAllToolsInCrib_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxOutputAllToolsInCrib.Checked)
            {
                CAM_Setup_Sheets_Addin.bOutputAllToolsinCrib = true;
            }
            if (!checkBoxOutputAllToolsInCrib.Checked)
            {
                CAM_Setup_Sheets_Addin.bOutputAllToolsinCrib = false;
            }
        }

        private void SelectSetupSheetFileToSaveToButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog fdlg = new SaveFileDialog();
            fdlg.DefaultExt = "xls";
            fdlg.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*";
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                CAM_Setup_Sheets_Addin.sExcelOutputFileName = fdlg.FileName;
                SetupSheetFileTextBox.Text = fdlg.FileName;
            }

            FileInfo fi;

            if (fdlg.FileName != String.Empty)
            {
                if (File.Exists(fdlg.FileName))
                {
                    fi = new FileInfo(fdlg.FileName);
                    if (IsFileLocked(fi))
                    {
                        MessageBox.Show(fdlg.FileName + " is open.\nPlease close and try again.", "File is in use!!!!!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
            }

            //if (NBlocksCheckBox.Checked)
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2)
            //        && (NCFileName_For_NBlocks_textbox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}

            //else
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}
            CheckForOKButtonEnable();
        }

        private void LoadingInstructionsButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog odlg = new OpenFileDialog();
            //odlg.InitialDirectory = Path.Combine(SOLIDWORKS_CAM_Setup_Sheets_Addin.swpath, PAC_Setup_Sheets.MainClass.SetupSheetPath);
            odlg.InitialDirectory = CAM_Setup_Sheets_Addin.swpath;
            odlg.DefaultExt = "SLDDRW";
            odlg.Filter = ".SLDDRW Files (*.SLDDRW)|*.SLDDRW|All Files (*.*)|*.*";
            if (odlg.ShowDialog() == DialogResult.OK)
            {
                CAM_Setup_Sheets_Addin.sSetupInstructionsFilename = odlg.FileName;
                SetupInstructionsTextBox.Text = odlg.FileName;
            }

            //if (NBlocksCheckBox.Checked)
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2)
            //        && (NCFileName_For_NBlocks_textbox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}

            //else
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}
            CheckForOKButtonEnable();
        }

        private void SelectTemplateFileButton_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.OutputTypeExcel)
            {
                OpenFileDialog odlg = new OpenFileDialog();
                odlg.DefaultExt = "xls";
                odlg.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*";
                if (odlg.ShowDialog() == DialogResult.OK)

                {
                    Properties.Settings.Default.ExcelDefaultTemplateFileName = odlg.FileName;
                    SetupSheetTemplateTextBox.Text = Properties.Settings.Default.ExcelDefaultTemplateFileName;
                }
            }

            if (Properties.Settings.Default.OutputTypeSWDrawing)
            {
                String TemplateLocation = CAM_Setup_Sheets_Addin.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates);
                string[] Locations = TemplateLocation.Split(';');
                OpenFileDialog odlg = new OpenFileDialog();
                odlg.DefaultExt = "drwdot";
                odlg.Filter = "SOLIDWORKS Drawing Templates (*.drwdot)|*.drwdot|All Files (*.*)|*.*";
                odlg.InitialDirectory = Locations[0];
                if (odlg.ShowDialog() == DialogResult.OK)

                {
                    Properties.Settings.Default.SOLIDWORKSDefaultDrawingTemplate = odlg.FileName;
                    SetupSheetTemplateTextBox.Text = Properties.Settings.Default.SOLIDWORKSDefaultDrawingTemplate;
                }
            }

            //if (NBlocksCheckBox.Checked)
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2)
            //        && (NCFileName_For_NBlocks_textbox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}

            //else
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}
            CheckForOKButtonEnable();
        }

        private void NCFileBrowseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog odlg = new OpenFileDialog();
            //odlg.InitialDirectory = Path.Combine(PAC_Setup_Sheets.MainClass.swpath, PAC_Setup_Sheets.MainClass.NCFilePath);
            odlg.InitialDirectory = CAM_Setup_Sheets_Addin.swpath;
            odlg.DefaultExt = "NC";
            odlg.Filter = ".NC Files (*.NC)|*.NC|All Files (*.*)|*.*";
            if (odlg.ShowDialog() == DialogResult.OK)
            {
                CAM_Setup_Sheets_Addin.sNCFilename = odlg.FileName;
                NCFileName_For_NBlocks_textbox.Text = odlg.FileName;
            }

            //if (NBlocksCheckBox.Checked)
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2)
            //        && (NCFileName_For_NBlocks_textbox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}

            //else
            //{
            //    if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //        && (SetupSheetFileTextBox.Text.Length > 2))
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}
            CheckForOKButtonEnable();
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private void checkBoxExcelScreenUpdating_CheckedChanged(object sender, EventArgs e)
        {
            CAM_Setup_Sheets_Addin.ExcelScreenUpdating = checkBoxExcelScreenUpdating.Checked;
        }

        private void EditSetupSheetItemsbutton_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            //this.Hide();
            Settings frm = new Settings();
            frm.ShowDialog();
            this.WindowState = FormWindowState.Normal;
            //this.Show();
        }

        private void ExcelFile_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.OutputTypeExcel = ExcelFile_radioButton.Checked;
            Properties.Settings.Default.Save();
            //if (ExcelFile_radioButton.Checked)
            //{
            //    if (NBlocksCheckBox.Checked)
            //    {
            //        if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //            && (SetupSheetFileTextBox.Text.Length > 2)
            //            && (NCFileName_For_NBlocks_textbox.Text.Length > 2))
            //        {
            //            OKButton.Enabled = true;
            //        }
            //    }

            //    else if ((ExcelSetupSheetTemplateTextBox.Text.Length > 2)
            //            && (SetupSheetFileTextBox.Text.Length > 2))
            //    {

            //        {
            //            OKButton.Enabled = true;
            //        }
            //    }
            //    else
            //    {
            //        OKButton.Enabled = false;
            //    }
            //}
            CheckForOKButtonEnable();
        }

        private void SOLIDWORKSDrawing_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.OutputTypeSWDrawing = SOLIDWORKSDrawing_radioButton.Checked;
            SetupSheetTemplateTextBox.Text = Properties.Settings.Default.SOLIDWORKSDefaultDrawingTemplate;
            Properties.Settings.Default.Save();
            //if (!NBlocksCheckBox.Checked)
            //{
            //    {
            //        OKButton.Enabled = true;
            //    }
            //}
            CheckForOKButtonEnable();
        }

        private void CheckForOKButtonEnable()
        {
            //Output Type = Excel
            if (Properties.Settings.Default.OutputTypeExcel)
            {
                SetupSheetTemplatelabel.Visible = true;
                SetupSheetTemplatelabel.Text = "Excel Setup Sheet Template";
                SetupSheetTemplateTextBox.Visible = true;
                SelectTemplateFileButton.Visible = true;
                SaveSetupSheetTo_label.Visible = true;
                SetupSheetFileTextBox.Visible = true;
                SelectSetupSheetFileToSaveToButton.Visible = true;
                SetupInstructionsFile_label.Visible = true;
                SetupInstructionsTextBox.Visible = true;
                LoadingInstructionsButton.Visible = true;
                checkBoxExcelScreenUpdating.Visible = true;
                if (NBlocksCheckBox.Checked)
                {
                    NCProgramtToExtractNBlocksFromabel.Visible = true;
                    NCFileName_For_NBlocks_textbox.Visible = true;
                    NCFileBrowseButton.Visible = true;
                }
                else
                {
                    NCProgramtToExtractNBlocksFromabel.Visible = false;
                    NCFileName_For_NBlocks_textbox.Visible = false;
                    NCFileBrowseButton.Visible = false;
                }

                // Now see if we can enable the OK Button
                if (NBlocksCheckBox.Checked)
                {
                    if ((SetupSheetTemplateTextBox.Text.Length > 2)
                        && (SetupSheetFileTextBox.Text.Length > 2)
                        && (NCFileName_For_NBlocks_textbox.Text.Length > 2))
                    {
                        OKButton.Enabled = true;
                    }
                }

                else if ((SetupSheetTemplateTextBox.Text.Length > 2)
                        && (SetupSheetFileTextBox.Text.Length > 2))
                {

                    {
                        OKButton.Enabled = true;
                    }
                }
                else
                {
                    OKButton.Enabled = false;
                }
            }

            // Output Type = SOLIDWORKS Drawing
            if (Properties.Settings.Default.OutputTypeSWDrawing)
            {
                SetupSheetTemplatelabel.Visible = true;
                SetupSheetTemplatelabel.Text = "SOLIDWORKS Setup Sheet Template";
                SetupSheetTemplateTextBox.Visible = true;
                SelectTemplateFileButton.Visible = true;
                SaveSetupSheetTo_label.Visible = false;
                SetupSheetFileTextBox.Visible = false;
                SelectSetupSheetFileToSaveToButton.Visible = false;
                SetupInstructionsFile_label.Visible = false;
                SetupInstructionsTextBox.Visible = false;
                LoadingInstructionsButton.Visible = false;
                checkBoxExcelScreenUpdating.Visible = false;
                if (NBlocksCheckBox.Checked)
                {
                    NCProgramtToExtractNBlocksFromabel.Visible = true;
                    NCFileName_For_NBlocks_textbox.Visible = true;
                    NCFileBrowseButton.Visible = true;
                }
                else
                {
                    NCProgramtToExtractNBlocksFromabel.Visible = false;
                    NCFileName_For_NBlocks_textbox.Visible = false;
                    NCFileBrowseButton.Visible = false;
                }

                // Now see if we can enable the OK Button
                if (!NBlocksCheckBox.Checked)
                {
                    {
                        OKButton.Enabled = true;
                    }
                }
            }
        }

        private void SOLIDWORKS_CAM_Setup_Sheets_Load(object sender, EventArgs e)
        {
            CheckForOKButtonEnable();
        }

        private void CreateToolModelsButton_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            CAM_Setup_Sheets_Addin.CreateCAMWorksSolidTool();
            this.WindowState = FormWindowState.Normal;

            CAM_Setup_Sheets_Addin.iSwApp.ActivateDoc(CAM_Setup_Sheets_Addin._SWModelDoc.GetTitle());
        }

        private void ExcelTemplateWizard_button_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            Form frm = new ExcelTemplateWizardStart();
            frm.ShowDialog();
            this.WindowState = FormWindowState.Normal;
        }

        private void SOLIDWORKS_CAM_Setup_Sheets_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
        }
    }
}
