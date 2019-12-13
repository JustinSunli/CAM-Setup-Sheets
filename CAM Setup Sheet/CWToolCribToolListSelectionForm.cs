using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace CAM_Setup_Sheets
{
    public partial class CWToolCribToolListSelectionForm : Form
    {
        private int[] selectedtools;
        public int[] SelectedTools
        {
            get
            {
                return this.selectedtools;
            }
            set
            {
                this.selectedtools = value;
            }
        }

        private String folderpath;
        public String FolderPath
        {
            get
            {
                return this.folderpath;
            }
            set
            {
                this.folderpath = value;
            }
        }

        private String swfilepath;
        public String SWFilePath
        {
            get
            {
                return this.swfilepath;
            }
            set
            {
                this.swfilepath = value;
            }
        }

        private Boolean bdisablescreenupdating;
        public Boolean BDisableScreenUpdating
        {
            get
            {
                return this.bdisablescreenupdating;
            }
            set
            {
                this.bdisablescreenupdating = value;
            }
        }

        private Boolean bcreateassemblyoftools;
        public Boolean BCreateAssemblyOfTools
        {
            get
            {
                return this.bcreateassemblyoftools;
            }
            set
            {
                this.bcreateassemblyoftools = value;
            }
        }

        private double xdistancebetweentools;
        public double XDistanceBetweentTools
        {
            get
            {
                return this.xdistancebetweentools;
            }
            set
            {
                this.xdistancebetweentools = value;
            }
        }

        private Boolean bcreatestltools;
        public Boolean BCreateSTLTools
        {
            get
            {
                return this.bcreateassemblyoftools;
            }
            set
            {
                this.bcreateassemblyoftools = value;
            }
        }

        public CWToolCribToolListSelectionForm()
        {
            InitializeComponent();
            BindingList<CWTools> BindingToolList = new BindingList<CWTools>(CAM_Setup_Sheets_Addin.SolidToolList);
            dataGridViewCWToolCribTools.DataSource = BindingToolList;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            int numtools = 0;
            foreach (DataGridViewRow r in dataGridViewCWToolCribTools.SelectedRows)
            {
                numtools++;
            }
            SelectedTools = new int[numtools];
            int counter = 0;
            foreach (DataGridViewRow r in dataGridViewCWToolCribTools.SelectedRows)
            {
                SelectedTools[counter] = r.Cells[0].RowIndex;
                counter++;
            }

            Array.Sort(SelectedTools,
                       new Comparison<int>((i1, i2) => i1.CompareTo(i2)));
        }

        private void pictureBoxCWToolCuttingPortionColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                pictureBoxCWToolCuttingPortionColor.BackColor = colorDialog1.Color;
            }
        }

        private void pictureBoxCWToolShankColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                pictureBoxCWToolShankColor.BackColor = colorDialog1.Color;
            }
        }

        private void pictureBoxCWToolHolderColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                pictureBoxCWToolHolderColor.BackColor = colorDialog1.Color;
            }
        }

        private void dataGridViewCWToolCribTools_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int numtools = 0;
            foreach (DataGridViewRow r in dataGridViewCWToolCribTools.SelectedRows)
            {
                numtools++;
            }
            SelectedTools = new int[numtools];
            int counter = 0;
            foreach (DataGridViewRow r in dataGridViewCWToolCribTools.SelectedRows)
            {
                SelectedTools[counter] = r.Cells[0].RowIndex;
                counter++;
            }

            Array.Sort(SelectedTools,
                       new Comparison<int>((i1, i2) => i1.CompareTo(i2)));

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void FolderPathTextBox_TextChanged(object sender, EventArgs e)
        {
            FolderPath = FolderPathTextBox.Text;
        }

        private void buttonSelectFolder_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = SWFilePath;
            folderBrowserDialog1.ShowNewFolderButton = true;
            DialogResult result = folderBrowserDialog1.ShowDialog();
            FolderPath = folderBrowserDialog1.SelectedPath;
            FolderPathTextBox.Text = FolderPath;
        }

        private void CWToolCribToolListSelectionForm_Load(object sender, EventArgs e)
        {
            FolderPathTextBox.Text = FolderPath;
            XDistanceBetweentTools = (double)XDistanceBetweentTools;
        }

        private void CWToolCribToolListSelectionForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!System.IO.Directory.Exists(FolderPath))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(FolderPath);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Cannot create folder \"" + FolderPath + "\"." +
                                    ex.Message.ToString());
                    e.Cancel = true;
                }
            }
        }

        private void checkBoxDisableScreenUpdating_CheckedChanged(object sender, EventArgs e)
        {
            BDisableScreenUpdating = checkBoxDisableScreenUpdating.Checked;
        }

        private void checkBoxCreatAssyOfTools_CheckedChanged(object sender, EventArgs e)
        {
            BCreateAssemblyOfTools = checkBoxCreatAssyOfTools.Checked;
        }

        private void CreateToolAsSTL_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            BCreateSTLTools = CreateToolAsSTL_checkBox.Checked;
        }

        private void numericUpDownXDistanceBetweenTools_ValueChanged(object sender, EventArgs e)
        {
            XDistanceBetweentTools = (double)numericUpDownXDistanceBetweenTools.Value;
        }

        private void dataGridViewCWToolCribTools_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int numtools = 0;
            foreach (DataGridViewRow r in dataGridViewCWToolCribTools.SelectedRows)
            {
                numtools++;
            }
            if (numtools > 1)
            {
                checkBoxCreatAssyOfTools.Checked = true;
                checkBoxDisableScreenUpdating.Checked = true;
            }
            else
            {
                checkBoxCreatAssyOfTools.Checked = false;
                checkBoxDisableScreenUpdating.Checked = false;
            }
        }
    }
}
