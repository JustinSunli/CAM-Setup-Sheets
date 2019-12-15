using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Font = System.Drawing.Font;
using Syncfusion.Windows.Forms.Tools;

namespace CAM_Setup_Sheets
{
    public partial class ExcelTemplateWizardStart : Form
    {
        private Worksheet worksheet = null;

        // Current TextBox
        TextBoxExt CurrentTextBox = null;
        public ExcelTemplateWizardStart()
        {
            InitializeComponent();

            // Set Radio Buttons
            radioButtonMill.Checked = Properties.Settings.Default.SetupSheetTypeMill;
            radioButtonTurn.Checked = Properties.Settings.Default.SetupSheetTypeTurn;
            radioButtonMillTurn.Checked = Properties.Settings.Default.SetupSheetTypeMillTurn;

            // Set a region at top of tab control to block out tabs
            TabControl1.Region = new Region(new RectangleF(tabPage1.Left, tabPage1.Top,
                tabPage1.Width, tabPage1.Height));


            // Get Post Parameters Items for Header Items List
            listBoxPostParameters.Items.Clear();
            listBoxPostParametersTab4.Items.Clear();
            listBoxPostParametersTab5.Items.Clear();
            listBoxPostParametersTab6.Items.Clear();
            listBoxPostParametersTab7.Items.Clear();
            listBoxPostParametersTab8.Items.Clear();
            listBoxPostParametersTab9.Items.Clear();

            if (CAM_Setup_Sheets_Addin._PostParameterNames != null)
            {
                foreach (var item in CAM_Setup_Sheets_Addin._PostParameterNames)
                {
                    listBoxPostParameters.Items.Add(item);
                    listBoxPostParametersTab4.Items.Add(item);
                    listBoxPostParametersTab5.Items.Add(item);
                    listBoxPostParametersTab6.Items.Add(item);
                    listBoxPostParametersTab7.Items.Add(item);
                    listBoxPostParametersTab8.Items.Add(item);
                    listBoxPostParametersTab9.Items.Add(item);
                }
            }
        }

        private void SET_DateTimeNow()
        {
            DateTime dt = DateTime.Now; // Or whatever
            string s = "This was created on ";
            s += dt.ToString("MM/dd/yyyy");
            s += " at ";
            s += dt.ToString("h:mm:ss tt");
            for (int i = 0; i < CAM_Setup_Sheets_Addin._PostParameterNames.Count; i++)
            {
                if (CAM_Setup_Sheets_Addin._PostParameterNames[i] == "Date/Time Created")
                {
                    CAM_Setup_Sheets_Addin._PostParameterValues[i] = s;
                }
            }
        }
        private void ButtonRow1Color_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor1stHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor1stHeaderRow = col.Color;
                
            }
        }

        private void ButtonRow2Color_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor2ndHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor2ndHeaderRow = col.Color;
                
            }
        }

        private void ButtonRow3Color_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor3rdHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor3rdHeaderRow = col.Color;
                
            }
        }

        private void ButtonRow4ForeColor_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor4thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor4thHeaderRow = col.Color;
                
            }
        }

        private void ButtonRow5ForeColor_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor5thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor5thHeaderRow = col.Color;
                
            }
        }

        private void ButtonRow6ForeColor_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor6thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor6thHeaderRow = col.Color;
                
            }
        }

        private void ButtonRow7ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor7thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor7thHeaderRow = col.Color;
                
            }
        }

        private void ButtonRow8ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor8thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor8thHeaderRow = col.Color;
                
            }
        }

        private void ButtonRow9ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor9thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor9thHeaderRow = col.Color;
                
            }
        }

        private void ButtonRow10ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor10thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor10thHeaderRow = col.Color;
                
            }
        }
        private void ButtonRow11ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor11thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor11thHeaderRow = col.Color;
                
            }
        }
        private void ButtonRow12ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor12thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor12thHeaderRow = col.Color;
                
            }
        }
        private void Button_RH_Row4ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_4thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_4thHeaderRow = col.Color;
                
            }
        }
        private void Button_RH_Row5ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_5thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_5thHeaderRow = col.Color;
                
            }
        }
        private void Button_RH_Row6ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_6thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_6thHeaderRow = col.Color;
                
            }
        }
        private void Button_RH_Row7ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_7thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_7thHeaderRow = col.Color;
                
            }
        }
        private void Button_RH_Row8ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_8thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_8thHeaderRow = col.Color;
                
            }
        }
        private void Button_RH_Row9ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_9thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_9thHeaderRow = col.Color;
                
            }
        }
        private void Button_RH_Row10ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_10thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_10thHeaderRow = col.Color;
                
            }
        }
        private void Button_RH_Row11ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_11thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_11thHeaderRow = col.Color;
                
            }
        }
        private void Button_RH_Row12ForeColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_12thHeaderRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_12thHeaderRow = col.Color;
                
            }
        }
        private void ButtonRow4ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor4thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor4thRowParameters = col.Color;
                
            }
        }

        private void ButtonRow5ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor5thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor5thRowParameters = col.Color;
                
            }
        }

        private void ButtonRow6ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor6thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor6thRowParameters = col.Color;
                
            }
        }

        private void ButtonRow7ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor7thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor7thRowParameters = col.Color;
                
            }
        }

        private void ButtonRow8ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor8thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor8thRowParameters = col.Color;
                
            }
        }

        private void ButtonRow9ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor9thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor9thRowParameters = col.Color;
                
            }
        }

        private void ButtonRow10ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor10thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor10thRowParameters = col.Color;
                
            }
        }
        private void ButtonRow11ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor11thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor11thRowParameters = col.Color;
                
            }
        }
        private void ButtonRow12ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor12thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor12thRowParameters = col.Color;
                
            }
        }
        private void Button_RH_Row4ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_4thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_4thRowParameters = col.Color;
                
            }
        }
        private void Button_RH_Row5ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_5thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_5thRowParameters = col.Color;
                
            }
        }
        private void Button_RH_Row6ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_6thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_6thRowParameters = col.Color;
                
            }
        }
        private void Button_RH_Row7ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_7thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_7thRowParameters = col.Color;
                
            }
        }
        private void Button_RH_Row8ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_8thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_8thRowParameters = col.Color;
                
            }
        }
        private void Button_RH_Row9ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_9thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_9thRowParameters = col.Color;
                
            }
        }
        private void Button_RH_Row10ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_10thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_10thRowParameters = col.Color;
                
            }
        }
        private void Button_RH_Row11ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_11thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_11thRowParameters = col.Color;
                
            }
        }
        private void Button_RH_Row12ForeColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor_RH_12thRowParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor_RH_12thRowParameters = col.Color;
                
            }
        }
        private void ButtonRow1BackColor_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row1BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row1BackColor = col.Color;
                
            }
        }

        private void ButtonRow2BackColor_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row2BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row2BackColor = col.Color;
                
            }
        }

        private void ButtonRow3BackColor_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row3BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row3BackColor = col.Color;
                
            }
        }

        private void ButtonRow4BackColor_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row4BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row4BackColor = col.Color;
                
            }
        }

        private void ButtonRow5BackColor_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row5BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row5BackColor = col.Color;
                
            }
        }

        private void ButtonRow6BackColor_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row6BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row6BackColor = col.Color;
                
            }
        }

        private void ButtonRow7BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row7BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row7BackColor = col.Color;
                
            }
        }

        private void ButtonRow8BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row8BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row8BackColor = col.Color;
                
            }
        }

        private void ButtonRow9BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row9BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row9BackColor = col.Color;
                
            }
        }

        private void ButtonRow10BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row10BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row10BackColor = col.Color;
                
            }
        }
        private void ButtonRow11BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row11BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row11BackColor = col.Color;
                
            }
        }
        private void ButtonRow12BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row12BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row12BackColor = col.Color;
                
            }
        }
        private void Button_RH_Row4BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row4_RH_BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row4_RH_BackColor = col.Color;
                
            }
        }
        private void Button_RH_Row5BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row5_RH_BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row5_RH_BackColor = col.Color;
                
            }
        }
        private void Button_RH_Row6BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row6_RH_BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row6_RH_BackColor = col.Color;
                
            }
        }
        private void Button_RH_Row7BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row7_RH_BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row7_RH_BackColor = col.Color;
                
            }
        }
        private void Button_RH_Row8BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row8_RH_BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row8_RH_BackColor = col.Color;
                
            }
        }
        private void Button_RH_Row9BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row9_RH_BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row9_RH_BackColor = col.Color;
                
            }
        }
        private void Button_RH_Row10BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row10_RH_BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row10_RH_BackColor = col.Color;
                
            }
        }
        private void Button_RH_Row11BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row11_RH_BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row11_RH_BackColor = col.Color;
                
            }
        }
        private void Button_RH_Row12BackColorText_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row12_RH_BackColor;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row12_RH_BackColor = col.Color;
                
            }
        }
        private void ButtonRow4BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row4BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row4BackColorParameters = col.Color;
                
            }
        }

        private void ButtonRow5BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row5BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row5BackColorParameters = col.Color;
                
            }
        }

        private void ButtonRow6BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row6BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row6BackColorParameters = col.Color;
                
            }
        }

        private void ButtonRow7BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row7BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row7BackColorParameters = col.Color;
                
            }
        }

        private void ButtonRow8BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row8BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row8BackColorParameters = col.Color;
                
            }
        }

        private void ButtonRow9BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row9BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row9BackColorParameters = col.Color;
                
            }
        }

        private void ButtonRow10BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row10BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row10BackColorParameters = col.Color;
                
            }
        }
        private void ButtonRow11BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row11BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row11BackColorParameters = col.Color;
                
            }
        }
        private void ButtonRow12BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row12BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row12BackColorParameters = col.Color;
                
            }
        }
        private void Button_RH_Row4BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row4_RH_BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row4_RH_BackColorParameters = col.Color;
                
            }
        }
        private void Button_RH_Row5BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row5_RH_BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row5_RH_BackColorParameters = col.Color;
                
            }
        }
        private void Button_RH_Row6BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row6_RH_BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row6_RH_BackColorParameters = col.Color;
                
            }
        }
        private void Button_RH_Row7BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row7_RH_BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row7_RH_BackColorParameters = col.Color;
                
            }
        }
        private void Button_RH_Row8BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row8_RH_BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row8_RH_BackColorParameters = col.Color;
                
            }
        }
        private void Button_RH_Row9BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row9_RH_BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row9_RH_BackColorParameters = col.Color;
                
            }
        }
        private void Button_RH_Row10BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row10_RH_BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row10_RH_BackColorParameters = col.Color;
                
            }
        }
        private void Button_RH_Row11BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row11_RH_BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row11_RH_BackColorParameters = col.Color;
                
            }
        }
        private void Button_RH_Row12BackColorParameters_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.Row12_RH_BackColorParameters;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.Row12_RH_BackColorParameters = col.Color;
                
            }
        }
        private void ButtonRow1Font_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor1stHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontFor1stHeaderRow;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontFor1stHeaderRow = fd.Font;
            
        }

        private void ButtonRow2Font_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor2ndHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontFor2ndHeaderRow;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontFor2ndHeaderRow = fd.Font;
            
        }

        private void ButtonRow3Font_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor3rdHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontFor3rdHeaderRow;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontFor3rdHeaderRow = fd.Font;
            
        }

        private void ButtonLeftSideRow4Font_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor4thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow4LeftSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow4LeftSide = fd.Font;
            
        }

        private void ButtonLeftSideRow5Font_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor5thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow5LeftSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow5LeftSide = fd.Font;
            
        }

        private void ButtonLeftSideRow6Font_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor6thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow6LeftSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow6LeftSide = fd.Font;
            
        }

        private void ButtonLeftSideRow7FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor7thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow7LeftSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow7LeftSide = fd.Font;
            
        }

        private void ButtonLeftSideRow8FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor8thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow8LeftSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow8LeftSide = fd.Font;
            
        }

        private void ButtonLeftSideRow9FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor9thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow9LeftSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow9LeftSide = fd.Font;
            
        }

        private void ButtonLeftSideRow10FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor10thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow10LeftSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow10LeftSide = fd.Font;
            
        }
        private void ButtonLeftSideRow11FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor11thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow11LeftSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow11LeftSide = fd.Font;
            
        }
        private void ButtonLeftSideRow12FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor12thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow12LeftSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow12LeftSide = fd.Font;
            
        }
        private void ButtonRightSideRow4FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_4thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow4RightSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow4RightSide = fd.Font;
            
        }
        private void ButtonRightSideRow5FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_5thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow5RightSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow5RightSide = fd.Font;
            
        }
        private void ButtonRightSideRow6FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_6thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow6RightSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow6RightSide = fd.Font;
            
        }
        private void ButtonRightSideRow7FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_7thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow7RightSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow7RightSide = fd.Font;
            
        }
        private void ButtonRightSideRow8FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_8thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow8RightSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow8RightSide = fd.Font;
            
        }
        private void ButtonRightSideRow9FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_9thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow9RightSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow9RightSide = fd.Font;
            
        }
        private void buttonRightSideRow10FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_10thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow10RightSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow10RightSide = fd.Font;
            
        }
        private void buttonRightSideRow11FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_11thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow11RightSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow11RightSide = fd.Font;
            
        }
        private void buttonRightSideRow12FontText_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_12thHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontForRow12RightSide;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow12RightSide = fd.Font;
            
        }
        private void ButtonLeftSideRow4FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor4thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow4LeftSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow4LeftSideParameters = fd.Font;
            
        }

        private void ButtonLeftSideRow5FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor5thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow5LeftSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow5LeftSideParameters = fd.Font;
            
        }

        private void ButtonLeftSideRow6FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor6thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow6LeftSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow6LeftSideParameters = fd.Font;
            
        }

        private void ButtonLeftSideRow7FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor7thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow7LeftSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow7LeftSideParameters = fd.Font;
            
        }

        private void ButtonLeftSideRow8FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor8thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow8LeftSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow8LeftSideParameters = fd.Font;
            
        }

        private void ButtonLeftSideRow9FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor9thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow9LeftSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow9LeftSideParameters = fd.Font;
            
        }

        private void ButtonLeftSideRow10FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor10thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow10LeftSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow10LeftSideParameters = fd.Font;
            
        }
        private void ButtonLeftSideRow11FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor11thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow11LeftSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow11LeftSideParameters = fd.Font;
            
        }
        private void ButtonLeftSideRow12FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor12thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow12LeftSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow12LeftSideParameters = fd.Font;
            
        }
        private void ButtonRightSideRow4FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_4thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow4RightSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow4RightSideParameters = fd.Font;
            
        }
        private void ButtonRightSideRow5FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_5thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow5RightSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow5RightSideParameters = fd.Font;
            
        }
        private void ButtonRightSideRow6FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_6thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow6RightSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow6RightSideParameters = fd.Font;
            
        }
        private void ButtonRightSideRow7FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_7thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow7RightSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow7RightSideParameters = fd.Font;
            
        }
        private void ButtonRightSideRow8FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_8thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow8RightSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow8RightSideParameters = fd.Font;
            
        }
        private void ButtonRightSideRow9FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_9thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow9RightSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow9RightSideParameters = fd.Font;
            
        }

        private void buttonRightSideRow10FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_10thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow10RightSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow10RightSideParameters = fd.Font;
            
        }
        private void buttonRightSideRow11FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_11thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow11RightSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow11RightSideParameters = fd.Font;
            
        }
        private void buttonRightSideRow12FontParameters_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor_RH_12thRowParameters;

            fd.Font = Properties.Settings.Default.TextFontForRow12RightSideParameters;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForRow12RightSideParameters = fd.Font;
            
        }
        private void ANYTextBox_DragDrop(object sender, DragEventArgs e)
        {
            if (sender is Syncfusion.Windows.Forms.Tools.TextBoxExt)
            {
                var s = e.Data.GetData(DataFormats.StringFormat).ToString();
                System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)sender;
                tb.Text += "<" + s + ">";
            }
        }
        private void ANYTextBox_DragEnter(object sender, DragEventArgs e)
        {
            if (sender is Syncfusion.Windows.Forms.Tools.TextBoxExt)
            {
                e.Effect = DragDropEffects.All;
            }
        }

        private void ANY_ListBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (!(sender is Syncfusion.Windows.Forms.Tools.TextBoxExt))
            {
                if (CurrentTextBox != null)
                {
                    CurrentTextBox.BorderColor = Color.Black;
                    CurrentTextBox.CornerRadius = 0;
                }
                CurrentTextBox = null;
            }

            if (sender is System.Windows.Forms.ListBox)
            {
                System.Windows.Forms.ListBox lb = (System.Windows.Forms.ListBox)sender;
	            if (lb.Items.Count == 0) return;
	
	            var index = lb.IndexFromPoint(e.X, e.Y);
	            if (index != -1)
	            {
	                var s = lb.Items[index].ToString();
	                var dde1 = DoDragDrop(s, DragDropEffects.All);
	            }
            }
        }

        private void ListBoxPostParametersTab4_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxPostParametersTab4.Items.Count == 0) return;

            var index = listBoxPostParametersTab4.IndexFromPoint(e.X, e.Y);
            if (index != -1)
            {
                var s = listBoxPostParametersTab4.Items[index].ToString();
                var dde1 = DoDragDrop(s, DragDropEffects.All);
            }
        }

        private void ListBoxPostParametersTab5_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxPostParametersTab5.Items.Count == 0) return;

            var index = listBoxPostParametersTab5.IndexFromPoint(e.X, e.Y);
            if (index != -1)
            {
                var s = listBoxPostParametersTab5.Items[index].ToString();
                var dde1 = DoDragDrop(s, DragDropEffects.All);
            }
        }

        private void ListBoxPostParametersTab6_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxPostParametersTab6.Items.Count == 0) return;

            var index = listBoxPostParametersTab6.IndexFromPoint(e.X, e.Y);
            if (index != -1)
            {
                var s = listBoxPostParametersTab6.Items[index].ToString();
                var dde1 = DoDragDrop(s, DragDropEffects.All);
            }
        }
        private void ListBoxPostParametersTab7_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxPostParametersTab7.Items.Count == 0) return;

            var index = listBoxPostParametersTab7.IndexFromPoint(e.X, e.Y);
            if (index != -1)
            {
                var s = listBoxPostParametersTab7.Items[index].ToString();
                var dde1 = DoDragDrop(s, DragDropEffects.All);
            }
        }
        private void ListBoxPostParametersTab8_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxPostParametersTab8.Items.Count == 0) return;

            var index = listBoxPostParametersTab8.IndexFromPoint(e.X, e.Y);
            if (index != -1)
            {
                var s = listBoxPostParametersTab8.Items[index].ToString();
                var dde1 = DoDragDrop(s, DragDropEffects.All);
            }
        }
        private void listBoxPostParametersTab9_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxPostParametersTab9.Items.Count == 0) return;

            var index = listBoxPostParametersTab9.IndexFromPoint(e.X, e.Y);
            if (index != -1)
            {
                var s = listBoxPostParametersTab9.Items[index].ToString();
                var dde1 = DoDragDrop(s, DragDropEffects.All);
            }
        }
        private void SetExcelFont(Range rng1, Color forecolor, Color backcolor, Font font)
        {
            rng1.Font.Color = ColorTranslator.ToOle(forecolor);
            rng1.Interior.Color = ColorTranslator.ToOle(backcolor);
            rng1.Font.FontStyle = font.Style;
            rng1.Font.Bold = font.Bold;
            rng1.Font.Size = font.Size;
            rng1.Font.Italic = font.Italic;
            rng1.Font.Underline = font.Underline;
            rng1.Font.Strikethrough = font.Strikeout;
            rng1.Font.Name = font.Name;
        }

        private void Next1_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 1;
            if (Properties.Settings.Default.SetupSheetTypeMill)
            {
                labelMachineType3.Text = radioButtonMill.Text + " Operations List";
                labelMachineType4.Text = radioButtonMill.Text + " Operations List";
                labelMachineType5.Text = radioButtonMill.Text + " Operations List";
                labelMachineType6.Text = radioButtonMill.Text + " Operations List";
                labelMachineType7.Text = radioButtonMill.Text + " Operations List";
                labelMachineType8.Text = radioButtonMill.Text + " Operations List";
                labelMachineType9.Text = radioButtonMill.Text + " Operations List";
                CAM_Setup_Sheets_Addin._SetupSheetType = radioButtonMill.Text + " Operations List";
            }
            if (Properties.Settings.Default.SetupSheetTypeTurn)
            {
                labelMachineType3.Text = radioButtonTurn.Text + " Operations List";
                labelMachineType4.Text = radioButtonTurn.Text + " Operations List";
                labelMachineType5.Text = radioButtonTurn.Text + " Operations List";
                labelMachineType6.Text = radioButtonTurn.Text + " Operations List";
                labelMachineType7.Text = radioButtonTurn.Text + " Operations List";
                labelMachineType8.Text = radioButtonTurn.Text + " Operations List";
                labelMachineType9.Text = radioButtonTurn.Text + " Operations List";
                CAM_Setup_Sheets_Addin._SetupSheetType = radioButtonTurn.Text + " Operations List";
            }
            if (Properties.Settings.Default.SetupSheetTypeMillTurn)
            {
                labelMachineType3.Text = radioButtonMillTurn.Text + " Operations List";
                labelMachineType4.Text = radioButtonMillTurn.Text + " Operations List";
                labelMachineType5.Text = radioButtonMillTurn.Text + " Operations List";
                labelMachineType6.Text = radioButtonMillTurn.Text + " Operations List";
                labelMachineType7.Text = radioButtonMillTurn.Text + " Operations List";
                labelMachineType8.Text = radioButtonMillTurn.Text + " Operations List";
                labelMachineType9.Text = radioButtonMillTurn.Text + " Operations List";
                CAM_Setup_Sheets_Addin._SetupSheetType = radioButtonMillTurn.Text + " Operations List";
            }

        }

        private void Next2_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 2;
            
        }

        private void Next3_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 3;
            
        }

        private void Next4_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 4;
            
        }

        private void Next5_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 5;
            
        }

        private void Next6_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 6;
            
        }
        private void Next7_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 7;
            
        }
        private void Next8_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 8;
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 9;
        }
        private void GOTO_OPerations_Parameters(object sender, EventArgs e)
        {
            
            this.WindowState = FormWindowState.Minimized;
            ExcelTemplateWizard_OperationParameters OP_Params = new ExcelTemplateWizard_OperationParameters();       
            OP_Params.ShowDialog();
            this.WindowState = FormWindowState.Normal;

        }
        private void Back2_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 0;
        }

        private void Back3_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 1;
            
        }

        private void Back4_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 2;
            
        }

        private void Back5_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 3;
            
        }


        private void Back6_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 4;
            
        }
        private void Back7_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 5;
            
        }
        private void Back8_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 6;
            
        }
        private void Back9_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 7;
            
        }
        private void RadioButtonMill_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SetupSheetTypeMill = radioButtonMill.Checked;
            labelMachineType3.Text = radioButtonMill.Text + " Operations List";
            labelMachineType4.Text = radioButtonMill.Text + " Operations List";
            labelMachineType5.Text = radioButtonMill.Text + " Operations List";
            labelMachineType6.Text = radioButtonMill.Text + " Operations List";
            labelMachineType7.Text = radioButtonMill.Text + " Operations List";
            labelMachineType8.Text = radioButtonMill.Text + " Operations List";
            labelMachineType9.Text = radioButtonMill.Text + " Operations List";
            CAM_Setup_Sheets_Addin._SetupSheetType = radioButtonMill.Text + " Operations List";
        }

        private void RadioButtonTurn_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SetupSheetTypeTurn = radioButtonTurn.Checked;
            labelMachineType3.Text = radioButtonTurn.Text + " Operations List";
            labelMachineType4.Text = radioButtonTurn.Text + " Operations List";
            labelMachineType5.Text = radioButtonTurn.Text + " Operations List";
            labelMachineType6.Text = radioButtonTurn.Text + " Operations List";
            labelMachineType7.Text = radioButtonTurn.Text + " Operations List";
            labelMachineType8.Text = radioButtonTurn.Text + " Operations List";
            labelMachineType9.Text = radioButtonTurn.Text + " Operations List";
            CAM_Setup_Sheets_Addin._SetupSheetType = radioButtonTurn.Text + " Operations List";
        }

        private void RadioButtonMillTurn_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SetupSheetTypeMillTurn = radioButtonMillTurn.Checked;
            labelMachineType3.Text = radioButtonMillTurn.Text + " Operations List";
            labelMachineType4.Text = radioButtonMillTurn.Text + " Operations List";
            labelMachineType5.Text = radioButtonMillTurn.Text + " Operations List";
            labelMachineType6.Text = radioButtonMillTurn.Text + " Operations List";
            labelMachineType7.Text = radioButtonMillTurn.Text + " Operations List";
            labelMachineType8.Text = radioButtonMillTurn.Text + " Operations List";
            labelMachineType9.Text = radioButtonMillTurn.Text + " Operations List";
            CAM_Setup_Sheets_Addin._SetupSheetType = radioButtonMillTurn.Text + " Operations List";
        }

        private void Button_ResetAllValues_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Reset();
        }

        public void ButtonPreviewInExcelTab3_Click(object sender, EventArgs e)
        {
            

            // Set the date and time right now in PostParameter values
            SET_DateTimeNow();

            int NumberOfColumns = 15;

            CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = Properties.Settings.Default.OperationItemsToUse.Count;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate < 15)
            {
                CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = 15;
            }

            else
            {
                NumberOfColumns = CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate;
            }

            TopMost = false;
            //Create Excel COM Objects. Create a COM object for everything that is referenced
            var xlApp = new Application();

            xlApp.Visible = true;
            var xlWorkbook = xlApp.Workbooks.Add();
            var newWorksheet = (Worksheet)xlWorkbook.Worksheets.Add();
            worksheet = newWorksheet;
            if (radioButtonMill.Checked)
                newWorksheet.Name = "Mill Operations";
            if (radioButtonMillTurn.Checked)
                newWorksheet.Name = "MillTurn Operations";
            if (radioButtonTurn.Checked)
                newWorksheet.Name = "Turn Operations";

            newWorksheet.Activate();


            // Set top row text
            newWorksheet.Cells[1, 1].Value = OperationsHeaderRow1.Text;
            newWorksheet.Cells[2, 1].Value = OperationsHeaderRow2.Text;
            newWorksheet.Cells[3, 1].Value = OperationsHeaderRow3.Text;

            // Set Fonts and Colors
            var rng1 = newWorksheet.Range["A1", Type.Missing];
            SetExcelFont(rng1, Properties.Settings.Default.TextColorFor1stHeaderRow,
                Properties.Settings.Default.Row1BackColor,
                Properties.Settings.Default.TextFontFor1stHeaderRow);

            rng1 = newWorksheet.Range["A2", Type.Missing];
            SetExcelFont(rng1, Properties.Settings.Default.TextColorFor2ndHeaderRow,
                Properties.Settings.Default.Row2BackColor,
                Properties.Settings.Default.TextFontFor2ndHeaderRow);

            rng1 = newWorksheet.Range["A3", Type.Missing];
            SetExcelFont(rng1, Properties.Settings.Default.TextColorFor3rdHeaderRow,
                Properties.Settings.Default.Row3BackColor,
                Properties.Settings.Default.TextFontFor3rdHeaderRow);

            // If DO NOT USE This row checked on anything, delete the row.
            if (Properties.Settings.Default.DoNotUseRow1Checked)
            {
                rng1 = newWorksheet.Range["A1"];
                rng1.EntireRow.Clear();
            }

            if (Properties.Settings.Default.DoNotUseRow2Checked)
            {
                rng1 = newWorksheet.Range["A2"];
                rng1.EntireRow.Clear();
            }

            if (Properties.Settings.Default.DoNotUseRow3Checked)
            {
                rng1 = newWorksheet.Range["A3"];
                rng1.EntireRow.Clear();
            }


            // Populate the parameters with values
            var usedrange = newWorksheet.UsedRange;

            foreach (Range row in usedrange)
            {
                for (var i = 0; i < row.Columns.Count; i++)
                {
                    if (row.Cells[1, i + 1].Value2 != null)
                    {
                        String input = row.Cells[1, i + 1].Value2.ToString();

                        Regex regex = new Regex(@"\<.*?\>");

                        var arr = regex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();

                        for (int j = 0; j < arr.Length; j++)
                        {
                            String str = arr[j];
                            String clean = str.Replace("<", "");
                            clean = clean.Replace(">", "");

                            for (int k = 0; k < CAM_Setup_Sheets_Addin._PostParameterNames.Count; k++)
                            {
                                String param = CAM_Setup_Sheets_Addin._PostParameterNames[k];
                                if (clean == param)
                                {
                                    input = input.Replace(str,
                                        CAM_Setup_Sheets_Addin._PostParameterValues[k]);
                                }
                            }
                        }

                        row.Cells[1, i + 1].Value2 = input;
                    }
                }
            }
            // Merge and center if needed
            rng1 = newWorksheet.Range[newWorksheet.Cells[1, 1], newWorksheet.Cells[1, NumberOfColumns]];
            rng1.Merge();
            if (Properties.Settings.Default.Row1MergeAndCenterChecked)
            {
                rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            }

            else
            {
                rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            }

            // Add Border if checked
            if (Properties.Settings.Default.Row1AddBorder)
            {
                rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            }

            rng1 = newWorksheet.Range[newWorksheet.Cells[2, 1], newWorksheet.Cells[2, NumberOfColumns]];
            rng1.Merge();
            if (Properties.Settings.Default.Row2MergeAndCenterChecked)
            {
                rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            }

            else
            {
                rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            }

            // Add Border if checked
            if (Properties.Settings.Default.Row2AddBorder)
            {
                rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            }

            rng1 = newWorksheet.Range[newWorksheet.Cells[3, 1], newWorksheet.Cells[3, NumberOfColumns]];
            rng1.Merge();
            if (Properties.Settings.Default.Row3MergeAndCenterChecked)
            {
                rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            }

            else
            {
                rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            }

            // Add Border if checked
            if (Properties.Settings.Default.Row3AddBorder)
            {
                rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            }
            //// Uncomment these when done
            //xlWorkbook.Close();
            //xlApp.Quit();

            //Marshal.ReleaseComObject(newWorksheet);
            //Marshal.ReleaseComObject(xlWorkbook);
            //Marshal.ReleaseComObject(xlApp);


            //xlApp = null;
            //xlWorkbook = null;
            //newWorksheet = null;
        }

        private void ButtonPreviewInExcelTab4_Click(object sender, EventArgs e)
        {
            int NumberOfColumns = 15;

            CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = Properties.Settings.Default.OperationItemsToUse.Count;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate < 15)
            {
                CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = 15;
            }

            else
            {
                NumberOfColumns = CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate;
            }

            ButtonPreviewInExcelTab3_Click(this, null);


            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[4, 1].Value = Properties.Settings.Default.Row4Text;
                worksheet.Cells[5, 1].Value = Properties.Settings.Default.Row5Text;
                worksheet.Cells[6, 1].Value = Properties.Settings.Default.Row6Text;

                worksheet.Cells[4, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows3to6].Value = Properties.Settings.Default.Row4LeftSideTextParameters;
                worksheet.Cells[5, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows3to6].Value = Properties.Settings.Default.Row5LeftSideTextParameters;
                worksheet.Cells[6, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows3to6].Value = Properties.Settings.Default.Row6LeftSideTextParameters;

                // Set Fonts and Colors
                var rng1 = worksheet.Range["A4", Type.Missing];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor4thHeaderRow,
                    Properties.Settings.Default.Row4BackColor,
                    Properties.Settings.Default.TextFontForRow4LeftSide);

                rng1 = worksheet.Range["A5", Type.Missing];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor5thHeaderRow,
                    Properties.Settings.Default.Row5BackColor,
                    Properties.Settings.Default.TextFontForRow5LeftSide);

                rng1 = worksheet.Range["A6", Type.Missing];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor6thHeaderRow,
                    Properties.Settings.Default.Row6BackColor,
                    Properties.Settings.Default.TextFontForRow6LeftSide);

                // Parameter value side
                rng1 = worksheet.Range[worksheet.Cells[4, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12],
                                       worksheet.Cells[4, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor4thRowParameters,
                    Properties.Settings.Default.Row4BackColorParameters,
                    Properties.Settings.Default.TextFontForRow4LeftSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[5, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12],
                                       worksheet.Cells[5, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor5thRowParameters,
                    Properties.Settings.Default.Row5BackColorParameters,
                    Properties.Settings.Default.TextFontForRow5LeftSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[6, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12],
                                       worksheet.Cells[6, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor6thRowParameters,
                    Properties.Settings.Default.Row6BackColorParameters,
                    Properties.Settings.Default.TextFontForRow6LeftSideParameters);

                // If DO NOT USE This Row checked on anything, delete the row.
                if (Properties.Settings.Default.DoNotUseRow4Checked)
                {
                    rng1 = worksheet.Range["A4"];
                    rng1.EntireRow.Clear();
                }

                if (Properties.Settings.Default.DoNotUseRow5Checked)
                {
                    rng1 = worksheet.Range["A5"];
                    rng1.EntireRow.Clear();
                }

                if (Properties.Settings.Default.DoNotUseRow6Checked)
                {
                    rng1 = worksheet.Range["A6"];
                    rng1.EntireRow.Clear();
                }

                // Populate the parameters with values
                var usedrange = worksheet.UsedRange;

                foreach (Range row in usedrange)
                {
                    for (var i = 0; i < row.Columns.Count; i++)
                    {
                        if (row.Cells[1, i + 1].Value2 != null)
                        {
                            String input = row.Cells[1, i + 1].Value2.ToString();

                            // List<string> list = new List<string>(input.Split('<', '>'));

                            Regex regex = new Regex(@"\<.*?\>");
                            //Regex regex = new Regex(@"\[([^]]*)\]");
                            //MatchCollection matches = regex.Matches(input);
                            //int count = matches.Count;

                            var arr = regex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();

                            for (int j = 0; j < arr.Length; j++)
                            {
                                String str = arr[j];
                                String clean = str.Replace("<", "");
                                clean = clean.Replace(">", "");

                                for (int k = 0; k < CAM_Setup_Sheets_Addin._PostParameterNames.Count; k++)
                                {
                                    String param = CAM_Setup_Sheets_Addin._PostParameterNames[k];
                                    if (clean == param)
                                    {
                                        input = input.Replace(str,
                                            CAM_Setup_Sheets_Addin._PostParameterValues[k].ToString());
                                    }
                                }
                            }

                            row.Cells[1, i + 1].Value2 = input;
                        }
                    }
                }
                // Merge and center if needed
                if (Properties.Settings.Default.NumberOfColumnsToMergeRows3to6 > 0)
                {
                    int numtomerge = (int)Properties.Settings.Default.NumberOfColumnsToMergeRows3to6;

                    rng1 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[4, 1 + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row4MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }

                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row4_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[5, 1], worksheet.Cells[5, 1 + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row5MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }

                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row5_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[6, 1], worksheet.Cells[6, 1 + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row6MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }

                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row6_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    // Merge Parameter Value Cells
                    rng1 = worksheet.Range[worksheet.Cells[4, 2 + numtomerge], worksheet.Cells[4, NumberOfColumns / 2]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row4MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row4_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[5, 2 + numtomerge], worksheet.Cells[5, NumberOfColumns / 2]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row5MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }

                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row5_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[6, 2 + numtomerge], worksheet.Cells[6, NumberOfColumns / 2]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row6MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row6_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }
                }
            }
        }

        private void ButtonPreviewInExcelTab5_Click(object sender, EventArgs e)
        {
            int NumberOfColumns = 15;

            CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = Properties.Settings.Default.OperationItemsToUse.Count;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate < 15)
            {
                CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = 15;
            }

            else
            {
                NumberOfColumns = CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate;
            }

            ButtonPreviewInExcelTab4_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[7, 1].Value = Properties.Settings.Default.Row7Text;
                worksheet.Cells[8, 1].Value = Properties.Settings.Default.Row8Text;
                worksheet.Cells[9, 1].Value = Properties.Settings.Default.Row9Text;

                worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows7to9].Value =
                    Properties.Settings.Default.Row7LeftSideTextParameters;
                worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows7to9].Value =
                    Properties.Settings.Default.Row8LeftSideTextParameters;
                worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows7to9].Value =
                    Properties.Settings.Default.Row9LeftSideTextParameters;

                // Set Fonts and Colors
                var rng1 = worksheet.Range["A7", Type.Missing];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor7thHeaderRow,
                    Properties.Settings.Default.Row7BackColor,
                    Properties.Settings.Default.TextFontForRow7LeftSide);

                rng1 = worksheet.Range["A8", Type.Missing];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor8thHeaderRow,
                    Properties.Settings.Default.Row8BackColor,
                    Properties.Settings.Default.TextFontForRow8LeftSide);

                rng1 = worksheet.Range["A9", Type.Missing];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor9thHeaderRow,
                    Properties.Settings.Default.Row9BackColor,
                    Properties.Settings.Default.TextFontForRow9LeftSide);

                // Parameter value side
                rng1 = worksheet.Range[worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12],
                                       worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor7thRowParameters,
                    Properties.Settings.Default.Row7BackColorParameters,
                    Properties.Settings.Default.TextFontForRow7LeftSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12],
                                       worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor8thRowParameters,
                    Properties.Settings.Default.Row8BackColorParameters,
                    Properties.Settings.Default.TextFontForRow8LeftSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12],
                                       worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor9thRowParameters,
                    Properties.Settings.Default.Row9BackColorParameters,
                    Properties.Settings.Default.TextFontForRow9LeftSideParameters);



                // If DO NOT USE This Row checked on anything, delete the row.
                if (Properties.Settings.Default.DoNotUseRow7Checked)
                {
                    rng1 = worksheet.Range["A7"];
                    rng1.EntireRow.Clear();
                }

                if (Properties.Settings.Default.DoNotUseRow8Checked)
                {
                    rng1 = worksheet.Range["A8"];
                    rng1.EntireRow.Clear();
                }

                if (Properties.Settings.Default.DoNotUseRow9Checked)
                {
                    rng1 = worksheet.Range["A9"];
                    rng1.EntireRow.Clear();
                }

                // Populate the parameters with values
                var usedrange = worksheet.UsedRange;

                foreach (Range row in usedrange)
                {
                    for (var i = 0; i < row.Columns.Count; i++)
                    {
                        if (row.Cells[1, i + 1].Value2 != null)
                        {
                            String input = row.Cells[1, i + 1].Value2.ToString();

                            // List<string> list = new List<string>(input.Split('<', '>'));

                            Regex regex = new Regex(@"\<.*?\>");
                            //Regex regex = new Regex(@"\[([^]]*)\]");
                            //MatchCollection matches = regex.Matches(input);
                            //int count = matches.Count;

                            var arr = regex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();

                            for (int j = 0; j < arr.Length; j++)
                            {
                                String str = arr[j];
                                String clean = str.Replace("<", "");
                                clean = clean.Replace(">", "");

                                for (int k = 0; k < CAM_Setup_Sheets_Addin._PostParameterNames.Count; k++)
                                {
                                    String param = CAM_Setup_Sheets_Addin._PostParameterNames[k];
                                    if (clean == param)
                                    {
                                        input = input.Replace(str,
                                            CAM_Setup_Sheets_Addin._PostParameterValues[k].ToString());
                                    }
                                }
                            }

                            row.Cells[1, i + 1].Value2 = input;
                        }
                    }
                }
                // Merge and center if needed
                if (Properties.Settings.Default.NumberOfColumnsToMergeRows7to9 > 0)
                {
                    int numtomerge = (int)Properties.Settings.Default.NumberOfColumnsToMergeRows7to9;

                    rng1 = worksheet.Range[worksheet.Cells[7, 1], worksheet.Cells[7, 1 + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row7MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row7_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[8, 1], worksheet.Cells[8, 1 + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row8MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row8_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[9, 1], worksheet.Cells[9, 1 + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row9MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row9_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    // Merge Parameter Value Cells
                    rng1 = worksheet.Range[worksheet.Cells[7, 2 + numtomerge], worksheet.Cells[7, NumberOfColumns / 2]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row7MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row7_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[8, 2 + numtomerge], worksheet.Cells[8, NumberOfColumns / 2]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row8MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row8_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[9, 2 + numtomerge], worksheet.Cells[9, NumberOfColumns / 2]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row9MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row9_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }
                }
            }
        }


        private void ButtonPreviewInExcelTab6_Click(object sender, EventArgs e)
        {
            int NumberOfColumns = 15;

            CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = Properties.Settings.Default.OperationItemsToUse.Count;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate < 15)
            {
                CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = 15;
            }

            else
            {
                NumberOfColumns = CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate;
            }

            ButtonPreviewInExcelTab5_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[10, 1].Value = Properties.Settings.Default.Row10Text;
                worksheet.Cells[11, 1].Value = Properties.Settings.Default.Row11Text;
                worksheet.Cells[12, 1].Value = Properties.Settings.Default.Row12Text;

                worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12].Value =
                    Properties.Settings.Default.Row10LeftSideTextParameters;
                worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12].Value =
                    Properties.Settings.Default.Row11LeftSideTextParameters;
                worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12].Value =
                    Properties.Settings.Default.Row12LeftSideTextParameters;

                // Set Fonts and Colors
                var rng1 = worksheet.Range["A10", Type.Missing];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor10thHeaderRow,
                    Properties.Settings.Default.Row10BackColor,
                    Properties.Settings.Default.TextFontForRow10LeftSide);

                rng1 = worksheet.Range["A11", Type.Missing];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor11thHeaderRow,
                    Properties.Settings.Default.Row11BackColor,
                    Properties.Settings.Default.TextFontForRow11LeftSide);

                rng1 = worksheet.Range["A12", Type.Missing];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor12thHeaderRow,
                    Properties.Settings.Default.Row12BackColor,
                    Properties.Settings.Default.TextFontForRow12LeftSide);

                // Parameter value side
                rng1 = worksheet.Range[worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12],
                                       worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor10thRowParameters,
                    Properties.Settings.Default.Row10BackColorParameters,
                    Properties.Settings.Default.TextFontForRow10LeftSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12],
                                       worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor11thRowParameters,
                    Properties.Settings.Default.Row11BackColorParameters,
                    Properties.Settings.Default.TextFontForRow11LeftSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12],
                                       worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor12thRowParameters,
                    Properties.Settings.Default.Row12BackColorParameters,
                    Properties.Settings.Default.TextFontForRow12LeftSideParameters);



                // If DO NOT USE This Row checked on anything, delete the row.
                if (Properties.Settings.Default.DoNotUseRow10Checked)
                {
                    rng1 = worksheet.Range["A10"];
                    rng1.EntireRow.Clear();
                }

                if (Properties.Settings.Default.DoNotUseRow11Checked)
                {
                    rng1 = worksheet.Range["A11"];
                    rng1.EntireRow.Clear();
                }

                if (Properties.Settings.Default.DoNotUseRow12Checked)
                {
                    rng1 = worksheet.Range["A12"];
                    rng1.EntireRow.Clear();
                }

                // Populate the parameters with values
                var usedrange = worksheet.UsedRange;

                foreach (Range row in usedrange)
                {
                    for (var i = 0; i < row.Columns.Count; i++)
                    {
                        if (row.Cells[1, i + 1].Value2 != null)
                        {
                            String input = row.Cells[1, i + 1].Value2.ToString();

                            // List<string> list = new List<string>(input.Split('<', '>'));

                            Regex regex = new Regex(@"\<.*?\>");
                            //Regex regex = new Regex(@"\[([^]]*)\]");
                            //MatchCollection matches = regex.Matches(input);
                            //int count = matches.Count;

                            var arr = regex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();

                            for (int j = 0; j < arr.Length; j++)
                            {
                                String str = arr[j];
                                String clean = str.Replace("<", "");
                                clean = clean.Replace(">", "");

                                for (int k = 0; k < CAM_Setup_Sheets_Addin._PostParameterNames.Count; k++)
                                {
                                    String param = CAM_Setup_Sheets_Addin._PostParameterNames[k];
                                    if (clean == param)
                                    {
                                        input = input.Replace(str,
                                            CAM_Setup_Sheets_Addin._PostParameterValues[k].ToString());
                                    }
                                }
                            }

                            row.Cells[1, i + 1].Value2 = input;
                        }
                    }
                }
                // Merge and center if needed
                if (Properties.Settings.Default.NumberOfColumnsToMergeRows10to12 > 0)
                {
                    int numtomerge = (int)Properties.Settings.Default.NumberOfColumnsToMergeRows10to12;

                    rng1 = worksheet.Range[worksheet.Cells[10, 1], worksheet.Cells[10, 1 + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row10MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row10_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[11, 1], worksheet.Cells[11, 1 + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row11MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row11_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[12, 1], worksheet.Cells[12, 1 + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row12MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row12_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    // Merge Parameter Value Cells
                    rng1 = worksheet.Range[worksheet.Cells[10, 2 + numtomerge], worksheet.Cells[10, NumberOfColumns / 2]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row10MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row10_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[11, 2 + numtomerge], worksheet.Cells[11, NumberOfColumns / 2]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row11MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row11_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

                    rng1 = worksheet.Range[worksheet.Cells[12, 2 + numtomerge], worksheet.Cells[12, NumberOfColumns / 2]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row12MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row12_LH_AddBorder)
                    {
                        rng1.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }
                }
            }
        }


        private void ButtonPreviewInExcelTab7_Click(object sender, EventArgs e)
        {
            int NumberOfColumns = 15;

            CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = Properties.Settings.Default.OperationItemsToUse.Count;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate < 15)
            {
                CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = 15;
            }

            else
            {
                NumberOfColumns = CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate;
            }

            ButtonPreviewInExcelTab6_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[4, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row4_RH_Text;
                worksheet.Cells[5, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row5_RH_Text;
                worksheet.Cells[6, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row6_RH_Text;

                worksheet.Cells[4, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)].Value =
                    Properties.Settings.Default.Row4_RH_TextParameters;
                worksheet.Cells[5, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)].Value =
                    Properties.Settings.Default.Row5_RH_TextParameters;
                worksheet.Cells[6, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)].Value =
                    Properties.Settings.Default.Row6_RH_TextParameters;

                // Set Fonts and Colors
                var rng1 = worksheet.Range[worksheet.Cells[4, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[4, 1 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_4thHeaderRow,
                    Properties.Settings.Default.Row4_RH_BackColor,
                    Properties.Settings.Default.TextFontForRow4RightSide);

                rng1 = worksheet.Range[worksheet.Cells[5, 1 + (NumberOfColumns / 2)],
                                       worksheet.Cells[5, 1 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_5thHeaderRow,
                    Properties.Settings.Default.Row5_RH_BackColor,
                    Properties.Settings.Default.TextFontForRow5RightSide);

                rng1 = worksheet.Range[worksheet.Cells[6, 1 + (NumberOfColumns / 2)],
                                       worksheet.Cells[6, 1 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_6thHeaderRow,
                    Properties.Settings.Default.Row6_RH_BackColor,
                    Properties.Settings.Default.TextFontForRow6RightSide);

                // Parameter value side
                rng1 = worksheet.Range[worksheet.Cells[4, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)],
                                       worksheet.Cells[4, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_4thRowParameters,
                    Properties.Settings.Default.Row4_RH_BackColorParameters,
                    Properties.Settings.Default.TextFontForRow4RightSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[5, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)],
                                       worksheet.Cells[5, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_5thRowParameters,
                    Properties.Settings.Default.Row5_RH_BackColorParameters,
                    Properties.Settings.Default.TextFontForRow5RightSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[6, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)],
                                       worksheet.Cells[6, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_6thRowParameters,
                    Properties.Settings.Default.Row6_RH_BackColorParameters,
                    Properties.Settings.Default.TextFontForRow6RightSideParameters);

                // If DO NOT USE This Row checked on anything, delete the row.
                if (Properties.Settings.Default.DoNotUse_RH_Row4Checked)
                {
                    rng1 = worksheet.Range[worksheet.Cells[4, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[4, NumberOfColumns]];
                    rng1.Clear();
                }

                if (Properties.Settings.Default.DoNotUse_RH_Row5Checked)
                {
                    rng1 = worksheet.Range[worksheet.Cells[5, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[5, NumberOfColumns]];
                    rng1.Clear();
                }

                if (Properties.Settings.Default.DoNotUse_RH_Row6Checked)
                {
                    rng1 = worksheet.Range[worksheet.Cells[6, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[6, NumberOfColumns]];
                    rng1.Clear();
                }
                // Populate the parameters with values
                var usedrange = worksheet.UsedRange;

                foreach (Range row in usedrange)
                {
                    for (var i = 0; i < row.Columns.Count; i++)
                    {
                        if (row.Cells[1, i + 1].Value2 != null)
                        {
                            String input = row.Cells[1, i + 1].Value2.ToString();

                            // List<string> list = new List<string>(input.Split('<', '>'));

                            Regex regex = new Regex(@"\<.*?\>");
                            //Regex regex = new Regex(@"\[([^]]*)\]");
                            //MatchCollection matches = regex.Matches(input);
                            //int count = matches.Count;

                            var arr = regex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();

                            for (int j = 0; j < arr.Length; j++)
                            {
                                String str = arr[j];
                                String clean = str.Replace("<", "");
                                clean = clean.Replace(">", "");

                                for (int k = 0; k < CAM_Setup_Sheets_Addin._PostParameterNames.Count; k++)
                                {
                                    String param = CAM_Setup_Sheets_Addin._PostParameterNames[k];
                                    if (clean == param)
                                    {
                                        input = input.Replace(str,
                                            CAM_Setup_Sheets_Addin._PostParameterValues[k].ToString());
                                    }
                                }
                            }

                            row.Cells[1, i + 1].Value2 = input;
                        }
                    }
                }
                // Merge and center if needed
                if (Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 > 0)
                {
                    int numtomerge = (int)Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6;

                    rng1 = worksheet.Range[worksheet.Cells[4, 1 + (NumberOfColumns / 2)], worksheet.Cells[4, 1 + (NumberOfColumns / 2) + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row4_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row4_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[5, 1 + (NumberOfColumns / 2)], worksheet.Cells[5, 1 + (NumberOfColumns / 2) + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row5_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }

                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row5_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[6, 1 + (NumberOfColumns / 2)], worksheet.Cells[6, 1 + (NumberOfColumns / 2) + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row5_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row6_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    // Merge Parameter Value Cells
                    rng1 = worksheet.Range[worksheet.Cells[4, 2 + (NumberOfColumns / 2) + numtomerge],
                                                           worksheet.Cells[4, NumberOfColumns]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row4_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row4_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[5, 2 + (NumberOfColumns / 2) + numtomerge],
                                                           worksheet.Cells[5, NumberOfColumns]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row5_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row5_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[6, 2 + (NumberOfColumns / 2) + numtomerge],
                                                           worksheet.Cells[6, NumberOfColumns]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row6_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row6_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;

                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }
                }
            }
        }

        private void buttonPreviewInExcelTab8_Click(object sender, EventArgs e)
        {
            int NumberOfColumns = 15;

            CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = Properties.Settings.Default.OperationItemsToUse.Count;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate < 15)
            {
                CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = 15;
            }

            else
            {
                NumberOfColumns = CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate;
            }

            ButtonPreviewInExcelTab7_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[7, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row7_RH_Text;
                worksheet.Cells[8, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row8_RH_Text;
                worksheet.Cells[9, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row9_RH_Text;

                worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)].Value =
                    Properties.Settings.Default.Row7_RH_TextParameters;
                worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)].Value =
                    Properties.Settings.Default.Row8_RH_TextParameters;
                worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)].Value =
                    Properties.Settings.Default.Row9_RH_TextParameters;

                // Set Fonts and Colors
                var rng1 = worksheet.Range[worksheet.Cells[7, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[7, 1 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_7thHeaderRow,
                    Properties.Settings.Default.Row7_RH_BackColor,
                    Properties.Settings.Default.TextFontForRow7RightSide);

                rng1 = worksheet.Range[worksheet.Cells[8, 1 + (NumberOfColumns / 2)],
                                       worksheet.Cells[8, 1 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_8thHeaderRow,
                    Properties.Settings.Default.Row8_RH_BackColor,
                    Properties.Settings.Default.TextFontForRow8RightSide);

                rng1 = worksheet.Range[worksheet.Cells[9, 1 + (NumberOfColumns / 2)],
                                       worksheet.Cells[9, 1 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_9thHeaderRow,
                    Properties.Settings.Default.Row9_RH_BackColor,
                    Properties.Settings.Default.TextFontForRow9RightSide);

                // Parameter value side
                rng1 = worksheet.Range[worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)],
                                       worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_7thRowParameters,
                    Properties.Settings.Default.Row7_RH_BackColorParameters,
                    Properties.Settings.Default.TextFontForRow7RightSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)],
                                       worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_8thRowParameters,
                    Properties.Settings.Default.Row8_RH_BackColorParameters,
                    Properties.Settings.Default.TextFontForRow8RightSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)],
                                       worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_9thRowParameters,
                    Properties.Settings.Default.Row9_RH_BackColorParameters,
                    Properties.Settings.Default.TextFontForRow9RightSideParameters);

                // If DO NOT USE This Row checked on anything, delete the row.
                if (Properties.Settings.Default.DoNotUse_RH_Row7Checked)
                {
                    rng1 = worksheet.Range[worksheet.Cells[7, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[7, NumberOfColumns]];
                    rng1.Clear();
                }

                if (Properties.Settings.Default.DoNotUse_RH_Row8Checked)
                {
                    rng1 = worksheet.Range[worksheet.Cells[8, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[8, NumberOfColumns]];
                    rng1.Clear();
                }

                if (Properties.Settings.Default.DoNotUse_RH_Row9Checked)
                {
                    rng1 = worksheet.Range[worksheet.Cells[9, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[9, NumberOfColumns]];
                    rng1.Clear();
                }

                // Populate the parameters with values
                var usedrange = worksheet.UsedRange;

                foreach (Range row in usedrange)
                {
                    for (var i = 0; i < row.Columns.Count; i++)
                    {
                        if (row.Cells[1, i + 1].Value2 != null)
                        {
                            String input = row.Cells[1, i + 1].Value2.ToString();

                            Regex regex = new Regex(@"\<.*?\>");

                            var arr = regex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();

                            for (int j = 0; j < arr.Length; j++)
                            {
                                String str = arr[j];
                                String clean = str.Replace("<", "");
                                clean = clean.Replace(">", "");

                                for (int k = 0; k < CAM_Setup_Sheets_Addin._PostParameterNames.Count; k++)
                                {
                                    String param = CAM_Setup_Sheets_Addin._PostParameterNames[k];
                                    if (clean == param)
                                    {
                                        input = input.Replace(str,
                                            CAM_Setup_Sheets_Addin._PostParameterValues[k].ToString());
                                    }
                                }
                            }

                            row.Cells[1, i + 1].Value2 = input;
                        }
                    }
                }

                // Merge and center if needed
                if (Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 > 0)
                {
                    int numtomerge = (int)Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9;

                    rng1 = worksheet.Range[worksheet.Cells[7, 1 + (NumberOfColumns / 2)], worksheet.Cells[7, 1 + (NumberOfColumns / 2) + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row7_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row7_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[8, 1 + (NumberOfColumns / 2)], worksheet.Cells[8, 1 + (NumberOfColumns / 2) + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row8_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }

                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row8_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[9, 1 + (NumberOfColumns / 2)], worksheet.Cells[9, 1 + (NumberOfColumns / 2) + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row9_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row9_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    // Merge Parameter Value Cells
                    rng1 = worksheet.Range[worksheet.Cells[7, 2 + (NumberOfColumns / 2) + numtomerge],
                                                           worksheet.Cells[7, NumberOfColumns]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row7_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row7_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[8, 2 + (NumberOfColumns / 2) + numtomerge],
                                                           worksheet.Cells[8, NumberOfColumns]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row8_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row8_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[9, 2 + (NumberOfColumns / 2) + numtomerge],
                                                           worksheet.Cells[9, NumberOfColumns]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row9_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row9_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;

                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }
                }
            }

        }

        private void buttonPreviewInExcelTab9_Click(object sender, EventArgs e)
        {
            int NumberOfColumns = 15;

            CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = Properties.Settings.Default.OperationItemsToUse.Count;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate < 15)
            {
                CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = 15;
            }

            else
            {
                NumberOfColumns = CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate;
            }

            buttonPreviewInExcelTab8_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[10, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row10_RH_Text;
                worksheet.Cells[11, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row11_RH_Text;
                worksheet.Cells[12, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row12_RH_Text;

                worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)].Value =
                    Properties.Settings.Default.Row10_RH_TextParameters;
                worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)].Value =
                    Properties.Settings.Default.Row11_RH_TextParameters;
                worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)].Value =
                    Properties.Settings.Default.Row12_RH_TextParameters;

                // Set Fonts and Colors
                var rng1 = worksheet.Range[worksheet.Cells[10, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[10, 1 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_10thHeaderRow,
                    Properties.Settings.Default.Row10_RH_BackColor,
                    Properties.Settings.Default.TextFontForRow10RightSide);

                rng1 = worksheet.Range[worksheet.Cells[11, 1 + (NumberOfColumns / 2)],
                                       worksheet.Cells[11, 1 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_11thHeaderRow,
                    Properties.Settings.Default.Row11_RH_BackColor,
                    Properties.Settings.Default.TextFontForRow11RightSide);

                rng1 = worksheet.Range[worksheet.Cells[12, 1 + (NumberOfColumns / 2)],
                                       worksheet.Cells[12, 1 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_12thHeaderRow,
                    Properties.Settings.Default.Row12_RH_BackColor,
                    Properties.Settings.Default.TextFontForRow12RightSide);

                // Parameter value side
                rng1 = worksheet.Range[worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)],
                                       worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_10thRowParameters,
                    Properties.Settings.Default.Row10_RH_BackColorParameters,
                    Properties.Settings.Default.TextFontForRow10RightSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)],
                                       worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_11thRowParameters,
                    Properties.Settings.Default.Row11_RH_BackColorParameters,
                    Properties.Settings.Default.TextFontForRow11RightSideParameters);

                rng1 = worksheet.Range[worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)],
                                       worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)]];
                SetExcelFont(rng1, Properties.Settings.Default.TextColorFor_RH_12thRowParameters,
                    Properties.Settings.Default.Row12_RH_BackColorParameters,
                    Properties.Settings.Default.TextFontForRow12RightSideParameters);

                // If DO NOT USE This Row checked on anything, delete the row.
                if (Properties.Settings.Default.DoNotUse_RH_Row10Checked)
                {
                    rng1 = worksheet.Range[worksheet.Cells[10, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[10, NumberOfColumns]];
                    rng1.Clear();
                }

                if (Properties.Settings.Default.DoNotUse_RH_Row11Checked)
                {
                    rng1 = worksheet.Range[worksheet.Cells[11, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[11, NumberOfColumns]];
                    rng1.Clear();
                }

                if (Properties.Settings.Default.DoNotUse_RH_Row12Checked)
                {
                    rng1 = worksheet.Range[worksheet.Cells[12, 1 + (NumberOfColumns / 2)],
                                           worksheet.Cells[12, NumberOfColumns]];
                    rng1.Clear();
                }

                // Populate the parameters with values
                var usedrange = worksheet.UsedRange;

                foreach (Range row in usedrange)
                {
                    for (var i = 0; i < row.Columns.Count; i++)
                    {
                        if (row.Cells[1, i + 1].Value2 != null)
                        {
                            String input = row.Cells[1, i + 1].Value2.ToString();

                            Regex regex = new Regex(@"\<.*?\>");

                            var arr = regex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();

                            for (int j = 0; j < arr.Length; j++)
                            {
                                String str = arr[j];
                                String clean = str.Replace("<", "");
                                clean = clean.Replace(">", "");

                                for (int k = 0; k < CAM_Setup_Sheets_Addin._PostParameterNames.Count; k++)
                                {
                                    String param = CAM_Setup_Sheets_Addin._PostParameterNames[k];
                                    if (clean == param)
                                    {
                                        input = input.Replace(str,
                                            CAM_Setup_Sheets_Addin._PostParameterValues[k].ToString());
                                    }
                                }
                            }

                            row.Cells[1, i + 1].Value2 = input;
                        }
                    }
                }

                // Merge and center if needed
                if (Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 > 0)
                {
                    int numtomerge = (int)Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12;

                    rng1 = worksheet.Range[worksheet.Cells[10, 1 + (NumberOfColumns / 2)], worksheet.Cells[10, 1 + (NumberOfColumns / 2) + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row10_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row10_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[11, 1 + (NumberOfColumns / 2)], worksheet.Cells[11, 1 + (NumberOfColumns / 2) + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row11_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }

                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row11_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[12, 1 + (NumberOfColumns / 2)], worksheet.Cells[12, 1 + (NumberOfColumns / 2) + numtomerge]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row12_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row12_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    // Merge Parameter Value Cells
                    rng1 = worksheet.Range[worksheet.Cells[10, 2 + (NumberOfColumns / 2) + numtomerge],
                                                           worksheet.Cells[10, NumberOfColumns]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row10_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row10_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[11, 2 + (NumberOfColumns / 2) + numtomerge],
                                                           worksheet.Cells[11, NumberOfColumns]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row11_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row11_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;
                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }

                    rng1 = worksheet.Range[worksheet.Cells[12, 2 + (NumberOfColumns / 2) + numtomerge],
                                                           worksheet.Cells[12, NumberOfColumns]];
                    rng1.Merge();
                    if (Properties.Settings.Default.Row12_RH_MergeAndCenterChecked)
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rng1.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }

                    // Add border if checked
                    if (Properties.Settings.Default.Row12_RH_AddBorder)
                    {
                        Borders border = rng1.Borders;

                        border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    }
                }
            }
        }

        private void SetButtonStatesFromTextBox()
        {
            // Bold Button
            if (CurrentTextBox.Font.Bold)
            {
                BoldButton.Checked = true;
            }

            else
            {
                BoldButton.Checked = false;
            }

            // Italic Button
            if (CurrentTextBox.Font.Italic)
            {
                ItalicButton.Checked = true;
            }

            else
            {
                ItalicButton.Checked = false;
            }

            // Underline Button
            if (CurrentTextBox.Font.Underline)
            {
                UnderlineButton.Checked = true;
            }

            else
            {
                UnderlineButton.Checked = false;
            }
        }

        private void ANY_TextBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (sender is Syncfusion.Windows.Forms.Tools.TextBoxExt)
            {
                if (CurrentTextBox != null)
                {
                    CurrentTextBox.BorderColor = Color.Black;
                    CurrentTextBox.CornerRadius = 0;
                }
                CurrentTextBox = (Syncfusion.Windows.Forms.Tools.TextBoxExt)sender;
                CurrentTextBox.DeselectAll();
                CurrentTextBox.BorderColor = Color.Red;
                CurrentTextBox.CornerRadius = CurrentTextBox.Height/2;
                SetButtonStatesFromTextBox();

            }
        }

        private void BoldButton_Click(object sender, EventArgs e)
        {
            if(BoldButton.Checked)
            {
                BoldButton.Checked = false;
            }

            else
            {
                BoldButton.Checked = true;
            }

            if (CurrentTextBox != null)
            {
                CurrentTextBox.Font = new Font(CurrentTextBox.Font, FontStyle.Bold ^ CurrentTextBox.Font.Style); ;
            }
        }

        private void ItalicButton_Click(object sender, EventArgs e)
        {
            if (ItalicButton.Checked)
            {
                ItalicButton.Checked = false;
            }

            else
            {
                ItalicButton.Checked = true;
            }

            if (CurrentTextBox != null)
            {
                CurrentTextBox.Font = new Font(CurrentTextBox.Font, FontStyle.Italic ^ CurrentTextBox.Font.Style); ;
            }
        }

        private void UnderlineButton_Click(object sender, EventArgs e)
        {
            if (UnderlineButton.Checked)
            {
                UnderlineButton.Checked = false;
            }

            else
            {
                UnderlineButton.Checked = true;
            }

            if (CurrentTextBox != null)
            {
                CurrentTextBox.Font = new Font(CurrentTextBox.Font, FontStyle.Underline ^ CurrentTextBox.Font.Style); ;
            }
        }

        private void FillcolorButton_Click(object sender, EventArgs e)
        {
            if (CurrentTextBox != null)
            {
                var col = new ColorDialog();
                col.Color = CurrentTextBox.BackColor;
                var res = col.ShowDialog();
                if (res == DialogResult.OK)
                {
                   CurrentTextBox.BackColor = col.Color;

                }
            }
        }

        private void FontcolorButton_Click(object sender, EventArgs e)
        {
            if (CurrentTextBox != null)
            {
                var col = new ColorDialog();
                col.Color = CurrentTextBox.ForeColor;
                var res = col.ShowDialog();
                if (res == DialogResult.OK)
                {
                    CurrentTextBox.ForeColor = col.Color;
                }
            }
        }

        private void FontSettingsButton_Click(object sender, EventArgs e)
        {
            if (CurrentTextBox != null)
            {
                var fd = new FontDialog();
                fd.Font = CurrentTextBox.Font;
                var res = fd.ShowDialog();
                if (res == DialogResult.OK)
                {
                    CurrentTextBox.Font = fd.Font;
                }
            }
        }

        private void AlignLeftButton_Click(object sender, EventArgs e)
        {
            if (CurrentTextBox != null)
            {
                CurrentTextBox.TextAlign = HorizontalAlignment.Left;
                CurrentTextBox.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void CenterTextButton_Click(object sender, EventArgs e)
        {
            if (CurrentTextBox != null)
            {
                CurrentTextBox.TextAlign = HorizontalAlignment.Center;
                CurrentTextBox.TextAlign = HorizontalAlignment.Center;
            }
        }

        private void AlignRightButton_Click(object sender, EventArgs e)
        {
            if (CurrentTextBox != null)
            {
                CurrentTextBox.TextAlign = HorizontalAlignment.Right;
                CurrentTextBox.TextAlign = HorizontalAlignment.Right;
            }
        }

        /// <summary>
        /// Sets controls to double buffered
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000;
                return cp;
            }
        }

        private void ANY_MouseDown_Restore_TextBox_State(object sender, MouseEventArgs e)
        {
            if (!(sender is Syncfusion.Windows.Forms.Tools.TextBoxExt))
            {
                if (CurrentTextBox != null)
                {
                    CurrentTextBox.BorderColor = Color.Black;
                    CurrentTextBox.CornerRadius = 0;
                }
                CurrentTextBox = null;
            }
        }
    }
}

internal class FlickerFreeListBox : System.Windows.Forms.ListBox
{
    public FlickerFreeListBox()
    {
        this.SetStyle(
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw |
            ControlStyles.UserPaint,
            true);
        this.DrawMode = DrawMode.OwnerDrawFixed;
    }
    //protected override void OnDrawItem(DrawItemEventArgs e)
    //{
    //    if (this.Items.Count > 0)
    //    {
    //        e.DrawBackground();
    //        e.Graphics.DrawString(this.Items[e.Index].ToString(), e.Font, new SolidBrush(this.ForeColor), new PointF(e.Bounds.X, e.Bounds.Y));
    //    }
    //    base.OnDrawItem(e);
    //}
    //protected override void OnPaint(PaintEventArgs e)
    //{
    //    Region iRegion = new Region(e.ClipRectangle);
    //    e.Graphics.FillRegion(new SolidBrush(this.BackColor), iRegion);
    //    if (this.Items.Count > 0)
    //    {
    //        for (int i = 0; i < this.Items.Count; ++i)
    //        {
    //            System.Drawing.Rectangle irect = this.GetItemRectangle(i);
    //            if (e.ClipRectangle.IntersectsWith(irect))
    //            {
    //                if ((this.SelectionMode == SelectionMode.One && this.SelectedIndex == i)
    //                || (this.SelectionMode == SelectionMode.MultiSimple && this.SelectedIndices.Contains(i))
    //                || (this.SelectionMode == SelectionMode.MultiExtended && this.SelectedIndices.Contains(i)))
    //                {
    //                    OnDrawItem(new DrawItemEventArgs(e.Graphics, this.Font,
    //                        irect, i,
    //                        DrawItemState.Selected, this.ForeColor,
    //                        this.BackColor));
    //                }
    //                else
    //                {
    //                    OnDrawItem(new DrawItemEventArgs(e.Graphics, this.Font,
    //                        irect, i,
    //                        DrawItemState.Default, this.ForeColor,
    //                        this.BackColor));
    //                }
    //                iRegion.Complement(irect);
    //            }
    //        }
    //    }
    //    base.OnPaint(e);
    //}
}