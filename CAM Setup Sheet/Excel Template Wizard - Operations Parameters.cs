using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Font = System.Drawing.Font;

namespace CAM_Setup_Sheets
{
    public partial class ExcelTemplateWizard_OperationParameters : Form
    {
        private Worksheet worksheet = null;
        private System.Drawing.Point mDownPos;

        public ExcelTemplateWizard_OperationParameters()
        {
            InitializeComponent();

            // Set a region at top of tab control to block out tabs
            TabControl1.Region = new Region(new RectangleF(tabPage4.Left, tabPage4.Top,
                tabPage4.Width, tabPage4.Height));

            // Get Operation Items TO Use
            Operations_ToUse_ListBox.Items.Clear();

            if (Properties.Settings.Default.OperationItemsToUse != null)
            {
                Operations_ToUse_ListBox.Items.Clear();
                foreach (String item in Properties.Settings.Default.OperationItemsToUse)
                {
                    Operations_ToUse_ListBox.Items.Add(item);
                }
            }

            // Get Operation Items NOT to Use
            Operations_NOT_ToUse_ListBox.Items.Clear();

            if (Properties.Settings.Default.OperationItemsNOTtoUse != null)
            {
                Operations_NOT_ToUse_ListBox.Items.Clear();
                foreach (String item in Properties.Settings.Default.OperationItemsNOTtoUse)
                {
                    Operations_NOT_ToUse_ListBox.Items.Add(item);
                }
            }


            // Get Post Parameters Items for Header Items List
            listBoxPostParametersTab5.Items.Clear();
            listBoxPostParametersTab6.Items.Clear();
            listBoxPostParametersTab7.Items.Clear();
            listBoxPostParametersTab8.Items.Clear();
            listBoxPostParametersTab9.Items.Clear();

            if (CAM_Setup_Sheets_Addin._PostParameterNames != null)
            {
                foreach (var item in CAM_Setup_Sheets_Addin._PostParameterNames)
                {
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
        private void ForeColorForOperationParameterHeader_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.ForeColorForOperationParameterHeadings;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.ForeColorForOperationParameterHeadings = col.Color;
                
            }
        }


        private void ForeColorForOperationParameterDataRow_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.ForeColorForOperationDataRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.ForeColorForOperationDataRow = col.Color;
                
            }
        }
        private void ForeColorForOperationParameterAlternatingDataRow_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.ForeColorForOperationAlternatingDataRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.ForeColorForOperationAlternatingDataRow = col.Color;
                
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
        private void BackColorForOperationParameterHeader_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.BackColorForOperationParameterHeadings;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.BackColorForOperationParameterHeadings = col.Color;
                
            }
        }


        private void BackColorForOperationParameterDataRow_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.BackColorForOperationDataRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.BackColorForOperationDataRow = col.Color;
                
            }
        }
        private void BackColorForOperationParameterAlternatingDataRow_Click(object sender, EventArgs e)
        {
            var col = new ColorDialog();
            col.Color = Properties.Settings.Default.BackColorForOperationAlternatingDataRow;
            var res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.BackColorForOperationAlternatingDataRow = col.Color;
                
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
        private void FontForOperationParameterHeader_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.ForeColorForOperationParameterHeadings;

            fd.Font = Properties.Settings.Default.FontForOperationParameterHeadings;

            fd.ShowDialog();

            Properties.Settings.Default.FontForOperationParameterHeadings = fd.Font;
            
        }


        private void FontForOperationParameterDataRow_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.ForeColorForOperationDataRow;

            fd.Font = Properties.Settings.Default.FontForOperationDataRow;

            fd.ShowDialog();

            Properties.Settings.Default.FontForOperationDataRow = fd.Font;
            
        }
        private void FontForOperationParameterAlternatingDataRow_Click(object sender, EventArgs e)
        {
            var fd = new FontDialog();
            fd.Color = Properties.Settings.Default.ForeColorForOperationAlternatingDataRow;

            fd.Font = Properties.Settings.Default.FontForOperationAlternatingDataRow;

            fd.ShowDialog();

            Properties.Settings.Default.FontForOperationAlternatingDataRow = fd.Font;
            
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
            if (sender.GetType().Name == "TextBox")
            {
                var s = e.Data.GetData(DataFormats.StringFormat).ToString();
                System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)sender;
                tb.Text += "<" + s + ">";
            }
        }
        private void ANYTextBox_DragEnter(object sender, DragEventArgs e)
        {
            if (sender.GetType().Name == "TextBox")
            {
                e.Effect = DragDropEffects.All;
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
        private void Next9_Click(object sender, EventArgs e)
        {
            TabControl1.SelectedIndex = 9;
            
        }

        private void Back4_Click(object sender, EventArgs e)
        {
            var newOperationItemsToUse = new System.Collections.Specialized.StringCollection();
            //Properties.Settings.Default.OperationItemsToUse.Clear();
            foreach (String item in Operations_ToUse_ListBox.Items)
            {
                newOperationItemsToUse.Add(item);
            }

            var newOperationItemsNOTToUse = new System.Collections.Specialized.StringCollection();
            //Properties.Settings.Default.OperationItemsNOTtoUse.Clear();
            foreach (String item in Operations_NOT_ToUse_ListBox.Items)
            {
                newOperationItemsNOTToUse.Add(item);
            }

            Properties.Settings.Default.OperationItemsToUse = newOperationItemsToUse;
            Properties.Settings.Default.OperationItemsNOTtoUse = newOperationItemsNOTToUse;



            this.Close();
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

        private void Button_ResetAllValues_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Reset();
        }


        private void ButtonPreviewInExcelTab4_Click(object sender, EventArgs e)
        {
            var newOperationItemsToUse = new System.Collections.Specialized.StringCollection();
            foreach (String item in Operations_ToUse_ListBox.Items)
            {
                newOperationItemsToUse.Add(item);
            }

            var newOperationItemsNOTToUse = new System.Collections.Specialized.StringCollection();
            foreach (String item in Operations_NOT_ToUse_ListBox.Items)
            {
                newOperationItemsNOTToUse.Add(item);
            }

            Properties.Settings.Default.OperationItemsToUse = newOperationItemsToUse;
            Properties.Settings.Default.OperationItemsNOTtoUse = newOperationItemsNOTToUse;


            

            CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = Properties.Settings.Default.OperationItemsToUse.Count;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate<15)
            {
                CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate = 15;
            }

            CreateExcel();  
        }

        private void ButtonPreviewInExcelTab5_Click(object sender, EventArgs e)
        {
            int NumberOfColumns = 0;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate == 0)
            {
                NumberOfColumns = 20;
            }

            ButtonPreviewInExcelTab4_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[7, 1].Value = textBoxRow7LeftSide.Text.ToString();
                worksheet.Cells[8, 1].Value = textBoxRow8LeftSide.Text.ToString();
                worksheet.Cells[9, 1].Value = textBoxRow9LeftSide.Text.ToString();

                worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows7to9].Value =
                    textBoxRow7LeftSideParameters.Text.ToString();
                worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows7to9].Value =
                    textBoxRow8LeftSideParameters.Text.ToString();
                worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows7to9].Value =
                    textBoxRow9LeftSideParameters.Text.ToString();

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
            int NumberOfColumns = 0;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate == 0)
            {
                NumberOfColumns = 20;
            }

            ButtonPreviewInExcelTab5_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[10, 1].Value = textBoxRow10LeftSide.Text.ToString();
                worksheet.Cells[11, 1].Value = textBoxRow11LeftSide.Text.ToString();
                worksheet.Cells[12, 1].Value = textBoxRow12LeftSide.Text.ToString();

                worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12].Value =
                    textBoxRow10LeftSideParameters.Text.ToString();
                worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12].Value =
                    textBoxRow11LeftSideParameters.Text.ToString();
                worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12].Value =
                    textBoxRow12LeftSideParameters.Text.ToString();

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
            int NumberOfColumns = 0;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate == 0)
            {
                NumberOfColumns = 20;
            }

            ButtonPreviewInExcelTab6_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[4, 1 + (NumberOfColumns / 2)].Value = TextBoxRow4RightSide.Text.ToString();
                worksheet.Cells[5, 1 + (NumberOfColumns / 2)].Value = TextBoxRow5RightSide.Text.ToString();
                worksheet.Cells[6, 1 + (NumberOfColumns / 2)].Value = TextBoxRow6RightSide.Text.ToString();

                worksheet.Cells[4, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)].Value =
                    textBoxRow4RightSideParameters.Text.ToString();
                worksheet.Cells[5, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)].Value =
                    textBoxRow5RightSideParameters.Text.ToString();
                worksheet.Cells[6, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)].Value =
                    textBoxRow6RightSideParameters.Text.ToString();

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
            int NumberOfColumns = 0;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate == 0)
            {
                NumberOfColumns = 20;
            }

            ButtonPreviewInExcelTab7_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[7, 1 + (NumberOfColumns / 2)].Value = TextBoxRow7RightSide.Text.ToString();
                worksheet.Cells[8, 1 + (NumberOfColumns / 2)].Value = TextBoxRow8RightSide.Text.ToString();
                worksheet.Cells[9, 1 + (NumberOfColumns / 2)].Value = TextBoxRow9RightSide.Text.ToString();

                worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)].Value =
                    textBoxRow7RightSideParameters.Text.ToString();
                worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)].Value =
                    textBoxRow8RightSideParameters.Text.ToString();
                worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)].Value =
                    textBoxRow9RightSideParameters.Text.ToString();

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
            int NumberOfColumns = 0;
            if (CAM_Setup_Sheets_Addin._NumberOfOperationParametersForTemplate == 0)
            {
                NumberOfColumns = 20;
            }

            buttonPreviewInExcelTab8_Click(this, null);

            if (worksheet != null)
            {
                // Set text
                worksheet.Cells[10, 1 + (NumberOfColumns / 2)].Value = TextBoxRow10RightSide.Text.ToString();
                worksheet.Cells[11, 1 + (NumberOfColumns / 2)].Value = TextBoxRow11RightSide.Text.ToString();
                worksheet.Cells[12, 1 + (NumberOfColumns / 2)].Value = TextBoxRow12RightSide.Text.ToString();

                worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)].Value =
                    textBoxRow10RightSideParameters.Text.ToString();
                worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)].Value =
                    textBoxRow11RightSideParameters.Text.ToString();
                worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)].Value =
                    textBoxRow12RightSideParameters.Text.ToString();

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

        private void Operations_NOT_ToUse_ListBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (Operations_NOT_ToUse_ListBox.Items.Count == 0) return;

            mDownPos = e.Location;

            //var index = Operations_NOT_ToUse_ListBox.IndexFromPoint(e.X, e.Y);
            //if (index != -1)
            //{
            //    var s = Operations_NOT_ToUse_ListBox.Items[index].ToString();
            //    var dde1 = DoDragDrop(s, DragDropEffects.All);
            //}
        }

        private void Operations_NOT_ToUse_ListBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            int index = Operations_NOT_ToUse_ListBox.IndexFromPoint(e.Location);
            if (index < 0) return;
            if (Math.Abs(e.X - mDownPos.X) >= SystemInformation.DragSize.Width ||
                Math.Abs(e.Y - mDownPos.Y) >= SystemInformation.DragSize.Height)
                DoDragDrop(new DragObject(Operations_NOT_ToUse_ListBox, Operations_NOT_ToUse_ListBox.Items[index]), DragDropEffects.Move);
        }

        private void Operations_NOT_ToUse_ListBox_DragDrop(object sender, DragEventArgs e)
        {
            DragObject obj = e.Data.GetData(typeof(DragObject)) as DragObject;
            Operations_NOT_ToUse_ListBox.Items.Add(obj.item);
            obj.source.Items.Remove(obj.item);
        }

        private void Operations_NOT_ToUse_ListBox_DragEnter(object sender, DragEventArgs e)
        {
            DragObject obj = e.Data.GetData(typeof(DragObject)) as DragObject;
            if (obj != null && obj.source != Operations_NOT_ToUse_ListBox) e.Effect = e.AllowedEffect;
        }

        private void Operations_ToUse_ListBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (Operations_ToUse_ListBox.Items.Count == 0) return;

            mDownPos = e.Location;
        }

        private void Operations_ToUse_ListBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            int index = Operations_ToUse_ListBox.IndexFromPoint(e.Location);
            if (index < 0) return;
            if (Math.Abs(e.X - mDownPos.X) >= SystemInformation.DragSize.Width ||
                Math.Abs(e.Y - mDownPos.Y) >= SystemInformation.DragSize.Height)
                DoDragDrop(new DragObject(Operations_ToUse_ListBox, Operations_ToUse_ListBox.Items[index]), DragDropEffects.Move);
        }

        private void Operations_ToUse_ListBox_DragDrop(object sender, DragEventArgs e)
        {
            DragObject obj = e.Data.GetData(typeof(DragObject)) as DragObject;
            Operations_ToUse_ListBox.Items.Add(obj.item);
            obj.source.Items.Remove(obj.item);
        }
        private void Operations_ToUse_ListBox_DragEnter(object sender, DragEventArgs e)
        {
            DragObject obj = e.Data.GetData(typeof(DragObject)) as DragObject;
            if (obj != null && obj.source != Operations_ToUse_ListBox) e.Effect = e.AllowedEffect;
        }
        private void Operations_ToUse_ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            OperationList_DownArrowbutton.Enabled = true;
            OperationsList_UpArrowbutton.Enabled = true;
        }

        private void OperationsList_UpArrowbutton_Click(object sender, EventArgs e)
        {
            MoveOperationToUse_ListItem(-1);
        }

        private void OperationsList_DownArrowbutton_Click(object sender, EventArgs e)
        {
            MoveOperationToUse_ListItem(1);
        }

        public void MoveOperationToUse_ListItem(int direction)
        {
            // Checking selected item
            if (Operations_ToUse_ListBox.SelectedItem == null || Operations_ToUse_ListBox.SelectedIndex < 0)
                return; // No selected item - nothing to do

            // Calculate new index using move direction
            int newIndex = Operations_ToUse_ListBox.SelectedIndex + direction;

            // Checking bounds of the range
            if (newIndex < 0 || newIndex >= Operations_ToUse_ListBox.Items.Count)
                return; // Index out of range - nothing to do

            object selected = Operations_ToUse_ListBox.SelectedItem;

            // Removing removable element
            Operations_ToUse_ListBox.Items.Remove(selected);
            // Insert it in new position
            Operations_ToUse_ListBox.Items.Insert(newIndex, selected);
            // Restore selection
            Operations_ToUse_ListBox.SetSelected(newIndex, true);
        }
        private void ExcelTemplateWizard_OperationParameters_FormClosing(object sender, FormClosingEventArgs e)
        {
            var newOperationItemsToUse = new System.Collections.Specialized.StringCollection();
            //Properties.Settings.Default.OperationItemsToUse.Clear();
            foreach (String item in Operations_ToUse_ListBox.Items)
            {
                newOperationItemsToUse.Add(item);
            }

            var newOperationItemsNOTToUse = new System.Collections.Specialized.StringCollection();
            //Properties.Settings.Default.OperationItemsNOTtoUse.Clear();
            foreach (String item in Operations_NOT_ToUse_ListBox.Items)
            {
                newOperationItemsNOTToUse.Add(item);
            }

            Properties.Settings.Default.OperationItemsToUse = newOperationItemsToUse;
            Properties.Settings.Default.OperationItemsNOTtoUse = newOperationItemsNOTToUse;



            this.Close();
        }

        private void PopulatePostParameterValues(Microsoft.Office.Interop.Excel.Range usedrange)
        {
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
        }

        private void PopulateOperationParameterValues(Microsoft.Office.Interop.Excel.Range usedrange, int startrow)
        {
            int currentcell = 0;
            foreach (Range cell in usedrange.Cells)
            {
                currentcell++;
                if (cell.Value2 != null)
                {

                    String input = cell.Value2.ToString();

                    Regex regex = new Regex(@"\<.*?\>");

                    var arr = regex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();

                    for (int j = 0; j < arr.Length; j++)
                    {
                        String str = arr[j];
                        String clean = str.Replace("<", "");
                        clean = clean.Replace(">", "");

                        for (int i = 0; i < CAM_Setup_Sheets_Addin._Operations.Count; i++)
                        {
                            Machine_Operation operation = CAM_Setup_Sheets_Addin._Operations[i];
                            var t = operation.GetType();

                            foreach (PropertyInfo pi in t.GetProperties())
                            {
                                var dn = DisplayNameHelper.GetDisplayName(pi);
                                if (clean == dn)
                                {
                                    double doublevariable;

                                    input = pi.GetValue(operation, null).ToString();
                                    if (input == "-999999")
                                    {
                                        input = String.Empty;
                                    }
                                    bool isDouble = Double.TryParse(input, out doublevariable);
                                    if (isDouble)
                                    {
                                        input = Math.Round(doublevariable,(int)Properties.Settings.Default.NumberOfDecimalPlaces).ToString();
                                    }

                                    worksheet.Cells[15 + i, currentcell].Value = input;
                                }
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < CAM_Setup_Sheets_Addin._Operations.Count; i++)
            {
                int columns = usedrange.Columns.Count;

                var rng1 = worksheet.Range[worksheet.Cells[15 + i, 1],
                               worksheet.Cells[15 + i, columns]];

                rng1.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

                //Borders border = rng1.Borders;
                //border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                //border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                //border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                //border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;

                if ((15 + i) % 2 == 0)
                {
                    //is even
                    SetExcelFont(rng1, Properties.Settings.Default.ForeColorForOperationAlternatingDataRow,
                    Properties.Settings.Default.BackColorForOperationAlternatingDataRow,
                    Properties.Settings.Default.FontForOperationAlternatingDataRow);
                }
                else
                {
                    // is odd
                    SetExcelFont(rng1, Properties.Settings.Default.ForeColorForOperationDataRow,
                    Properties.Settings.Default.BackColorForOperationDataRow,
                    Properties.Settings.Default.FontForOperationDataRow);
                }
            }
        }

        public static class DisplayNameHelper
        {
            public static string GetDisplayName(object obj, string propertyName)
            {
                if (obj == null) return null;
                return GetDisplayName(obj.GetType(), propertyName);

            }

            public static string GetDisplayName(Type type, string propertyName)
            {
                var property = type.GetProperty(propertyName);
                if (property == null) return null;

                return GetDisplayName(property);
            }

            public static string GetDisplayName(PropertyInfo property)
            {
                var attrName = GetAttributeDisplayName(property);
                if (!string.IsNullOrEmpty(attrName))
                    return attrName;

                //var metaName = GetMetaDisplayName(property);
                //if (!string.IsNullOrEmpty(metaName))
                //    return metaName;

                return property.Name.ToString(CultureInfo.InvariantCulture);
            }

            private static string GetAttributeDisplayName(PropertyInfo property)
            {
                var atts = property.GetCustomAttributes(
                    typeof(DisplayNameAttribute), true);
                if (atts.Length == 0)
                    return null;
                var displayNameAttribute = atts[0] as DisplayNameAttribute;
                return displayNameAttribute != null ? displayNameAttribute.DisplayName : null;
            }

            //private static string GetMetaDisplayName(PropertyInfo property)
            //{
            //    if (property.DeclaringType != null)
            //    {
            //        var atts = property.DeclaringType.GetCustomAttributes(
            //            typeof(MetadataTypeAttribute), true);
            //        if (atts.Length == 0)
            //            return null;

            //        var metaAttr = atts[0] as MetadataTypeAttribute;
            //        if (metaAttr != null)
            //        {
            //            var metaProperty =
            //                metaAttr.MetadataClassType.GetProperty(property.Name);
            //            return metaProperty == null ? null : GetAttributeDisplayName(metaProperty);
            //        }
            //    }
            //    return null;
            //}
        }

        private void CreateExcel()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

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

            xlApp.Visible = false;
            xlApp.ScreenUpdating = false;
            var xlWorkbook = xlApp.Workbooks.Add();
            var newWorksheet = (Worksheet)xlWorkbook.Worksheets.Add();
            worksheet = newWorksheet;

            newWorksheet.Name = CAM_Setup_Sheets_Addin._SetupSheetType;

            newWorksheet.Activate();


            // Set top row text
            worksheet.Cells[1, 1].Value = Properties.Settings.Default.Row1Text;
            worksheet.Cells[2, 1].Value = Properties.Settings.Default.Row2Text;
            worksheet.Cells[3, 1].Value = Properties.Settings.Default.Row3Text;

            // Set row 4-6 left side text
            worksheet.Cells[4, 1].Value = Properties.Settings.Default.Row4Text;
            worksheet.Cells[5, 1].Value = Properties.Settings.Default.Row5Text;
            worksheet.Cells[6, 1].Value = Properties.Settings.Default.Row6Text;

            worksheet.Cells[4, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows3to6].Value = Properties.Settings.Default.Row4LeftSideTextParameters;
            worksheet.Cells[5, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows3to6].Value = Properties.Settings.Default.Row5LeftSideTextParameters;
            worksheet.Cells[6, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows3to6].Value = Properties.Settings.Default.Row6LeftSideTextParameters;

            // Set row 7-9 left side text
            worksheet.Cells[7, 1].Value = Properties.Settings.Default.Row7Text;
            worksheet.Cells[8, 1].Value = Properties.Settings.Default.Row8Text;
            worksheet.Cells[9, 1].Value = Properties.Settings.Default.Row9Text;

            worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows7to9].Value =
                Properties.Settings.Default.Row7LeftSideTextParameters;
            worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows7to9].Value =
                Properties.Settings.Default.Row8LeftSideTextParameters;
            worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows7to9].Value =
                Properties.Settings.Default.Row9LeftSideTextParameters;

            // Set row 10-12 left side text
            worksheet.Cells[10, 1].Value = Properties.Settings.Default.Row10Text;
            worksheet.Cells[11, 1].Value = Properties.Settings.Default.Row11Text;
            worksheet.Cells[12, 1].Value = Properties.Settings.Default.Row12Text;

            worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12].Value =
                Properties.Settings.Default.Row10LeftSideTextParameters;
            worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12].Value =
                Properties.Settings.Default.Row11LeftSideTextParameters;
            worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOfColumnsToMergeRows10to12].Value =
                Properties.Settings.Default.Row12LeftSideTextParameters;

            // Set row 4-6 Right Side text
            worksheet.Cells[4, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row4_RH_Text;
            worksheet.Cells[5, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row5_RH_Text;
            worksheet.Cells[6, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row6_RH_Text;

            worksheet.Cells[4, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)].Value =
                Properties.Settings.Default.Row4_RH_TextParameters;
            worksheet.Cells[5, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)].Value =
                Properties.Settings.Default.Row5_RH_TextParameters;
            worksheet.Cells[6, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows4to6 + (NumberOfColumns / 2)].Value =
                Properties.Settings.Default.Row6_RH_TextParameters;

            // Set Row 7-9 Right Side text
            worksheet.Cells[7, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row7_RH_Text;
            worksheet.Cells[8, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row8_RH_Text;
            worksheet.Cells[9, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row9_RH_Text;

            worksheet.Cells[7, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)].Value =
                Properties.Settings.Default.Row7_RH_TextParameters;
            worksheet.Cells[8, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)].Value =
                Properties.Settings.Default.Row8_RH_TextParameters;
            worksheet.Cells[9, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows7to9 + (NumberOfColumns / 2)].Value =
                Properties.Settings.Default.Row9_RH_TextParameters;

            // Set Row 10-12 Right Side text
            worksheet.Cells[10, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row10_RH_Text;
            worksheet.Cells[11, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row11_RH_Text;
            worksheet.Cells[12, 1 + (NumberOfColumns / 2)].Value = Properties.Settings.Default.Row12_RH_Text;

            worksheet.Cells[10, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)].Value =
                Properties.Settings.Default.Row10_RH_TextParameters;
            worksheet.Cells[11, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)].Value =
                Properties.Settings.Default.Row11_RH_TextParameters;
            worksheet.Cells[12, 2 + Properties.Settings.Default.NumberOf_RH_ColumnsToMergeRows10to12 + (NumberOfColumns / 2)].Value =
                Properties.Settings.Default.Row12_RH_TextParameters;

            // Set Row 1-3 Fonts and Colors
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

            // Set Set Left Side Row 4-6 Fonts and Colors
            rng1 = worksheet.Range["A4", Type.Missing];
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

            // Row 4-6 Left Side Parameter value Fonts and Colors
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

            // Set Row 7-9 Fonts and Colors
            rng1 = worksheet.Range["A7", Type.Missing];
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

            // Row 7-9 Left Side Parameter value Fonts and Colors
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

            // Row 10-12 Left Side Set Fonts and Colors
            rng1 = worksheet.Range["A10", Type.Missing];
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

            // Row 10-12 Left Side Parameter value side Fonts and Colors
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

            // Set Right Side Row 4-6 Fonts and Colors
            rng1 = worksheet.Range[worksheet.Cells[4, 1 + (NumberOfColumns / 2)],
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

            // Set Right Side Row 4-6 Parameter value side Fonts and Colors
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

            // Set Right Side Row 7-9 Fonts and Colors
            rng1 = worksheet.Range[worksheet.Cells[7, 1 + (NumberOfColumns / 2)],
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

            // Set Right Side Row 7-9 Parameter value side Fonts and Colors
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

            // Set Right Side Row 10-12 Fonts and Colors
            rng1 = worksheet.Range[worksheet.Cells[10, 1 + (NumberOfColumns / 2)],
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

            // Set Right Side Row 10-12 Parameter value Fonts and Colors
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
            var usedrange = newWorksheet.UsedRange;

            PopulatePostParameterValues(usedrange);

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

            // Rows 4-6 Left Side Merge and center if needed
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

            // Row 7-9 Left Side Merge and center if needed
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

            // Left Side Row 10-12 Merge and center if needed
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

            // Right Side Row 4-6 Merge and center if needed
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

            // Row 7-9 Merge and center if needed
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

            // Row 10-12 Merge and center if needed
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

            // Add Operation Parameters to Row 13

            for (int i = 0; i < Properties.Settings.Default.OperationItemsToUse.Count; i++)
            {
                var commands = Properties.Settings.Default.OperationItemsToUse[i].Split(' ');

                String[] headertext = new string[2];
                if (commands.Length == 1)
                {
                    headertext[0] = commands[0];
                    headertext[1] = string.Empty;
                }

                if (commands.Length == 2)
                {
                    headertext[0] = commands[0];
                    headertext[1] = commands[1];
                }
                if (commands.Length>2)
                {
                    for (int j = 0; j < (commands.Length) / 2; j++)
                    {
                        headertext[0] += commands[j] + " ";
                    }

                    for (int j = (commands.Length) / 2; j < commands.Length; j++)
                    {
                        headertext[1] += commands[j] + " ";
                    }
                }

                newWorksheet.Cells[13, i+1].Value = headertext[0];
                newWorksheet.Cells[14, i + 1].Value = headertext[1];
                newWorksheet.Cells[15, i + 1] = "<" + Properties.Settings.Default.OperationItemsToUse[i] + ">";

                rng1 = worksheet.Range[worksheet.Cells[13, i+1],worksheet.Cells[14, i+1]];
                SetExcelFont(rng1, Properties.Settings.Default.ForeColorForOperationParameterHeadings,
                    Properties.Settings.Default.BackColorForOperationParameterHeadings,
                    Properties.Settings.Default.FontForOperationParameterHeadings);

                Borders border = rng1.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;

                rng1.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                rng1= worksheet.Range[worksheet.Cells[15, 1], worksheet.Cells[15, NumberOfColumns]];



            }

            // Fill Operation Parameter Values
            PopulateOperationParameterValues(rng1, 15);

            rng1 = worksheet.UsedRange;

            rng1.Columns.AutoFit();

            // Excel Page Setup
            worksheet.PageSetup.PaperSize = XlPaperSize.xlPaperLetter;
            worksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            worksheet.PageSetup.BottomMargin = xlApp.InchesToPoints(0);
            worksheet.PageSetup.TopMargin = xlApp.InchesToPoints(0);
            worksheet.PageSetup.LeftMargin = xlApp.InchesToPoints(0);
            worksheet.PageSetup.RightMargin = xlApp.InchesToPoints(0);
            worksheet.PageSetup.BottomMargin = xlApp.InchesToPoints(0);
            worksheet.PageSetup.HeaderMargin = xlApp.InchesToPoints(0);
            worksheet.PageSetup.FooterMargin = xlApp.InchesToPoints(0);
            worksheet.PageSetup.FitToPagesWide = 1;
            worksheet.PageSetup.Zoom = false;
            worksheet.PageSetup.FitToPagesTall = false;
            worksheet.PageSetup.PrintArea = rng1.Address;
            ///worksheet.PageSetup.PrintTitleColumns = worksheet.Columns[worksheet.Cells[1, 1], worksheet.Cells[1, NumberOfColumns]];
            worksheet.PageSetup.PrintTitleRows = "$1:$15";


            //// Uncomment these when done
            //xlWorkbook.Close();
            //xlApp.Quit();

            xlApp.Visible = true;
            xlApp.ScreenUpdating = true;
            Marshal.ReleaseComObject(newWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);


            xlApp = null;
            xlWorkbook = null;
            newWorksheet = null;

            stopwatch.Stop();
            MessageBox.Show("Elapsed Time = " + stopwatch.Elapsed);
        }
        private class DragObject
        {
            public System.Windows.Forms.ListBox source;
            public object item;
            public DragObject(System.Windows.Forms.ListBox box, object data) { source = box; item = data; }
        }
    }
}



