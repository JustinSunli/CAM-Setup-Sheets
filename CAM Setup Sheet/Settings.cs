using System;
using System.Windows.Forms;
using ListBox = System.Windows.Forms.ListBox;

namespace CAM_Setup_Sheets
{
    public partial class Settings : Form
    {
        public object LbItem = null;

        public Settings()
        {
            InitializeComponent();



            //Properties.Settings.Default.Reset();

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

            // Populate Tool List Items
            Tool_Items_NOT_ToUse_listBox.Items.Clear();

            if (Properties.Settings.Default.Tool_ItemsNOTtoUse != null)
            {
                Tool_Items_NOT_ToUse_listBox.Items.Clear();
                foreach (String item in Properties.Settings.Default.Tool_ItemsNOTtoUse)
                {
                    Tool_Items_NOT_ToUse_listBox.Items.Add(item);
                }
            }

            // Get Tool Items to Use
            Tool_Items_ToUse_listBox.Items.Clear();

            if (Properties.Settings.Default.Tool_ItemsToUse != null)
            {

                foreach (String item in Properties.Settings.Default.Tool_ItemsToUse)
                {
                    Tool_Items_ToUse_listBox.Items.Add(item);
                }
            }

            // Get Post Parameters Items for Header Items List
            listBoxPostParameters.Items.Clear();

            if (CAM_Setup_Sheets_Addin._PostParameterNames != null)
            {

                foreach (String item in CAM_Setup_Sheets_Addin._PostParameterNames)
                {
                    listBoxPostParameters.Items.Add(item);
                }
            }


            // Show Operations Panel
            Operations_splitContainer.Dock = DockStyle.Fill;
            Operations_splitContainer.Show();

        }

        private void MoveListBoxItems(ListBox source, ListBox destination)
        {
            ListBox.SelectedObjectCollection sourceItems = source.SelectedItems;
            foreach (var item in sourceItems)
            {
                destination.Items.Add(item);
            }
            while (source.SelectedItems.Count > 0)
            {
                source.Items.Remove(source.SelectedItems[0]);
            }

            OperationsList_RightArrowbutton.Enabled = false;
            OperationsList_UpArrowbutton.Enabled = false;
            OperationList_DownArrowbutton.Enabled = false;
            OperationList_LeftArrowbutton.Enabled = false;
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

        public void MoveToolListToUse_ListItem(int direction)
        {
            // Checking selected item
            if (Tool_Items_ToUse_listBox.SelectedItem == null || Tool_Items_ToUse_listBox.SelectedIndex < 0)
                return; // No selected item - nothing to do

            // Calculate new index using move direction
            int newIndex = Tool_Items_ToUse_listBox.SelectedIndex + direction;

            // Checking bounds of the range
            if (newIndex < 0 || newIndex >= Tool_Items_ToUse_listBox.Items.Count)
                return; // Index out of range - nothing to do

            object selected = Tool_Items_ToUse_listBox.SelectedItem;

            // Removing removable element
            Tool_Items_ToUse_listBox.Items.Remove(selected);
            // Insert it in new position
            Tool_Items_ToUse_listBox.Items.Insert(newIndex, selected);
            // Restore selection
            Tool_Items_ToUse_listBox.SetSelected(newIndex, true);
        }


        private void Settings_Load(object sender, EventArgs e)
        {
            //this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void Settings_FormClosing(object sender, FormClosingEventArgs e)
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

            var newToolItemsToUse = new System.Collections.Specialized.StringCollection();
            foreach (String item in Tool_Items_ToUse_listBox.Items)
            {
                newToolItemsToUse.Add(item);
            }

            var newToolItemsNOTToUse = new System.Collections.Specialized.StringCollection();
            foreach (String item in Tool_Items_NOT_ToUse_listBox.Items)
            {
                newToolItemsNOTToUse.Add(item);
            }

            Properties.Settings.Default.OperationItemsToUse = newOperationItemsToUse;
            Properties.Settings.Default.OperationItemsNOTtoUse = newOperationItemsNOTToUse;

            Properties.Settings.Default.Tool_ItemsToUse = newToolItemsToUse;
            Properties.Settings.Default.Tool_ItemsNOTtoUse = newToolItemsNOTToUse;
            Properties.Settings.Default.Save();
        }

        private void OperationItems_label_Click(object sender, EventArgs e)
        {
            Operations_splitContainer.Dock = DockStyle.Fill;
            Operations_splitContainer.Visible = true;

            Tools_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void Operations_ToolListItems_label_Click(object sender, EventArgs e)
        {
            Tools_splitContainer.Dock = DockStyle.Fill;
            Tools_splitContainer.Visible = true;

            Operations_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void Operations_OtherOptions_Click(object sender, EventArgs e)
        {
            OtherOptionsContainer.Dock = DockStyle.Fill;
            OtherOptionsContainer.Visible = true;

            Operations_splitContainer.Visible = false;
            Tools_splitContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void Operations_OperationHeaderItems_label_Click(object sender, EventArgs e)
        {
            OperationListHeaderItems_Container.Visible = true;
            OperationListHeaderItems_Container.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            Tools_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void Operations_ToolListHeaderItems_label_Click(object sender, EventArgs e)
        {
            ToolListHeaderItems_Container.Visible = true;
            ToolListHeaderItems_Container.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            Tools_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
        }

        private void ToolList_ToolListItems_label_Click(object sender, EventArgs e)
        {
            Tools_splitContainer.Dock = DockStyle.Fill;
            Tools_splitContainer.Visible = true;

            Operations_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void ToolList_OperationItems_label_Click(object sender, EventArgs e)
        {
            Operations_splitContainer.Dock = DockStyle.Fill;
            Operations_splitContainer.Visible = true;

            Tools_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void ToolList_OperationHeaderItems_label_Click(object sender, EventArgs e)
        {
            OperationListHeaderItems_Container.Visible = true;
            OperationListHeaderItems_Container.Dock = DockStyle.Fill;

            Tools_splitContainer.Visible = false;
            Operations_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void ToolList_ToolListHeaderItems_label_Click(object sender, EventArgs e)
        {
            ToolListHeaderItems_Container.Visible = true;
            ToolListHeaderItems_Container.Dock = DockStyle.Fill;

            Tools_splitContainer.Visible = false;
            Operations_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
        }

        private void ToolList_OtherOptions_Click(object sender, EventArgs e)
        {
            OtherOptionsContainer.Dock = DockStyle.Fill;
            OtherOptionsContainer.Visible = true;

            Operations_splitContainer.Visible = false;
            Tools_splitContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void OpHeader_OperationsList_label_Click(object sender, EventArgs e)
        {
            Operations_splitContainer.Visible = true;
            Operations_splitContainer.Dock = DockStyle.Fill;

            OtherOptionsContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            Tools_splitContainer.Visible = false;
        }

        private void OpHeader_ToolList_label_Click(object sender, EventArgs e)
        {
            Tools_splitContainer.Visible = true;
            Tools_splitContainer.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
            OperationListHeaderItems_Container.Visible = false;

        }

        private void OpHeader_ToolListHeader_label_Click(object sender, EventArgs e)
        {
            ToolListHeaderItems_Container.Visible = true;
            ToolListHeaderItems_Container.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            Tools_splitContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
        }

        private void OpHeader_OtherSettings_label_Click(object sender, EventArgs e)
        {
            OtherOptionsContainer.Visible = true;
            OtherOptionsContainer.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
            Tools_splitContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
        }

        private void ToolListHeader_OperationsList_label_Click(object sender, EventArgs e)
        {
            Operations_splitContainer.Visible = true;
            Operations_splitContainer.Dock = DockStyle.Fill;

            Tools_splitContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
        }

        private void ToolListHeader_ToolListItems_label_Click(object sender, EventArgs e)
        {
            Tools_splitContainer.Visible = true;
            Tools_splitContainer.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
        }

        private void ToolListHeader_OperationsListHeader_label_Click(object sender, EventArgs e)
        {
            OperationListHeaderItems_Container.Visible = true;
            OperationListHeaderItems_Container.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
            OtherOptionsContainer.Visible = false;
            Tools_splitContainer.Visible = false;
        }

        private void ToolListHeader_OtherSettings_label_Click(object sender, EventArgs e)
        {
            OtherOptionsContainer.Visible = true;
            OtherOptionsContainer.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            Tools_splitContainer.Visible = false;
        }

        private void OtherOtions_OperationList_Click(object sender, EventArgs e)
        {
            Operations_splitContainer.Dock = DockStyle.Fill;
            Operations_splitContainer.Visible = true;

            Tools_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void OtherOptions_ToolList_Click(object sender, EventArgs e)
        {
            Tools_splitContainer.Dock = DockStyle.Fill;
            Tools_splitContainer.Visible = true;

            Operations_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void OtherOptions_OperationListHeader_label_Click(object sender, EventArgs e)
        {
            OperationListHeaderItems_Container.Visible = true;
            OperationListHeaderItems_Container.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            Tools_splitContainer.Visible = false;
            ToolListHeaderItems_Container.Visible = false;
        }

        private void OtherOptions_ToolListHeader_label_Click(object sender, EventArgs e)
        {
            ToolListHeaderItems_Container.Visible = true;
            ToolListHeaderItems_Container.Dock = DockStyle.Fill;

            Operations_splitContainer.Visible = false;
            OtherOptionsContainer.Visible = false;
            Tools_splitContainer.Visible = false;
            OperationListHeaderItems_Container.Visible = false;
        }

        private void TextColorforTitleRowbutton_Click(object sender, EventArgs e)
        {
            ColorDialog col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorFor1stHeaderRow;
            DialogResult res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorFor1stHeaderRow = col.Color;
            }
        }
        private void TextColorforTopRowbutton_Click(object sender, EventArgs e)
        {
            ColorDialog col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorForHeaderRow;
            DialogResult res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorForHeaderRow = col.Color;
            }
        }

        private void TextColorForDataRowsButton_Click(object sender, EventArgs e)
        {
            ColorDialog col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorForDataRows;
            DialogResult res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorForDataRows = col.Color;
            }
        }
        private void TextColorForAlternatingRowsButton_Click(object sender, EventArgs e)
        {
            ColorDialog col = new ColorDialog();
            col.Color = Properties.Settings.Default.TextColorForAlternatingRows;
            DialogResult res = col.ShowDialog();
            if (res == DialogResult.OK)
            {
                Properties.Settings.Default.TextColorForAlternatingRows = col.Color;
            }
        }
        private void OperationsList_RightArrowbutton_Click(object sender, EventArgs e)
        {
            MoveListBoxItems(Operations_NOT_ToUse_ListBox, Operations_ToUse_ListBox);
        }

        private void OperationsList_LeftArrowbutton_Click(object sender, EventArgs e)
        {
            MoveListBoxItems(Operations_ToUse_ListBox, Operations_NOT_ToUse_ListBox);
        }

        private void OperationsList_UpArrowbutton_Click(object sender, EventArgs e)
        {
            MoveOperationToUse_ListItem(-1);
        }

        private void OperationsList_DownArrowbutton_Click(object sender, EventArgs e)
        {
            MoveOperationToUse_ListItem(1);
        }

        private void ToolList_RightArrow_Click(object sender, EventArgs e)
        {
            MoveListBoxItems(Tool_Items_NOT_ToUse_listBox, Tool_Items_ToUse_listBox);
        }

        private void ToolList_LeftArrow_Click(object sender, EventArgs e)
        {
            MoveListBoxItems(Tool_Items_ToUse_listBox, Tool_Items_NOT_ToUse_listBox);
        }

        private void ToolList_UpArrow_Click(object sender, EventArgs e)
        {
            MoveToolListToUse_ListItem(-1);
        }

        private void ToolList_DownArrow_Click(object sender, EventArgs e)
        {
            MoveToolListToUse_ListItem(1);
        }

        private void Operations_NOT_ToUse_ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            OperationsList_RightArrowbutton.Enabled = true;
            OperationsList_UpArrowbutton.Enabled = false;
            OperationList_DownArrowbutton.Enabled = false;
            OperationList_LeftArrowbutton.Enabled = false;
        }

        private void Operations_ToUse_ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            OperationsList_UpArrowbutton.Enabled = true;
            OperationList_DownArrowbutton.Enabled = true;
            OperationList_LeftArrowbutton.Enabled = true;
            OperationsList_RightArrowbutton.Enabled = false;
        }

        private void Tool_Items_NOT_ToUse_listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ToolList_RightArrow.Enabled = true;
            ToolList_UpArrow.Enabled = false;
            ToolList_DownArrow.Enabled = false;
            ToolList_LeftArrow.Enabled = false;
        }

        private void Tool_Items_ToUse_listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ToolList_UpArrow.Enabled = true;
            ToolList_DownArrow.Enabled = true;
            ToolList_LeftArrow.Enabled = true;
            ToolList_RightArrow.Enabled = false;
        }

        private void TextFontForTitleRowButton_Click(object sender, EventArgs e)
        {
            FontDialog fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorFor1stHeaderRow;

            fd.Font = Properties.Settings.Default.TextFontFor1stHeaderRow;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontFor1stHeaderRow = fd.Font;
            Properties.Settings.Default.Save();
        }
        private void TextFontForDataRows_Click(object sender, EventArgs e)
        {
            FontDialog fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorForDataRows;

            fd.Font = Properties.Settings.Default.TextFontForDataRows;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForDataRows = fd.Font;
            Properties.Settings.Default.Save();
        }

        private void TextFontForHeaderRowButton_Click(object sender, EventArgs e)
        {
            FontDialog fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorForHeaderRow;


            fd.Font = Properties.Settings.Default.TextFontForHeaderRow;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForHeaderRow = fd.Font;
            Properties.Settings.Default.Save();
        }

        private void TextFontForAlternatingRowsButton_Click(object sender, EventArgs e)
        {
            FontDialog fd = new FontDialog();
            fd.Color = Properties.Settings.Default.TextColorForAlternatingRows;


            fd.Font = Properties.Settings.Default.TextFontForAlternatingRows;

            fd.ShowDialog();

            Properties.Settings.Default.TextFontForAlternatingRows = fd.Font;
            Properties.Settings.Default.Save();
        }

        private void listBoxPostParameters_MouseDown(object sender, MouseEventArgs e)
        {
            LbItem = null;

            if (listBoxPostParameters.Items.Count == 0)
            {
                return;
            }

            int index = listBoxPostParameters.IndexFromPoint(e.X, e.Y);
            if (index != -1)
            {
                string s = listBoxPostParameters.Items[index].ToString();
                DragDropEffects dde1 = DoDragDrop(s, DragDropEffects.All);
            }
        }

        private void Column1Row1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column1Row1_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column1Row1.Text != String.Empty)
            {
                oldtext = Column1Row1.Text;
            }
            Column1Row1.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column1Row2_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column1Row2_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column1Row2.Text != String.Empty)
            {
                oldtext = Column1Row2.Text;
            }
            Column1Row2.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column1Row3_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column1Row3_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column1Row3.Text != String.Empty)
            {
                oldtext = Column1Row3.Text;
            }
            Column1Row3.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column1Row4_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column1Row4_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column1Row4.Text != String.Empty)
            {
                oldtext = Column1Row4.Text;
            }
            Column1Row4.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column1Row5_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column1Row5_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column1Row5.Text != String.Empty)
            {
                oldtext = Column1Row5.Text;
            }
            Column1Row5.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column1Row6_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column1Row6_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column1Row6.Text != String.Empty)
            {
                oldtext = Column1Row6.Text;
            }
            Column1Row6.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column1Row7_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column1Row7_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column1Row7.Text != String.Empty)
            {
                oldtext = Column1Row7.Text;
            }
            Column1Row7.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column1Row8_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column1Row8_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column1Row8.Text != String.Empty)
            {
                oldtext = Column1Row8.Text;
            }
            Column1Row8.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column1Row9_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column1Row9_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column1Row9.Text != String.Empty)
            {
                oldtext = Column1Row9.Text;
            }
            Column1Row9.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column2Row1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column2Row1_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column2Row1.Text != String.Empty)
            {
                oldtext = Column2Row1.Text;
            }
            Column2Row1.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column2Row2_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column2Row2_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column2Row2.Text != String.Empty)
            {
                oldtext = Column2Row2.Text;
            }
            Column2Row2.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column2Row3_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column2Row3_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column2Row3.Text != String.Empty)
            {
                oldtext = Column2Row3.Text;
            }
            Column2Row3.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column2Row4_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column2Row4_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column2Row4.Text != String.Empty)
            {
                oldtext = Column2Row4.Text;
            }
            Column2Row4.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column2Row5_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column2Row5_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column2Row5.Text != String.Empty)
            {
                oldtext = Column2Row5.Text;
            }
            Column2Row5.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column2Row6_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column2Row6_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column2Row6.Text != String.Empty)
            {
                oldtext = Column2Row6.Text;
            }
            Column2Row6.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column2Row7_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column2Row7_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column2Row7.Text != String.Empty)
            {
                oldtext = Column2Row7.Text;
            }
            Column2Row7.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column2Row8_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column2Row8_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column2Row8.Text != String.Empty)
            {
                oldtext = Column2Row8.Text;
            }
            Column2Row8.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void Column2Row9_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Column2Row9_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (Column2Row9.Text != String.Empty)
            {
                oldtext = Column2Row9.Text;
            }
            Column2Row9.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void TopHeaderLine1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void TopHeaderLine1_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (TopHeaderLine1.Text != String.Empty)
            {
                oldtext = TopHeaderLine1.Text;
            }
            TopHeaderLine1.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }

        private void TopHeaderLine2_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void TopHeaderLine2_DragDrop(object sender, DragEventArgs e)
        {
            var s = e.Data.GetData(DataFormats.StringFormat).ToString();
            String oldtext = String.Empty;
            if (TopHeaderLine2.Text != String.Empty)
            {
                oldtext = TopHeaderLine2.Text;
            }
            TopHeaderLine2.Text = s;
            listBoxPostParameters.Items.Remove(s);
            if (oldtext != String.Empty)
            {
                listBoxPostParameters.Items.Add(oldtext);
            }
        }
    }
}
