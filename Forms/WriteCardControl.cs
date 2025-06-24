using StressUtilities.FEM;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace StressUtilities.Forms
{
    public partial class WriteCardControl : UserControl
    {
        private string[] CardsList = {"SUBCASE", "FORCE", "FORCE1", "FORCE2", "MOMENT", "MOMENT1", "MOMENT2",
                "SLOAD", "PLOAD", "PLOAD1", "PLOAD2", "PLOADB3", "PLOAD4", "PRESAX", "PLOADX1",
                "GRAV", "RFORCE", "ACCEL", "ACCEL1", "TEMP", "TEMPD", "TEMPAX", "TEMPBC", "SPC", "SPC1",
                "SPCD", "SPCAX", "DEFORM", "DAREA", "RLOAD1", "RLOAD2", "DLOAD", "TLOAD1", "TLOAD2"};
        private List<string> SupportedCards = new List<string>();

        public WriteCardControl()
        {
            InitializeComponent();
            this.Load += new EventHandler(WriteCardControl_Load);
        }


        private void BtnExport_Click(object sender, EventArgs e)
        {
            string FileName, Directory, FileNameFull;
            string TargetOS = OSBox.SelectedItem.ToString();
            string CardFormat = CardFormatBox.SelectedItem.ToString();

            if (!string.IsNullOrEmpty(PathBox.Text))
            {
                Directory = PathBox.Text;
            }
            else
            {
                MessageBox.Show("File Path Cannot be empty");
                return;
            }
            if (!string.IsNullOrEmpty(FileBox.Text))
            {
                FileName = FileBox.Text;
            }
            else
            {
                MessageBox.Show("File Name Cannot be empty");
                return;
            }
            if (!Directory.EndsWith(@"\"))
            {
                Directory += @"\";
            }

            FileNameFull = Directory + FileName;
            WriteNastranCards WriteCards = new WriteNastranCards();
            if (AllCardsList.Items.Count != 0)
            {
                List<string> CardList = new List<string>();
                foreach (string listitem in AllCardsList.Items)
                {
                    if (!string.IsNullOrEmpty(listitem))
                        CardList.Add(listitem);
                }
                //CardList=AllCardsList.Items.Cast<string>().ToList(); 
                if (TargetOS != Properties.Settings.Default.TargetOS)
                {
                    Properties.Settings.Default.TargetOS = TargetOS;
                }
                if (PathBox.Text != Properties.Settings.Default.WorkingDirectory)
                {
                    Properties.Settings.Default.WorkingDirectory = PathBox.Text;
                }

                

                WriteCards.WriteCards(CardList, FileNameFull, TargetOS, CardFormat);
            }
            else
            {
                MessageBox.Show("The listbox is empty");
            }

        }


        private void WriteCardControl_Load(object sender, EventArgs e)
        {
            string TargetOS = Properties.Settings.Default.TargetOS;
            SupportedCards.AddRange(CardsList);
            SupportedCards.Sort();
            switch(TargetOS)
            {
                case "Unix/Linux":
                    OSBox.SelectedIndex = 0;
                    break;
                case "Windows":
                    OSBox.SelectedIndex = 1;
                    break;
            }
            PathBox.Text = Properties.Settings.Default.WorkingDirectory;
            CardFormatBox.SelectedIndex = 1;
        }

        private void AddCards_Click(object sender, EventArgs e)
        {
            if (!AllCardsList.Items.Contains(CardNameBox.Text) && CardNameBox.Text!=null)
            {
                AllCardsList.Items.Add(CardNameBox.Text);
            }
        }

        private void btnAutofill_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string SheetName;
            foreach(Excel.Worksheet ws in wb.Worksheets)
            {
                SheetName = ws.Name;
                if (SupportedCards.Contains(SheetName) && !AllCardsList.Items.Contains(SheetName))
                {
                    AllCardsList.Items.Add(SheetName);
                }
            }
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //TextBox t = sender as TextBox;
            CardNameBox.AutoCompleteMode = AutoCompleteMode.Suggest;
            CardNameBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection collection = new AutoCompleteStringCollection();
            collection.AddRange(CardsList);
            this.CardNameBox.AutoCompleteCustomSource = collection;
        }

        private void AllCardsList_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            AllCardsList.Items.RemoveAt(AllCardsList.SelectedIndex);
        }

        private void BrowseBox_Click(object sender, EventArgs e)
        {
            this.PathBox.Text = General.BrowseFolder();
        }
    }
}
