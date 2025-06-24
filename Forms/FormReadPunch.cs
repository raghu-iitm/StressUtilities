using FEM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace StressUtilities.Forms
{
    public partial class FormReadPunch : Form
    {
        int PreSelectionSolType = -1;
        int PreSelectionRequest = -1;
        int PreSelectionTarget = -1;

        public FormReadPunch()
        {
            InitializeComponent();
        }

        private void FormReadPunch_Load(object sender, EventArgs e)
        {
            //string[] SolutionType;
            string[] TarEntity = new[] { "ELEMENTAL", "NODAL" };

            string[] SolutionType = new[] { "SOL 101", "SOL 103", "SOL 105", "SOL 106", "SOL 107", "SOL 108", "SOL 109", "SOL 110", "SOL 111", "SOL 112", "SOL 114", "SOL 115", "SOL 118", "SOL 129", "SOL 144", "SOL 145", "SOL 146", "SOL 153", "SOL 159", "SOL 200", "SOL 400", "SOL 600", "SOL 700" };


            // Try

            foreach (ComboBox ctrl in this.SelectionGroup.Controls.OfType<ComboBox>())
            {
                switch (ctrl.Name)
                {
                    case "SolutionTypeBox":
                        {
                            ctrl.Items.AddRange(SolutionType);
                            if (ctrl.SelectedIndex < 0)
                                ctrl.SelectedIndex = 0;
                            else
                                ctrl.SelectedIndex = PreSelectionSolType;
                            break;
                        }

                    case "RequestBox":
                        {
                            ctrl.Items.AddRange(ReadPunch.PunchRequestList());
                            if (ctrl.SelectedIndex < 0)
                                ctrl.SelectedIndex = 0;
                            else
                                ctrl.SelectedIndex = PreSelectionRequest;
                            break;
                        }

                    case "TargetEntityBox":
                        {
                            ctrl.Items.AddRange(TarEntity);
                            if (ctrl.SelectedIndex < 0)
                                ctrl.SelectedIndex = 0;
                            else
                                ctrl.SelectedIndex = PreSelectionRequest;
                            break;
                        }
                }
            }
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            //string[] FEFileList;
            //string FEFilePath;

            OpenFileDialog ofd = new OpenFileDialog();

            ofd.InitialDirectory = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
            ofd.Filter = @"Nastran Punch files (*.pch)|*.pch|All files (*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.Multiselect = true;
            ofd.RestoreDirectory = true;

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    string[] FEFileList = ofd.FileNames;
                    string FEFilePath = ofd.FileName.Replace(ofd.SafeFileName, "");
                    foreach (var filename in FEFileList)
                        FileListBox.Items.Add(filename);
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(@"Cannot read file from disk. Original error: " + Ex.Message);
                    return;
                }
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void BtnExtractPunch_Click(object sender, EventArgs e)
        {
            string Request;
            string ElemList = "";
            List<string> FileList = new List<string>();
            int i = 0;
            string EntityType;
            ReadPunch PunchResult = new ReadPunch();

            ElemList = EntityList.Text;
            Request = this.RequestBox.SelectedItem.ToString();
            EntityType = this.TargetEntityBox.SelectedItem.ToString();


            foreach (string filename in FileListBox.Items)
            {
                FileList.Add(filename);
                i += 1;
            }

            if (i == 0)
            {
                MessageBox.Show(@"No files are selected.");
                return;
            }

            if (Request == "")
            {
                MessageBox.Show(@"Please select the output request.");
                return;
            }

            PunchResult.ReadPunchResults(FileList, ElemList, Request, EntityType);
        }

        private void FileListBox_DoubleClick(object sender, System.EventArgs e)
        {
            if (FileListBox.SelectedIndex != -1)
                FileListBox.Items.RemoveAt(FileListBox.SelectedIndex);
        }

        private void SolutionTypeBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            PreSelectionSolType = SolutionTypeBox.SelectedIndex;
        }

        private void RequestBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            PreSelectionRequest = RequestBox.SelectedIndex;
        }
        private void TargetEntityBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            PreSelectionTarget = TargetEntityBox.SelectedIndex;
        }

        private void SolutionTypeBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void TargetEntityBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TargetEntityBox.SelectedIndex == 0)
                this.EntityList.Text = @"elm ";
            else
                this.EntityList.Text = @"node ";
        }
    }
}
