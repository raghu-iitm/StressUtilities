using FEM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace StressUtilities.Forms
{
    public partial class ImportF06Form : Form
    {
        private int PreSelectionResult;

        public ImportF06Form()
        {
            InitializeComponent();
        }

        private void ImportF06Form_Load(object sender, EventArgs e)
        {
            string[]  ResultList = new[] { "STRESSES", "STRAINS", "FORCES", "DISPLACEMENTS", "SPC FORCES", "LOADS" };

            // Try

            foreach (ComboBox ctrl in this.Controls.OfType<ComboBox>())
            {
                if (ctrl.Name == "RequestBox")
                {
                    ctrl.Items.AddRange(ResultList);
                    if (ctrl.SelectedIndex < 0)
                        ctrl.SelectedIndex = 0;
                    else
                        ctrl.SelectedIndex = PreSelectionResult;
                }
            }
        }


        private void btnF06Add_Click(object sender, EventArgs e)
        {
            //string[] FEFileList;
            //string FEFilePath;
            // Dim DataSource As String = DataSourceOption()
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
            openFileDialog1.Filter = @"Nastran f06 files (*.f06)|*.f06|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = true;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    string[] FEFileList = openFileDialog1.FileNames;
                    string FEFilePath = openFileDialog1.FileName.Replace(openFileDialog1.SafeFileName, "");
                    foreach (var filename in FEFileList)
                        FilesListBox.Items.Add(filename);
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(@"Cannot read file from disk. Original error: " + Ex.Message);
                    return;
                }
            }
        }

        private void btnImportF06_Click(object sender, EventArgs e)
        {
            string Request;
            string ElemList = "";
            List<string> f06FileList = new List<string>();
            // Dim FilesListBox As ListBox
            int i = 0;

            ElemList = ElementList.Text;
            Request = this.RequestBox.SelectedItem.ToString();


            foreach (string filename in FilesListBox.Items)
            {
                f06FileList.Add(filename);
                i += 1;
            }

            if (i == 0)
            {
                MessageBox.Show("No files are selected.");
                return;
            }

            if (Request == "")
            {
                MessageBox.Show("Please select the result to be extracted.");
                return;
            }
            Readf06 ReadFiles = new Readf06();
            ReadFiles.f06Read(f06FileList, Request, ElemList);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void FilesListBox_DoubleClick(object sender, System.EventArgs e)
        {
            if (FilesListBox.SelectedIndex != -1)
                FilesListBox.Items.RemoveAt(FilesListBox.SelectedIndex);
        }

        private void RequestBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            PreSelectionResult = RequestBox.SelectedIndex;
        }
    }
}
