using FEM;
using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace StressUtilities.Forms
{
    public partial class CombinationForm : Form
    {
        public CombinationForm()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnMapFile_Click(object sender, EventArgs e)
        {
            // Dim FEFileList() As String
            //string MapFilePath;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            openFileDialog1.InitialDirectory = wb.Path;
            openFileDialog1.Filter = @"Mapping Files (*.smp)|*.smp|Text Files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    string MapFilePath = openFileDialog1.FileName;
                    // FEFilePath = Replace(openFileDialog1.FileName, openFileDialog1.SafeFileName, "")

                    TextBoxMapFile.Text = MapFilePath;
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(@"Cannot read file from disk. Original error: " + Ex.Message);
                    return;
                }
            }
        }


        private void BrowseBox_Click(object sender, EventArgs e)
        {
            //string[] FEFileList;
            //string FEFilePath;
            string DataSource = DataSourceOption();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            openFileDialog1.InitialDirectory = wb.Path;
            openFileDialog1.Filter = @"Patran rpt files (*.rpt)|*.rpt|Comma Separated values (*.csv)|*.csv|Nastran HDF5 Files (*.h5)|*.h5|All files (*.*)|*.*";
            switch (DataSource)
            {
                case "rpt":
                    {
                        openFileDialog1.FilterIndex = 1;
                        break;
                    }

                case "csv":
                    {
                        openFileDialog1.FilterIndex = 2;
                        break;
                    }
                case "h5":
                    {
                        openFileDialog1.FilterIndex = 2;
                        break;
                    }

                default:
                    {
                        openFileDialog1.FilterIndex = 3;
                        break;
                    }
            }

            openFileDialog1.Multiselect = true;
            openFileDialog1.RestoreDirectory = true;


            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    string[] FEFileList = openFileDialog1.FileNames;
                    string FEFilePath = openFileDialog1.FileName.Replace(openFileDialog1.SafeFileName, "");
                    foreach (var filename in FEFileList)

                        FileListBox.Items.Add(filename.Replace(FEFilePath, ""));

                    FEPathBox.Text = FEFilePath;
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(@"Cannot read file from disk. Original error: " + Ex.Message);
                    return;
                }
            }
        }

        private string ElemTypeSelection()
        {
            if (this.SelTypeOption1D.Checked == true)
                return "1D";
            else if (this.SelTypeOption2D.Checked == true)
                return "2D";
            else if (this.SelTypeOption3D.Checked == true)
                return "3D";
            else if (SelTypeOptionNode.Checked == true)
                return "NODE";
            else
                return "NONE";
        }

        private string OperationOption()
        {
            if (this.OpOptionCombine.Checked == true)
                return "Combine";
            else if (this.OpOptionCombineAvg.Checked == true)
                return "CombineAverage";
            else
                return "NONE";
        }

        private string FormulaValueOption()
        {
            if (this.ValueOption.Checked == true)
                return "Value";
            else if (this.FormulaOption.Checked == true)
                return "Formula";
            else
                return "Value";
        }

        private string ThermalOption()
        {
            if (this.ThermalOptionBox.Checked == true)
                return "AddThermal";
            else
                return "NONE";
        }
        private string DataSourceOption()
        {
            if (this.DataSourceRPT.Checked == true)
                return "rpt";
            else if (this.DataSourceCSV.Checked == true)
                return "csv";
            else if (this.DataSourceh5.Checked == true)
                return "csv";
            else
                return "NONE";
        }

        private void FileListBox_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            FileListBox.SelectedIndex = FileListBox.IndexFromPoint(new Point(e.X, e.Y));
        }


        private void FileListBox_MouseWheel(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (FileListBox.SelectedIndex == -1)
            {
                FileListBox.SelectedIndex = 0;
                return;
            }
            if (e.Delta < 0)
            {
                if (FileListBox.SelectedIndex == FileListBox.Items.Count - 1)
                    return;
                FileListBox.SelectedIndex = FileListBox.SelectedIndex + 1;
            }
            else
            {
                if (FileListBox.SelectedIndex == 0)
                    return;
                FileListBox.SelectedIndex = FileListBox.SelectedIndex - 1;
            }
        }




        private void FileListBox_DoubleClick(object sender, System.EventArgs e)
        {
            if (FileListBox.SelectedIndex != -1)
                FileListBox.Items.RemoveAt(FileListBox.SelectedIndex);
        }

        private void btbBrowseLC_Click(object sender, EventArgs e)
        {
            //string FEFileList;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            openFileDialog1.InitialDirectory = wb.Path;
            openFileDialog1.Filter = @"Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = true;
            openFileDialog1.RestoreDirectory = true;


            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    string FEFileList = openFileDialog1.FileName;
                    LCFileBox.Text = FEFileList;
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(@"Cannot read file from disk. Original error: " + Ex.Message);
                    return;
                }
            }
        }

        private void btnCombLoads_Click(object sender, EventArgs e)
        {
            //string LCFileName, FilePath, LCTypesList, ThermCaseList, ElemType, OperationType;
            //string DataSource, ImportOption, ElmList;
            //string[] FileList;
            string MapFileName = "";
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;


            string FilePath = FEPathBox.Text;
            string LCFileName = LCFileBox.Text;
            string LCTypesList = LoadTypesBox.Text;
            string ThermCaseList = ThermalSourceBox.Text;
            string ElemType = ElemTypeSelection();
            string OperationType = OperationOption();
            string DataSource = DataSourceOption();
            string ImportOption = FormulaValueOption();
            string ElmList = ElemList.Text;
            string[] FileList = new string[FileListBox.Items.Count - 1 + 1];
            for (int i = 0; i <= FileListBox.Items.Count - 1; i++)
            {
                FileList[i] = FileListBox.Items[i].ToString();
            }

            string UnitThermalLoads = unitThermalBox.Text;

            // Check errors
            if (ElemType == "NONE")
            {
                MessageBox.Show(@"Please select the element type");
                return;
            }
            else if (OperationType == "NONE")
            {
                MessageBox.Show(@"Please identify the operation");
                return;
            }
            else if (DataSource == "NONE")
            {
                MessageBox.Show(@"Please select the datasource");
                return;
            }
            else if (FilePath == "")
            {
                MessageBox.Show(@"Please list the result files");
                return;
            }
            else if (LCFileName == "")
            {
                MessageBox.Show(@"Please identify the Load Case file");
                return;
            }
            else if (LCTypesList == "")
            {
                MessageBox.Show(@"At least one load source must be identified. Example 1,2,3,4 or 1;2;3;4");
                return;
            }
            if (OperationType == "CombineAverage" && ElmList.ToUpper() == "ALL")
            {
                MessageBox.Show(@"Please enter the list of elements seperated by Comma or semicolon or Line feed");
                return;
            }

            if (DataSource == "csv")
                MapFileName = TextBoxMapFile.Text;

            LoadCombination LCo = new LoadCombination();

            LCo.CombineLoads(FilePath, FileList, DataSource, LCFileName, LCTypesList, ThermCaseList, UnitThermalLoads, ElemType, ElmList, OperationType, ImportOption, MapFileName);

            //Marshal.ReleaseComObject(wb);
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            FEPathBox.Text = "";
            LCFileBox.Text = "";
            LoadTypesBox.Text = null;
            ThermalSourceBox.Text = "";
            ElemList.Text = "";
            DataSourceRPT.Checked = false;
            DataSourceCSV.Checked = false;
            ThermalOptionBox.Checked = false;
            OpOptionCombine.Checked = false;
            OpOptionCombineAvg.Checked = false;
            OpOptionCombine.Enabled = true;
            OpOptionCombineAvg.Enabled = true;
            ThermalOptionBox.Enabled = true;
            CellSelectBox.Enabled = false;
            TextBoxMapFile.Enabled = false;
            TextBoxMapFile.Text = "";
        }

        private void CellSelectBox_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Range Rng;
            try
            {
                WindowState = FormWindowState.Minimized;

                Rng = wb.Application.InputBox("Select the 1 to 3 cells containing the unit thermal loads.", "Obtain Range Object", Type: 8);
                WindowState = FormWindowState.Normal;
            }
            catch (Exception Ex)
            {
                Rng = null;
            }

            if (Rng != null)
                this.unitThermalBox.Text = Rng.AddressLocal;
        }

        private void SelTypeOption1D_CheckedChanged(object sender, EventArgs e)
        {
            if (SelTypeOption1D.Enabled == true)
            {
                OpOptionCombine.Enabled = true;
                OpOptionCombineAvg.Enabled = false;
                ThermalOptionBox.Enabled = true;
            }
        }

        private void SelTypeOption2D_CheckedChanged(object sender, EventArgs e)
        {
            if (SelTypeOption2D.Enabled == true)
            {
                OpOptionCombine.Enabled = true;
                OpOptionCombineAvg.Enabled = true;
                ThermalOptionBox.Enabled = true;
            }
        }

        private void SelTypeOption3D_CheckedChanged(object sender, EventArgs e)
        {
            if (SelTypeOption3D.Enabled == true)
            {
                OpOptionCombine.Enabled = true;
                OpOptionCombineAvg.Enabled = false;
                ThermalOptionBox.Enabled = false;
            }
        }

        private void SelTypeOptionNode_CheckedChanged(object sender, EventArgs e)
        {
            if (SelTypeOptionNode.Enabled == true)
            {
                OpOptionCombine.Enabled = true;
                OpOptionCombineAvg.Enabled = false;
                ThermalOptionBox.Enabled = true;
            }
        }

        private void ThermalOptionBox_CheckedChanged(object sender, EventArgs e)
        {
            if (ThermalOptionBox.Checked == true)
                CellSelectBox.Enabled = true;
            else if (ThermalOptionBox.Checked == false)
                CellSelectBox.Enabled = false;
        }

        private void DataSourceCSV_CheckedChanged(object sender, EventArgs e)
        {
            if (DataSourceCSV.Checked == true)
            {
                TextBoxMapFile.Enabled = true;
                btnMapFile.Enabled = true;
            }
            else
            {
                TextBoxMapFile.Enabled = false;
                btnMapFile.Enabled = false;
            }
        }



        /* private void BrowseBox_Click(object sender, EventArgs e)
         {

         }*/
    }
}
