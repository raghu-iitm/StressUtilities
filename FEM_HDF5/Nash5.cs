using FEM;
using StressUtilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Nastranh5
{
    public partial class Nash5 : Form
    {
        public Nash5()
        {
            InitializeComponent();
            InitializeTreeView();
            InitializeFormComponents();
        }

        public object My { get; private set; }


        private void InitializeTreeView()
        {
            treeView1.BeginUpdate();
            treeView1.Nodes.Clear();
            treeView1.EndUpdate();
            BtnAllNodeExpand.CheckOnClick = true;
            BtnAllNodeExpand.CheckedChanged += new EventHandler(BtnAllNodeExpand_CheckedChanged);
        }

        private void InitializeFormComponents()
        {
            string[] items = new string[] { "Derive/Average", "Average/Derive" };
            cBOptions.DataSource = items;
            cBOptions.SelectedIndex = 0;

            items = new string[] { "Centroid", "Gauss Point" }; //,"Nodal"
            cBLocation.DataSource = items;
            cBLocation.SelectedIndex = 0;

            toolStripStatusLabel1.Text = "Ready...";
        }

        private void btnDBExtract_Click(object sender, EventArgs e)
        {
            ExtractResults();
        }

        private void ExtractResults()
        {
            H5DBread readh5 = new H5DBread();
            List<string> h5FileList = new List<string>();
            bool Success = true;
            toolStripStatusLabel1.Text = @"Extracting the database results...";
            Stopwatch sw = new Stopwatch();
            sw.Start();

            //string grpName;
            try
            {
                /*foreach (string FileItem in FileListBox.Items)
                {
                    h5FileList.Add(FileItem);
                }*/
                h5FileList.AddRange(FileListBox.Items.Cast<string>().ToArray());
                if (h5FileList.Count == 0)
                {
                    MessageBox.Show(@"The database list is empty.");
                    return;
                }

                string datasetpath = textBoxDataSet.Text;
                if (string.IsNullOrEmpty(datasetpath))
                {
                    MessageBox.Show(@"Please select the dataset from the tree");
                    return;
                }

                string datasetName = datasetpath.Substrin­g(datasetpath.LastI­n­dexOf("/") + 1);
                string datasetparent = treeView1.SelectedNode.Parent.Text;
                string Prefix = datasetparent + "_" + datasetName;
                string RangeStart = StartCell.Text;
                string grpName = datasetpath.Substring(0, datasetpath.Length - datasetName.Length - 1);

                //entity list needs to be reworked since a larger entity list has overhead and slows down the code.
                List<long> entitylist = new List<long>();
                List<long> SubCaseList = new List<long>();
                if (grpName.StartsWith("/NASTRAN/RESULT") && !datasetpath.EndsWith("/DOMAINS"))
                {
                    entitylist = General.GetEntityList(textBoxEntity.Text);
                    SubCaseList = General.GetEntityList(SubCaseBox.Text);
                }

                    bool[] Requests = new bool[] { PrincRequest(), vMRequest() };
                string LocRequest = RequestLocation();
                readh5.ExtractH5File(ref h5FileList, grpName, ref datasetName, ref datasetparent, ref entitylist, ref SubCaseList, Requests, ref RangeStart, ref LocRequest, ref Success);
                if (!Success)
                    //MessageBox.Show($"The database is extracted Successfully. Elapsed Time: {sw.Elapsed/*:00.00.00.0000*/}");
                    //else */
                    MessageBox.Show($"Not all entity results are extracted. Please refer to log file at the Workbook location");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            toolStripStatusLabel1.Text = "Ready...";
            sw.Reset();
            readh5.Dispose();
        }

        public void InsertTree(string fileName)
        {
            ImageList h5ImageList = new ImageList();
            h5ImageList.Images.Add(StressUtilities.Properties.Resources.Database);
            h5ImageList.Images.Add(StressUtilities.Properties.Resources.Folder);
            h5ImageList.Images.Add(StressUtilities.Properties.Resources.Dataset);

            TreeView trlist = new TreeView();

            List<string> datasetList = trlist.GetGroupList(fileName);

            Dictionary<string, NodeEntry> tItem = trlist.GetGroups(datasetList);


            TreeNode rootNode = new TreeNode("db");
            rootNode.Tag = "file";
            treeView1.ImageList = h5ImageList;
            rootNode.ImageIndex = 0;
            rootNode.SelectedImageIndex = 0;
            treeView1.Nodes.Add(rootNode);
            TreeNode ParentNode = rootNode;

            PopulateTreeView(tItem, ParentNode);

            h5ImageList = null;

        }

        private void PopulateTreeView(Dictionary<string, NodeEntry> tItem, TreeNode parentNode)
        {
            TreeNode TempNode = parentNode;
            TreeNode ChildNode;
            Dictionary<string, NodeEntry> childDict;
            foreach (string item in tItem.Keys)
            {
                if (tItem[item].Children.Count > 0)
                {
                    childDict = new Dictionary<string, NodeEntry>();
                    childDict = tItem[item].Children;
                    ChildNode = new TreeNode(item);
                    PopulateTreeView(childDict, ChildNode);

                    parentNode = TempNode;
                    ChildNode.Tag = "group";
                    ChildNode.ImageIndex = 1;
                    ChildNode.SelectedImageIndex = 1;
                    parentNode.Nodes.Add(ChildNode);
                }
                else
                {
                    ChildNode = new TreeNode(item);
                    ChildNode.Tag = "dataset";
                    ChildNode.ImageIndex = 2;
                    ChildNode.SelectedImageIndex = 2;
                    parentNode.Nodes.Add(ChildNode);
                }
            }
        }

        private void btnOpenDB_Click(object sender, EventArgs e) => Openh5db();



        private void Openh5db()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            string[] FEFileList;
            dlg.InitialDirectory = StressUtilities.Properties.Settings.Default.DBpath;
            dlg.Filter = @"Nastran h5 database (*.h5)|*.h5";
            dlg.FilterIndex = 1;
            dlg.Multiselect = true;
            dlg.RestoreDirectory = true;

            toolStripStatusLabel1.Text = @"Navigating Database";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                //string fileName;
                //fileName = dlg.FileName;
                treeView1.Nodes.Clear();
                try
                {
                    FEFileList = dlg.FileNames;

                    foreach (string filename in FEFileList)
                    {
                        FileListBox.Items.Add(filename);
                    }
                    string StartfileName = FileListBox.Items[0].ToString();
                    string FEFilePath = Path.GetDirectoryName(StartfileName);

                    InsertTree(StartfileName); //Function to insert the trees
                    StressUtilities.Properties.Settings.Default.DBpath = FEFilePath;
                    StressUtilities.Properties.Settings.Default.Save();
                    treeView1.ExpandAll();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Cannot read file from disk. Original error: {ex.Message} ");
                    toolStripStatusLabel1.Text = @"Ready...";
                    return;
                }
            }
            toolStripStatusLabel1.Text = "Ready...";
        }

        private void FileListBox_MouseWheel(object sender, MouseEventArgs e)
        {
            if (FileListBox.SelectedIndex == -1)
            {
                FileListBox.SelectedIndex = 0;
                return;
            }
            if (e.Delta < 0)
            {
                if (FileListBox.SelectedIndex == FileListBox.Items.Count - 1) return;
                FileListBox.SelectedIndex = FileListBox.SelectedIndex + 1;
            }
            else
            {
                if (FileListBox.SelectedIndex == 0) return;
                FileListBox.SelectedIndex = FileListBox.SelectedIndex - 1;
            }
        }

        private void FileListBox_MouseDoubleClick(Object sender, MouseEventArgs e)
        {
            if (FileListBox.SelectedIndex != -1)
            {
                FileListBox.Items.RemoveAt(FileListBox.SelectedIndex);
                if (FileListBox.Items.Count == 0)
                {
                    treeView1.Nodes.Clear();
                }
            }
        }

        private void btnCloseDB_Click(object sender, EventArgs e)
        {
            dbclose();
        }

        private void dbclose()
        {
            if (FileListBox.SelectedIndex != -1)
            {
                FileListBox.Items.RemoveAt(FileListBox.SelectedIndex);
                if (FileListBox.Items.Count == 0)
                {
                    treeView1.Nodes.Clear();
                }
            }
        }

        private void btnCloseAllDB_Click(object sender, EventArgs e)
        {
            FileListBox.Items.Clear();
            treeView1.Nodes.Clear();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                int myNodeCount = treeView1.SelectedNode.GetNodeCount(true);
                treeView1.PathSeparator = "/";
                string datasetpath = treeView1.SelectedNode.FullPath;
                datasetpath = datasetpath.Replace("db", "");
                if ((string)treeView1.SelectedNode.Tag == "dataset") { textBoxDataSet.Text = datasetpath; }
                else { textBoxDataSet.Text = null; }

                if (treeView1.SelectedNode.Parent != null)
                {
                    string tnName = treeView1.SelectedNode.Parent.Text;

                    if (tnName == "STRESS" && !datasetpath.Contains("INDEX"))
                    {
                        vMRequestBox.Enabled = true;
                        if (!datasetpath.Contains("CPLX"))
                            PrincRequestBox.Enabled = true;
                        else
                            PrincRequestBox.Enabled = false;
                    }
                    else if (tnName == "STRAIN" && !datasetpath.Contains("INDEX"))
                    {
                        vMRequestBox.Enabled = false;
                        if (!datasetpath.Contains("CPLX"))
                            PrincRequestBox.Enabled = true;
                        else
                            PrincRequestBox.Enabled = false;
                    }
                    else
                    {
                        vMRequestBox.Enabled = false;
                        PrincRequestBox.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                //Pass  
            }
        }

        //private void btnBrowse_Click(object sender, EventArgs e)
        //{
        //    toolStripStatusLabel1.Text = "Busy";
        //    using (var ofd = new OpenFileDialog())
        //    {
        //        ofd.InitialDirectory = StressUtilities.Properties.Settings.Default.WorkingDirectory;
        //        ofd.Multiselect = false;
        //        ofd.RestoreDirectory = true;
        //        ofd.ValidateNames = false;
        //        ofd.CheckFileExists = false;
        //        ofd.CheckPathExists = true;
        //        ofd.FileName = "Select Folder.";

        //        DialogResult result = ofd.ShowDialog();

        //    if (result == DialogResult.OK)
        //        {
        //            try
        //            {
        //                string FEFilepath = Path.GetDirectoryName(ofd.FileName);

        //                //textBoxDirectory.Text = FEFilepath;
        //                StressUtilities.Properties.Settings.Default.WorkingDirectory= FEFilepath;
        //            }
        //    catch( Exception Ex)
        //        {
        //            MessageBox.Show(string.Format("Cannot read file from disk. Original error: {0} ", Ex.Message));
        //            return;
        //        }
        //    }
        //        toolStripStatusLabel1.Text = "Ready...";
        //    }

        //    StressUtilities.Properties.Settings.Default.Save();
        //}


        private bool vMRequest()
        {
            bool Request = false;

            if (vMRequestBox.Checked && vMRequestBox.Enabled)
            {
                Request = true;
            }
            return Request;
        }

        private bool PrincRequest()
        {
            bool Request = false;
            if (PrincRequestBox.Checked && PrincRequestBox.Enabled)
            {
                Request = true;
            }
            return Request;
        }

        private string RequestLocation()
        {
            string request = cBLocation.Text;
            return request;
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Openh5db();
        }

        private void singleFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExtractResults();
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dbclose();
        }

        private void closeAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileListBox.Items.Clear();
            treeView1.Nodes.Clear();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            //H5General test = new H5General();
            toolStripStatusLabel1.Text = "Busy...";
            //test.looptime();
            //toolStripStatusLabel1.Text = "Ready...";
        }

        private void btnLCTemplate_Click(object sender, EventArgs e)
        {
            LCTable InstLCtbl = new LCTable();
            InstLCtbl.LCTableTemplate();
        }

        private void btnLCFile_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = @"Select the Load Case Combination File...";
            using (var ofd = new OpenFileDialog())
            {
                ofd.InitialDirectory = StressUtilities.Properties.Settings.Default.WorkingDirectory;
                ofd.Filter = @"Excel files (*.xlsx)|*.xlsx";
                ofd.FilterIndex = 1;
                ofd.Multiselect = false;
                ofd.RestoreDirectory = true;
                ofd.ValidateNames = false;
                ofd.CheckFileExists = false;
                ofd.CheckPathExists = true;

                DialogResult result = ofd.ShowDialog();

                if (result == DialogResult.OK)
                {
                    try
                    {
                        textBoxLC.Text = ofd.FileName; // FEFilepath;
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show($"Cannot read file from disk. Original error: {Ex.Message} ");
                        return;
                    }
                }
                toolStripStatusLabel1.Text = @"Ready...";
            }

            StressUtilities.Properties.Settings.Default.Save();
        }

        //private void ExtractionBox_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (ExtractionBox.Text == "CSV")
        //    {
        //        cBFileOption.Enabled = true;
        //        cBFileOption.Text = "Single File";
        //        textBoxDirectory.Enabled = true;
        //        btnBrowse.Enabled = true;
        //    }
        //    else
        //    {
        //        cBFileOption.Enabled = false;
        //        textBoxDirectory.Enabled = false;
        //        btnBrowse.Enabled = false;
        //    }

        //}



        //            if (btnExpand.CheckOnClick == true)
        //    treeView1.ExpandAll();
        //else
        //    treeView1.CollapseAll();


        private void BtnAllNodeExpand_CheckedChanged(object sender, EventArgs e)
        {
            if (BtnAllNodeExpand.Checked)
            {
                treeView1.ExpandAll();
            }
            else
            {
                treeView1.CollapseAll();
            }
        }

        private void BtnNodeExpand_Click(object sender, EventArgs e)
        {
            if (treeView1.Nodes.Count > 0)
            {
                if (treeView1.SelectedNode.IsSelected)
                    treeView1.SelectedNode.ExpandAll();
            }

        }

        private void SelectCellBox_Click(object sender, EventArgs e)
        {
            //string oldrange = StartCell.Text;
            this.SendToBack();
            try
            {
                Microsoft.Office.Interop.Excel.Range RangeStartCell = Globals.ThisAddIn.Application.InputBox("Start Cell", "Select Start Cell", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type: 8); //e.Range[Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1];
                StartCell.Text = RangeStartCell.Address.ToString();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RangeStartCell);
            }
            catch
            {
                //StartCell.Text = "B2";
            }

            //this.BringToFront();

            this.Focus();

        }
    }
}
