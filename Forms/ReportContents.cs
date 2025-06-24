using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace StressUtilities.Forms
{
    public partial class ReportContents : Form
    {
        public ReportContents()
        {
            InitializeComponent();
            ToolTipTbl.SetToolTip(SummaryParam, $"Use vertical bar {"|"} to split the tables.");
        }

        private void DeleteBox_Click(object sender, EventArgs e)
        {
            Report.WriteReport.DeleteTableListData();
        }

        private void AddSelectedBox_Click(object sender, EventArgs e)
        {
            bool Status = true;
            string TableID = this.cmbTableList.SelectedItem.ToString();
            if (TableID != "No Table Data")
                Status = AddTablesData(TableID);
        }


        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void ReportContents_Load(object sender, EventArgs e)
        {
            Excel.Worksheet ReportSheet;
            //string[] TableList;
            string IDTable;
            int count = 0;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            int i = 1;
            //Excel.Range TableRange;
            string SummaryTables = Report.WriteReport.GetCellNameReport(Report.CellNameReport.NAME_CELL_START_TABLE);

            ReportSheet = wb.Worksheets[Report.WriteReport.GetSheetNameReport(Report.SheetNameReport.SHEET_NAME_REPORT)];

            IDTable = ReportSheet.Range["IDTable"].Text;

            foreach (ComboBox ctrl in this.GroupBox2.Controls.OfType<ComboBox>())
            {
                if (ctrl.Name == "cmbTableList")
                {
                    string[] TableList = Report.WriteReport.GetTableListData().Split(';');
                    ctrl.Items.AddRange(TableList);
                    if (TableList[0] != "No Table Data")
                    {
                        try
                        {
                            while (IDTable != TableList[count])
                                count += 1;
                        }
                        catch (Exception ex)
                        {
                            count = 0;
                        }
                    }
                    ctrl.SelectedIndex = count;
                }
            }

            Excel.Range TableRange = ReportSheet.Range[SummaryTables].End[Excel.XlDirection.xlDown];

            if (!string.IsNullOrEmpty(IDTable))
            {
                // IDTable = "No Table Data"
                foreach (Excel.Range Rng in TableRange)
                {
                    if (string.IsNullOrEmpty(Rng.Text) & Rng.Text != IDTable)
                        i += 1;
                }
            }
            else
            {
            }


            // If Not String.IsNullOrEmpty(IDTable) Then
            // While Not String.IsNullOrEmpty(ReportSheet.Range[SummaryTables].Offset[i, 0].Value)
            // If Not ReportSheet.Range[SummaryTables].Offset[i, 0].Value = IDTable And Not ReportSheet.Range[SummaryTables].Offset[i, 0].Value = "Header Level" Then
            // i += 1
            // End If
            // End While
            // End If

            foreach (TextBox tbx in this.GroupBox2.Controls.OfType<TextBox>())
            {
                switch (tbx.Name)
                {
                    case "CaptionBox":
                        {
                            tbx.Text = ReportSheet.Range[SummaryTables].Offset[i, 4].Text;
                            break;
                        }

                    case "Parameters":
                        {
                            tbx.Text = ReportSheet.Range[SummaryTables].Offset[i, 1].Text;
                            break;
                        }

                    case "CriticalItem":
                        {
                            tbx.Text = ReportSheet.Range[SummaryTables].Offset[i, 3].Text;
                            break;
                        }

                    case "SummaryParam":
                        {
                            tbx.Text = ReportSheet.Range[SummaryTables].Offset[i, 2].Text;
                            break;
                        }

                    case "ReportName":
                        {
                            tbx.Text = ReportSheet.Range["ReportName"].Text;
                            break;
                        }

                    case "SectionBox":
                        {
                            break;
                        }
                }
            }
        }



        // Private Sub WriteReport_Click(sender As Object, e As EventArgs) Handles WriteReport.Click
        // Dim TableId As String
        // Dim Component As String
        // Dim Parameters As String
        // Dim CritItem As String
        // Dim SummaryParam As String
        // Dim TableOption As String
        // Dim ReportName As String, AnalysisType As String
        // Dim ReportSheet As Excel.Worksheet
        // Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook


        // ReportSheet = wb.Worksheets(GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT))

        // Component = Me.CaptionBox.Text
        // TableId = Me.cmbTableList.SelectedItem
        // Parameters = Me.Parameters.Text
        // CritItem = Me.CriticalItem.Text
        // SummaryParam = Me.SummaryParam.Text
        // TableOption = SumTableOption()
        // ReportName = Me.ReportName.Text
        // AnalysisType = Me.SectionBox.Text


        // ReportSheet.Range("ReportName").Value = ReportName
        // 'ReportSheet.Range("Component").Value = Component
        // ReportSheet.Range("IDTable").Value = TableId
        // 'ReportSheet.Range("Parameters").Value = Parameters
        // 'ReportSheet.Range("CriticalCase").Value = CritItem
        // 'ReportSheet.Range("RptOptions").Value = TableOption
        // 'ReportSheet.Range("Summary").Value = SummaryParam
        // 'ReportSheet.Range("AnalysisType").Value = AnalysisType
        // GenerateReport()
        // End Sub



        private string SumTableOption()
        {
            string TableOption;
            if (this.optCalc.Checked == true)
                TableOption = "Calculation";
            else if (this.optCalcTbl.Checked == true)
                TableOption = "Calculation+Table";
            else if (this.optTbl.Checked == true)
                TableOption = "Table";
            else
                TableOption = "Calculation";

            return TableOption;
        }

        private void btnPictures_Click(object sender, EventArgs e)
        {
            string[] DrawingFileList;
            string FileExt;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            bool SelectXL = ChartOption();
            //int Count;
            //string ChartItem;

            if (SelectXL)
            {
                List<string> ChartList;
                ChartList = Report.WriteReport.GetChartList();
                foreach (string ChartItem in ChartList)
                    ListBoxGeometry.Items.Add(ChartItem);
            }
            else
            {
                openFileDialog1.InitialDirectory = General.GetFolderpath();
                openFileDialog1.Filter = @"Image Files |*.jpg;*.png;*.bmp;*.tif;*.tiff|JPEG Files (*.jpg)|*.jpg|Portable Network Graphics (*.png)|*.png|TIFF Files (*.tif;*.tiff)|*.tif;*.tiff|Bitmap Files (*.bmp)|*.bmp";
                openFileDialog1.FilterIndex = 1;
                openFileDialog1.Multiselect = true;
                openFileDialog1.RestoreDirectory = true;
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        DrawingFileList = openFileDialog1.FileNames;
                        foreach (var filename in DrawingFileList)
                        {
                            FileExt = Path.GetExtension(filename);
                            if (FileExt == ".png" | FileExt == ".jpg" | FileExt == ".tif" | FileExt == ".tiff" | FileExt == ".bmp")
                                ListBoxGeometry.Items.Add(filename);
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show($"Cannot read file from disk. Original error: {Ex.Message} ");
                        return;
                    }
                }
            }
        }

        private bool ChartOption()
        {
            bool SelectXL = false;
            if (this.OptionCharts.Checked == true)
                SelectXL = true;
            return SelectXL;
        }

        private void ListBoxGeometry_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            ListBoxGeometry.SelectedIndex = ListBoxGeometry.IndexFromPoint(new Point(e.X, e.Y));
        }

        private void ListBoxGeometry_MouseWheel(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (ListBoxGeometry.SelectedIndex == -1 & ListBoxGeometry.Items.Count != 0)
            {
                ListBoxGeometry.SelectedIndex = 0;
                return;
            }
            if (e.Delta < 0)
            {
                if (ListBoxGeometry.SelectedIndex == ListBoxGeometry.Items.Count - 1)
                    return;
                ListBoxGeometry.SelectedIndex = ListBoxGeometry.SelectedIndex + 1;
            }
            else
            {
                if (ListBoxGeometry.SelectedIndex == 0)
                    return;
                ListBoxGeometry.SelectedIndex = ListBoxGeometry.SelectedIndex - 1;
            }
        }

        private void ListBoxGeometry_DoubleClick(object sender, System.EventArgs e)
        {
            if (ListBoxGeometry.SelectedIndex != -1)
                ListBoxGeometry.Items.RemoveAt(ListBoxGeometry.SelectedIndex);
        }

        private void BtnAddtoList_Click(object sender, EventArgs e)
        {
            bool Status = true;
            string[] TableList;
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            TableList = Report.WriteReport.GetTableListData().Split(';');
            if (TableList[0] != "No Table Data")
            {
                for (var i = 0; i < TableList.Length; i++)
                {
                    if (Status)
                        Status = AddTablesData(TableList[i], true);
                }
            }

            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        }

        private bool AddTablesData(string TableId, bool Multiple = false)
        {
            // Dim TableId As String
            string CaptionName;
            string Parameters;
            string CritItem;
            string SummaryParam;
            //string ReportName;
            // Dim TableOption As String
            string SectionHeading;
            Excel.Worksheet ReportSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string[] FileList = { };
            int i = 1;
            int j = 1;
            bool Status = false;
            //bool CheckOveralp = false;
            DialogResult Results;
            string SummaryTables = Report.WriteReport.GetCellNameReport(Report.CellNameReport.NAME_CELL_START_TABLE);

            ReportSheet = wb.Worksheets[Report.WriteReport.GetSheetNameReport(Report.SheetNameReport.SHEET_NAME_REPORT)];
            //ReportName = this.ReportName.Text;

            if (!Multiple)
            {
                CaptionName = this.CaptionBox.Text;
                // TableId = Me.cmbTableList.SelectedItem
                Parameters = this.Parameters.Text;
                CritItem = this.CriticalItem.Text;
                SummaryParam = this.SummaryParam.Text;
                // TableOption = SumTableOption()
                // ReportName = Me.ReportName.Text
                SectionHeading = this.SectionBox.Text;
            }
            else
            {
                CaptionName = "";
                Parameters = Report.WriteReport.GetParameterList(TableId, CaptionName);
                // TableId = Me.cmbTableList.SelectedItem
                CritItem = this.CriticalItem.Text;
                SummaryParam = Parameters;
            }

            while (ReportSheet.Range[SummaryTables].Offset[j, 0].Text != "Header Level")
                j += 1;

            while (!string.IsNullOrEmpty(ReportSheet.Range[SummaryTables].Offset[i, 0].Text))
            {
                if (TableId == ReportSheet.Range[SummaryTables].Offset[i, 0].Text)
                {
                    Results = MessageBox.Show($"The {TableId} parameters are already populated. Are you sure to overwrite the parameter? ", "Warning!", MessageBoxButtons.YesNo);
                    if (Results == DialogResult.Yes)
                    {
                        ReportSheet.Range[SummaryTables].Offset[i, 0].Value = TableId;
                        ReportSheet.Range[SummaryTables].Offset[i, 1].Value = Parameters;
                        ReportSheet.Range[SummaryTables].Offset[i, 2].Value = SummaryParam;
                        ReportSheet.Range[SummaryTables].Offset[i, 3].Value = CritItem;
                        ReportSheet.Range[SummaryTables].Offset[i, 4].Value = CaptionName;
                    }
                    Status = true;
                }
                i += 1;
            }

            if (!Status)
            {
                if (i < j - 4)
                {
                    ReportSheet.Range[SummaryTables].Offset[i, 0].Value = TableId;
                    ReportSheet.Range[SummaryTables].Offset[i, 1].Value = Parameters;
                    ReportSheet.Range[SummaryTables].Offset[i, 2].Value = SummaryParam;
                    ReportSheet.Range[SummaryTables].Offset[i, 3].Value = CritItem;
                    ReportSheet.Range[SummaryTables].Offset[i, 4].Value = CaptionName;
                    Status = true;
                }
                else
                {
                    MessageBox.Show(@"The number of filled rows in the list of tables has reached the Reports content table. Please insert additional rows before proceeding.");
                    Status = false;
                }
            }
            return Status;
        }

        private void cmbTableList_SelectedIndexChanged(object sender, EventArgs e)
        {
            string TableId;
            string Parameters;
            string CaptionText = "";

            TableId = this.cmbTableList.Text;
            if (TableId != "No Table Data")
            {
                Parameters = Report.WriteReport.GetParameterList(TableId, CaptionText);
                this.Parameters.Text = Parameters;
                this.SummaryParam.Text = Parameters;
                this.CaptionBox.Text = CaptionText;
            }
            else
                MessageBox.Show(@"No valid table ID selected.");
        }

        private void BtnPicturesList_Click(object sender, EventArgs e)
        {
            string[] FileList = { };
            int i; // , j As Integer
            Excel.Worksheet ReportSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string CaptionPrefix;
            string PictureID;
            string FileNameList;
            DialogResult Results;
            bool Status = false;
            string ObjType = "Picture";
            string listSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ListSeparator;
            string StartCellName = Report.WriteReport.GetCellNameReport(Report.CellNameReport.NAME_CELL_START_FIGURE);

            if (ListBoxGeometry.Items.Count > 0)
            {
                FileList = new string[ListBoxGeometry.Items.Count - 1 + 1];
                for (i = 0; i < ListBoxGeometry.Items.Count; i++)
                    FileList[i] = ListBoxGeometry.Items[i].ToString();
            }

            if (ChartOption() == true)
                ObjType = "xlChart";

            ReportSheet = wb.Worksheets[Report.WriteReport.GetSheetNameReport(Report.SheetNameReport.SHEET_NAME_REPORT)];

            PictureID = this.PictureID.Text;
            FileNameList = string.Join(listSeparator, FileList);
            CaptionPrefix = this.CaptionText.Text;
            SummaryParam = Parameters;

            i = 1;
            while (!string.IsNullOrEmpty(ReportSheet.Range[StartCellName].Offset[i, 0].Text))
            {
                if (PictureID == ReportSheet.Range[StartCellName].Offset[i, 0].Text)
                {
                    Results = MessageBox.Show($"The {PictureID} parameters are already populated. Are you sure to overwrite the parameter? ", "Warning!", MessageBoxButtons.YesNo);
                    if (Results == DialogResult.Yes)
                    {
                        ReportSheet.Range[StartCellName].Offset[i, 0].Value = PictureID;
                        ReportSheet.Range[StartCellName].Offset[i, 1].Value = ObjType;
                        ReportSheet.Range[StartCellName].Offset[i, 2].Value = FileNameList;
                        ReportSheet.Range[StartCellName].Offset[i, 3].Value = CaptionPrefix;
                    }
                    Status = true;
                }
                i += 1;
            }

            if (!Status)
            {
                ReportSheet.Range[StartCellName].Offset[i, 0].Value = PictureID;
                ReportSheet.Range[StartCellName].Offset[i, 1].Value = ObjType;
                ReportSheet.Range[StartCellName].Offset[i, 2].Value = FileNameList;
                ReportSheet.Range[StartCellName].Offset[i, 3].Value = CaptionPrefix;
                Status = true;
            }
        }


    }
}
