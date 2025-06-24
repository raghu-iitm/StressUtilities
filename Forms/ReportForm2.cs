using Report;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace StressUtilities.Forms
{
    public partial class ReportForm2 : Form
    {
        public ReportForm2()
        {
            InitializeComponent();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ReportSheet;
            string Path, ReportName;
            WriteReport WR = new WriteReport();

            ReportSheet = wb.Worksheets[WriteReport.GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT)];

            Path = this.PathBox.Text;
            ReportName = this.FileBox.Text;

            ReportSheet.Range[WriteReport.GetCellNameReport(CellNameReport.NAME_CELL_PATH_NAME)].Value = this.PathBox.Text;
            ReportSheet.Range[WriteReport.GetCellNameReport(CellNameReport.NAME_CELL_REPORT_NAME)].Value = this.FileBox.Text;

            WR.GenerateReport();

        }

        private void BrowseIcon_Click(object sender, EventArgs e) 
        {
            this.PathBox.Text=General.BrowseFolder();
        }
        /*{
            string MapFilePath;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = wb.Path,
                Filter = string.Empty,
                Multiselect = false,
                RestoreDirectory = true,
                Title = "Select Folder...",
                CheckFileExists = false,
                CheckPathExists = false,
                FileName = "dummy"
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    MapFilePath = openFileDialog1.FileName.Replace(openFileDialog1.SafeFileName, "");

                    this.PathBox.Text = MapFilePath;
                }
                catch (Exception Ex)
                {
                    MessageBox.Show("Cannot read file from disk. Original error: " + Ex.Message);
                    return;
                }
            }
        }*/


        private void ReportForm_Load(object sender, EventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ReportSheet;
            ReportSheet = wb.Worksheets[WriteReport.GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT)];

            this.PathBox.Text = ReportSheet.Range[WriteReport.GetCellNameReport(CellNameReport.NAME_CELL_PATH_NAME)].Text;
            this.FileBox.Text = ReportSheet.Range[WriteReport.GetCellNameReport(CellNameReport.NAME_CELL_REPORT_NAME)].Text;
        }
    }
}
