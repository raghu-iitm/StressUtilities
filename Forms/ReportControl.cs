using Report;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace StressUtilities.Forms
{
    public partial class ReportControl : UserControl
    {
        public ReportControl()
        {
            InitializeComponent();
            this.Load += new EventHandler(ReportForm_Load);
        }

        private void FillForm_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                Excel.Worksheet ReportSheet;
                ReportSheet = wb.Worksheets[WriteReport.GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT)];

                this.PathBox.Text = ReportSheet.Range[WriteReport.GetCellNameReport(CellNameReport.NAME_CELL_PATH_NAME)].Text;
                this.FileBox.Text = ReportSheet.Range[WriteReport.GetCellNameReport(CellNameReport.NAME_CELL_REPORT_NAME)].Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("No data stored");
            }
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

        private void BrowseBox_Click(object sender, EventArgs e)
        {
            this.PathBox.Text = General.BrowseFolder();
        }

        private void ReportForm_Load(object sender, EventArgs e)
        {
            //FileBox.Text= Properties.Settings.Default.MaxHDFRows.ToString();
            //PathBox.Text=Properties.Settings.Default.WorkingDirectory;


        }
    }
}
