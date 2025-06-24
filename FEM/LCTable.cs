using StressUtilities;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FEM
{
    class LCTable
    {
        string _TemplateName;
        string _SheetName;

        private string TemplateName
        {
            get
            {
                return _TemplateName;
            }
            set
            {
                _TemplateName = "LC_Table.xlsx";
            }
        }
        private string LCSheetName
        {
            get
            {
                return _SheetName;
            }
            set
            {
                _SheetName = "LoadCases";
            }
        }

        public LCTable()
        {

        }
        public void LCTableTemplate()
        {
            //string LCTable_Path;
            //string Template_Path;

            string LcTablePath = StressUtilities.Properties.Settings.Default.WorkingDirectory; // @"C:\Temp";


            string TemplatePath = Path.Combine(LcTablePath, TemplateName);
            if (File.Exists(TemplatePath))
            {
                MessageBox.Show($"LC Template already exists in the folder {LcTablePath} \nPlease remove the old Template before generating the new one");
            }
            else
            {
                Excel.Workbook wb;
                Excel.Application xlApp = Globals.ThisAddIn.Application;
                wb = xlApp.Workbooks.Add();

                wb.Title = "LC Template";
                wb.Subject = "LC";
                wb.SaveAs(Filename: TemplatePath);


                Excel.Worksheet wsobj = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                wsobj.Name = LCSheetName;

                wsobj.Range["A1"].Value2 = "PLEASE DO NOT MODIFY THE FORMAT OF THE TABLE. THE INFORMATION REGARDING THE PROGRAM/PROJECT CAN BE MANUALLY ENTERED IN ROW 2. HEADER DESCRIPTION IN ROW 3 CAN BE CHANGED BUT NOT IN ROW 4";
                wsobj.Range["A3"].Value2 = "ID";
                wsobj.Range["B3"].Value2 = "Load Case";
                wsobj.Range["C3"].Value2 = "Factor 1";
                wsobj.Range["D3"].Value2 = "Additional Case 2";
                wsobj.Range["E3"].Value2 = "Factor 2";
                wsobj.Range["F3"].Value2 = "Additional Case 3";
                wsobj.Range["G3"].Value2 = "Factor 3";
                wsobj.Range["H3"].Value2 = "Additional Case 4";
                wsobj.Range["I3"].Value2 = "Factor 4";
                //wsobj.Range["J3"].Value = "Delta Temperature [deg. C)";
                wsobj.Range["J3"].Value2 = "Description";
                wsobj.Range["A4"].Value2 = "SID";
                wsobj.Range["A5"].Value2 = "1";
                wsobj.Range["A6"].Value2 = "2";
                wsobj.Range["B4"].Value2 = "LC1";
                wsobj.Range["C4"].Value2 = "LF1";
                wsobj.Range["D4"].Value2 = "LC2";
                wsobj.Range["E4"].Value2 = "LF2";
                wsobj.Range["F4"].Value2 = "LC3";
                wsobj.Range["G4"].Value2 = "LF3";
                wsobj.Range["H4"].Value2 = "LC4";
                wsobj.Range["I4"].Value2 = "LF4";
                //wsobj.Range["J4"].Value = "DT1";
                //wsobj.Range["K4"].Value = "DT2";
                //wsobj.Range["L4"].Value = "DT3";
                //wsobj.Range["M4"].Value = "DT4";
                //wsobj.Range["N4"].Value = "DT5";
                wsobj.Range["J4"].Value2 = "INFO";


                /*Excel.Range Selection = wsobj.Range["J3:N3"]; 
                Selection.Merge();
                Selection.HorizontalAlignment = Excel.Constants.xlCenter;*/

                wsobj.Range["A3:J200"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                wsobj.Range["A3:J200"].HorizontalAlignment = Excel.Constants.xlCenter;

                xlApp.DisplayAlerts = false;
                wb.Worksheets["Sheet1"].Delete();
                xlApp.DisplayAlerts = false;
                wb.Close(SaveChanges: true);
                wb = null;

                xlApp.DisplayAlerts = true;
                MessageBox.Show($"LC Template Created in the folder {LcTablePath}. \n Please update the LC Tables for combining more than one load case");
                Marshal.ReleaseComObject(xlApp);
            }
        }


        public Dictionary<string, object> LoadCases(int NCases, ref string FileName, ref string[] HeadersText)
        {
            Dictionary<string, object> LoadCaseDict;
            string[] HeaderKeys;
            string SheetName = LCSheetName;
            //int NThermal;
            //int NCases=4;
            //long  StartColNdx;
            object[] LCData;
            //Excel.Workbook wb = null;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            //Dictionary<string, object> AddnlCaseDict;
            //NCases = 4;
            long RowNdx = 5;
            long ColNdx = 2;
            int NThermal = 5;
            long StartRowNdx = RowNdx;
            long StartColNdx = ColNdx;

            LCData = new object[2 * NCases + 1];
            HeaderKeys = new string[2 * NCases + 1];
            HeadersText = new string[2 * NCases + 1];

            if (File.Exists(FileName))
            {
                Excel.Workbook wb = xlApp.Workbooks.Open(FileName, true, true);
                wb.Activate();
                wb.Windows[1].Visible = false;

                LoadCaseDict = new Dictionary<string, object>();

                for (int j = 0; j <= NCases * 2 - 1; j++)
                {
                    HeaderKeys[j] = wb.Sheets[SheetName].Cells[StartRowNdx - 1, ColNdx + j].Text;
                    HeadersText[j] = wb.Sheets[SheetName].Cells[StartRowNdx - 2, ColNdx + j].Text;
                }


                RowNdx = StartRowNdx;
                int count = 0;
                Dictionary<string, object> AddnlCaseDict;
                while (wb.Sheets[SheetName].Cells[RowNdx, ColNdx - 1].text != "")
                {
                    for (int j = 0; j <= NCases * 2 - 1; j++)
                        LCData[j] = wb.Sheets[SheetName].Cells[RowNdx, ColNdx + j].value;

                    AddnlCaseDict = AddCaseIds(ref HeaderKeys, ref LCData, ref NCases, NThermal);
                    LoadCaseDict.Add(System.Convert.ToString(LCData[0]), AddnlCaseDict);
                    RowNdx = RowNdx + 1;
                    count = count + 1;
                }

                if (count == 0)
                {
                    MessageBox.Show(@"The load case table is empty. Please update the load case table");
                    //LoadCaseDict = null;
                    wb.Close(SaveChanges: false);
                    //wb = null;
                    xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    xlApp.ScreenUpdating = true;
                    return null;
                }
                wb.Close(SaveChanges: false);
                //wb = null;

            }
            else
            {
                MessageBox.Show(@"Load Case file does not exist. Please check and rerun again.");
                //wsobj = null;
                //AddnlCaseDict = null;
                //Marshal.ReleaseComObject(xlApp);
                return null;
            }


            //wsobj = null;
            //AddnlCaseDict = null;
            //Marshal.ReleaseComObject(xlApp);
            return LoadCaseDict;
        }


        private Dictionary<string, object> AddCaseIds(ref string[] HeaderKeys, ref object[] LCData, ref int NCases, int NThermal)
        { //int i;
            Dictionary<string, object> AddCaseId = new Dictionary<string, object>();
            for (int i = 0; i <= NCases * 2 + (NThermal - 1); i++)
            {
                AddCaseId.Add(HeaderKeys[i], LCData[i]);
            }

            return AddCaseId;
        }
    }
}
