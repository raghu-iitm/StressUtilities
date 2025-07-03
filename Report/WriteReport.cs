using StressUtilities;
using StressUtilities.Forms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

/*
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

namespace Report
{
    public enum SheetNameReport
    {
        SHEET_NAME_REPORT = 1,
        SHEET_NAME_REFERENCE = 2
    }

    public enum CellNameReport
    {
        NAME_CELL_REPORT_NAME = 1,
        NAME_CELL_TABLE_ID = 2,
        NAME_CELL_REPORT_FONT = 3,
        NAME_CELL_LIST_ABBREVIATION = 4,
        NAME_CELL_START_TABLE = 5,
        NAME_CELL_START_FIGURE = 6,
        NAME_CELL_START_CONTENT = 7,
        NAME_CELL_PATH_NAME = 8
    }

    class WriteReport
    {

        private bool WriteStatus = true;
        private bool _RefSheetStatus;
        private string bullet { get; set; }
        private char[] listSeparator
        {
            get { return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator.ToCharArray(); }
        }

        private bool RefSheetStatus
        {
            get { return _RefSheetStatus; }
            set { _RefSheetStatus = value; }
        }

        public WriteReport()
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            RefSheetStatus = WorksheetExists("References", wb);
            //RefSheetStatus = false;
            /* for (int i = 0; i < wb.Worksheets.Count; i++)
             {
                 if (wb.Worksheets[i].Name == "References")
                 {
                     RefSheetStatus = true;
                     break;
                 /
             }*/
        }

        public static string GetSheetNameReport(SheetNameReport eNumValue)
        {
            string[] SheetNames = new[] { "ReportOptions", "References" };
            return SheetNames[(int)eNumValue - 1];
        }

        public static string GetCellNameReport(CellNameReport eNumValue)
        {
            string[] CellNames = new[]
            {
                "ReportName", "IDTable", "DefaultFont", "ListAbbr", "ListTables", "ListImages", "ReportContents",
                "ReportPath"
            };
            return CellNames[(int)eNumValue - 1];
        }

        public void LaunchReportContentsForm()
        {
            IEnumerable<ReportContents> FrmCollection =
                System.Windows.Forms.Application.OpenForms.OfType<ReportContents>();

            CheckReportSheet();
            if (FrmCollection.Any())
                FrmCollection.First().Focus();
            else
            {
                ReportContents myUsrform = new ReportContents();
                myUsrform.Show();
            }
        }

        //public void LaunchReportForm()
        //{
        //    Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    try
        //    {
        //        DateTime LastSavedTime = (DateTime)wb.BuiltinDocumentProperties.item["Last Save Time"].value;
        //    }
        //    catch (Exception)
        //    {
        //        MessageBox.Show(@"The workbook is not saved. Please save the workbook before proceeding.");
        //        return;
        //    }

        //    IEnumerable<ReportForm2> FrmCollection = System.Windows.Forms.Application.OpenForms.OfType<ReportForm2>();

        //    CheckReportSheet();
        //    if (FrmCollection.Any())
        //        FrmCollection.First().Focus();
        //    else
        //    {
        //        ReportForm2 myUsrform = new ReportForm2();
        //        myUsrform.Show();
        //    }
        //}

        public void GenerateReport()
        {
            string ReportTitle;
            //string listSeparator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            Excel.Worksheet ReportSheet;

            string ReportSheetName;
            string TableList = "";
            bool ParamCheck;
            bool InputCheck = true;

            string wrdPath;
            string wrdFileName;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Word.Application wrdApp;
            Word.Document wrdDoc;
            Word.Range wrdRng;
            string Defaultfont = "Arial";
            string UserFont;
            DialogResult Response;
            Stopwatch sw = new Stopwatch();
            sw.Start();

            ReportSheetName = GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT);
            if (WorksheetExists(ReportSheetName, wb))
                ReportSheet = wb.Worksheets[GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT)];
            else
            {
                MessageBox.Show(@"Report Input Sheet does not exists.");
                return;
            }

            ReportSheet = wb.Worksheets[ReportSheetName];
            ReportTitle = ReportSheet.Range[GetCellNameReport(CellNameReport.NAME_CELL_REPORT_NAME)].Value;
            wrdPath = ReportSheet.Range[GetCellNameReport(CellNameReport.NAME_CELL_PATH_NAME)].Value;

            wrdFileName = ReportTitle + ".docx";
            wrdFileName = wrdPath + Path.DirectorySeparatorChar + wrdFileName;

            FileInfo fileInf = new FileInfo(wrdFileName);

            if (IsFileOpen(fileInf))
            {
                MessageBox.Show(@"The Word document is alredy open. Please close the document and rerun the program");
                return;
            }

            if (File.Exists(wrdFileName))
            {
                Response = MessageBox.Show(
                    $"The document \n{wrdFileName} exists in the drive.\nAre you sure to append the document?",
                    "Warning!", MessageBoxButtons.YesNo);
                if (Response == DialogResult.No)
                    return;
            }

            xlApp.StatusBar = "Validating Inputs...";
            ParamCheck = CheckSummaryTable(ref TableList);
            InputCheck = ValidateInputs();
            if (ParamCheck == false || InputCheck == false)
                return;

            xlApp.StatusBar = "Preparing the Data...";

            UserFont = Defaultfont;
            wrdApp = new Word.Application();

            if (!File.Exists(wrdFileName))
            {
                wrdDoc = wrdApp.Documents.Add();
                wrdApp.Visible = false;
                {

                    wrdDoc.SaveAs(FileName: wrdFileName);
                }
            }
            else
            {
                wrdDoc = wrdApp.Documents.Open(wrdFileName);
                wrdApp.Visible = false;
                wrdDoc.Activate();
                wrdRng = wrdDoc.Range();
                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();

                wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
                wrdRng.MoveEnd();
                wrdRng.InsertParagraphAfter();
                wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);

            }

            xlApp.ScreenUpdating = false;
            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

            wrdApp.Application.ScreenUpdating = false;
            wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleNormal].Font.Name = Defaultfont;
            try
            {
                wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleNormal].Font.Name = UserFont;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Incorrect Font. The Program will continue with default Arial Font\n{ex.ToString()}");
            }

            bool userOvertype = wrdApp.Options.Overtype;

            // Make sure Overtype is turned off.
            if (wrdApp.Options.Overtype)
                wrdApp.Options.Overtype = false;

            WordReportGeneral(ref wb, ref wrdApp, ref wrdDoc, UserFont);

            // Restore the user's Overtype selection
            wrdApp.Options.Overtype = userOvertype;

            if (WriteStatus)
            {
                wrdDoc.Close(SaveChanges: Word.WdSaveOptions.wdSaveChanges);
                xlApp.StatusBar = "Word Document Saved.";
                MessageBox.Show($"Report Generated.  Elapsed Time: {sw.Elapsed}", @"Stress Utilities");
            }
            else
            {
                wrdDoc.Close(SaveChanges: Word.WdSaveOptions.wdSaveChanges);
                xlApp.StatusBar = $"Word Document Saved.";
                WriteStatus = true;
                MessageBox.Show(@"Error Reported while writing the report. Document Saved Partially",
                    @"Stress Utilities");
            }

            sw.Reset();
            wrdRng = null;
            wrdDoc = null;
            wrdApp.Application.ScreenUpdating = true;
            wrdApp.Application.Quit();
            wrdApp = null;

            xlApp.ScreenUpdating = true;
            xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            xlApp.StatusBar = false;
        }

        private void WordReportGeneral(ref Excel.Workbook wb, ref Word.Application wrdApp, ref Word.Document wrdDoc, string UserFont)
        {
            string[] FEFileList, TblParamReq, Request;
            string[] TblBreakParam = { };
            int TblCount, StartIndex = 0;
            string wrdTableTitle, TableRef="", prevTableRef, wrkShtName, Parameters, TableParam, CriticalItem;
            string Item = "", SetRef, ObjType, ObjList, ObjCaptionPrefix, ParameterID;
            string HeadingText, ParagraphText = "";
            int ListLevel;
            long RowNdx, ColNdx, StartRow, StartCol;
            bool WriteLoA = true, OptionDescription=false;
            int count, ObjectID;
            Excel.Worksheet CalcSheet;

            Dictionary<string, dynamic> DataDict = new Dictionary<string, dynamic>();

            Dictionary<string, dynamic> FormulaTables, SolSymb;

            SortedDictionary<string, Dictionary<string, string>> LoA = new SortedDictionary<string, Dictionary<string, string>>();
            Dictionary<string, dynamic> DictTables= new Dictionary<string, dynamic>();
            //Dictionary<string, dynamic> SolSymb;

            Dictionary<string, string> DictAutoCorrect = new Dictionary<string, string>();
            Word.OMathAutoCorrectEntries aEntries = wrdApp.OMathAutoCorrect.Entries;

            Reference RefData = new Reference();

            Excel.Worksheet ReportSheet = wb.Worksheets[GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT)];
            //string listSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ListSeparator;
            Word.Range currentRange;

            foreach (Word.OMathAutoCorrectEntry aCorrect in aEntries)
                DictAutoCorrect.Add(aCorrect.Name, aCorrect.Value);

            bullet = DictAutoCorrect[@"\bullet"];

            /*count = aEntries.Count;
            for (int k = 1; k <= count; k++)
                DictAutoCorrect.Add(aEntries[k].Name.ToString(),aEntries[k].Value.ToString());*/

            try
            {
                DataDict = InputDataTable(ReportSheet);

                Word.Range myRange = wrdApp.ActiveDocument.Content;
                myRange.Find.Execute(FindText: "LIST OF REFERENCES", Forward: false);

                if (RefSheetStatus && myRange.Find.Found == false)
                {
                    Excel.Worksheet RefSheet = wb.Worksheets[GetSheetNameReport(SheetNameReport.SHEET_NAME_REFERENCE)];
                    Dictionary<string, dynamic> RefDict;
                    RowNdx = RefSheet.Range["REFTABLE"].Row;
                    ColNdx = RefSheet.Range["REFTABLE"].Column;
                    long RefStartRow = RowNdx;
                    string Refkey, ParamKey;
                    StartIndex++;
                    Dictionary<string, dynamic> ParamDict;
                    int CountColumn = RefSheet.Range[RefSheet.Range["REFTABLE"],
                        RefSheet.Range["REFTABLE"].End[Excel.XlDirection.xlToRight]].Count;

                    RefDict = new Dictionary<string, dynamic>();

                    while (!string.IsNullOrEmpty(RefSheet.Cells[RowNdx, ColNdx].text))
                    {
                        ParamDict = new Dictionary<string, dynamic>();
                        Refkey = RefSheet.Cells[RowNdx, ColNdx].Text;

                        for (int j = 0; j < CountColumn; j++)
                        {
                            ParamKey = RefSheet.Cells[RefStartRow, ColNdx + j].Text;
                            ParamDict.Add(ParamKey, RefSheet.Cells[RowNdx, ColNdx + j].text);
                        }

                        RefDict.Add(Refkey, ParamDict);
                        RowNdx++;
                    }

                    RefData.wrdRefTable(wrdApp, wrdDoc, RefDict);
                }

                RowNdx = ReportSheet.Range[GetCellNameReport(CellNameReport.NAME_CELL_START_CONTENT)].Row + 1;
                ColNdx = ReportSheet.Range[GetCellNameReport(CellNameReport.NAME_CELL_START_CONTENT)].Column;
                StartIndex++;
                if (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx].Text) && wrdDoc.TablesOfContents.Count == 0)
                {
                    TblofContents(wrdApp, wrdDoc);
                }

                /*myRange = wrdApp.ActiveDocument.Content;
                myRange.Find.Execute(FindText: "List of Abbreviations", Forward: false);

                if (RefSheetStatus && myRange.Find.Found == false)
                {
                    WriteLoA = true;
                    wrdApp.ActiveDocument.Characters.Last.Select();
                    wrdApp.Selection.Collapse();
                    wrdApp.Selection.TypeParagraph();
                    HeadingListLevel(wrdApp, 2);
                    wrdApp.Selection.ParagraphFormat.set_Style(wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading2].NameLocal);
                    wrdApp.Selection.TypeText(@"List of Abbreviations");
                    wrdApp.Selection.TypeParagraph();
                    wrdApp.Selection.TypeText(@"The list of abbreviations are presented below.");

                    wrdApp.Selection.TypeParagraph();
                }*/


                while (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx].Text))
                {
                    ListLevel = (int)ReportSheet.Cells[RowNdx, ColNdx].Value;
                    HeadingText = ReportSheet.Cells[RowNdx, ColNdx + 1].Text;
                    ParagraphText = ReportSheet.Cells[RowNdx, ColNdx + 2].Text;

                    WriteParagraph(ref wrdApp, ListLevel, HeadingText, ParagraphText);
                    if (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx + 3].Text))
                    {
                        SetRef = ReportSheet.Cells[RowNdx, ColNdx + 3].Text;
                        ObjType = DataDict["FIGURES"][SetRef]["TYPE"];
                        ObjList = DataDict["FIGURES"][SetRef]["FILELIST"];
                        ObjCaptionPrefix = DataDict["FIGURES"][SetRef]["CAPTION"];

                        currentRange = wrdApp.ActiveDocument.Range(wrdApp.ActiveDocument.Content.End - 1,
                            wrdApp.ActiveDocument.Content.End - 1);

                        if (ObjType == "xlChart")
                            FEFileList = ConvertChart2PNG(SetRef, ObjList);
                        else
                            FEFileList = ObjList.Split(listSeparator);

                        if (FEFileList.Length > 0)
                            WrdInsertFigures(ref FEFileList, wrdApp, wrdDoc, ObjCaptionPrefix);

                        ObjectID = 1; // wrdDoc.Tables.Count
                        InsertCrossRefTF(ref wrdApp, currentRange, ref ObjectID, ref ObjCaptionPrefix, "Figure");
                        wrdApp.ActiveDocument.Range(wrdApp.ActiveDocument.Content.End - 1,
                            wrdApp.ActiveDocument.Content.End - 1).Select();

                        wrdApp.ActiveDocument.Characters.Last.Select();
                        wrdApp.Selection.Collapse();
                        wrdApp.Selection.TypeParagraph();
                    }

                    if (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx + 5].Text))
                    {
                        TableRef = ReportSheet.Cells[RowNdx, ColNdx + 5].Text;
                        wrkShtName = wb.Names.Item(TableRef).RefersToRange.Parent.Name;
                        CalcSheet = wb.Worksheets[wrkShtName];
                        CalcSheet.Select();
                        StartRow = CalcSheet.Range[TableRef].Row;

                        Parameters = DataDict["TABLES"][TableRef]["PARAMCALC"];
                        TableParam = DataDict["TABLES"][TableRef]["PARAMTABLE"];
                        CriticalItem = DataDict["TABLES"][TableRef]["ITEMS"];
                        wrdTableTitle = DataDict["TABLES"][TableRef]["CAPTION"];

                        Request = Parameters.Split(listSeparator);

                        StartRow = CalcSheet.Range[TableRef].Row;
                        StartCol = CalcSheet.Range[TableRef].Column;
                        DictTables = CalcDict(TableRef, CalcSheet, "Value");

                        if (TableRef.StartsWith("TableC"))
                        {
                            Dictionary<string, string> ParamDict;
                            foreach (string ParamKey in DictTables["PARAMETER"].Keys)
                            {
                                ParamDict = new Dictionary<string, string>();
                                if (!LoA.ContainsKey(DictTables["SYMBOL"][ParamKey]) && ParamKey!= "PARAMETER")
                                {
                                    ParamDict.Add("SYMBOL", DictTables["SYMBOL"][ParamKey]);
                                    ParamDict.Add("UNIT", DictTables["UNIT"][ParamKey]); 
                                    if (DictTables.ContainsKey("DESCRIPTION"))
                                    {
                                        ParamDict.Add("DESCRIPTION", DictTables["DESCRIPTION"][ParamKey]);
                                    }
                                    else if (DictTables.ContainsKey("TBLDESCR"))
                                    {
                                        ParamDict.Add("DESCRIPTION", DictTables["TBLDESCR"][ParamKey]);
                                    }

                                    if (DictTables["SYMBOL"][ParamKey]!="-")
                                        LoA.Add(DictTables["SYMBOL"][ParamKey], ParamDict);
                                }
                            }
                        }


                        if (!double.TryParse(CriticalItem, out double _))
                        {
                            if (CalcSheet.Range[TableRef].Text == "DESCRIPTION")
                                Item = CriticalItem;
                            else if (CalcSheet.Range[TableRef].Text == "TBLDESCR")
                            {
                                count = 0;
                                while (!string.IsNullOrEmpty(CalcSheet.Cells[StartRow + count, StartCol].Text))
                                {
                                    ParameterID = CalcSheet.Cells[StartRow + count, StartCol].Text;
                                    // Item = ReportSheet.Cells[StartRow + count, ColNdx].Value
                                    if (ParameterID != "TBLDESCR" && ParameterID != "PARAMETER" &&
                                        ParameterID != "SYMBOL" && ParameterID != "UNIT" && ParameterID != "REFERENCE")
                                    {
                                        if (CriticalItem == CalcSheet.Cells[StartRow + count, StartCol + 1].text)
                                            Item = CalcSheet.Cells[StartRow + count, StartCol].text;
                                    }

                                    count++;
                                }
                            }
                            else
                                Item = "DUMMYDATATOESCAPE";
                        }
                        else
                            Item = CriticalItem;


                        if (DictTables.ContainsKey(Item))
                        {
                            Globals.ThisAddIn.Application.StatusBar =
                                $"Preparing the Calculation Steps for the {TableRef} ...";
                            FormulaTables = CalcDict(TableRef, CalcSheet, "Formula");
                            SolSymb = ColMapDict(TableRef, ref CalcSheet);
                            WriteCalculationSteps(wrdApp, wrdDoc, ref DictTables, ref FormulaTables, ref SolSymb,
                                ref Request, ref Item, ref StartRow, ref DictAutoCorrect, ref TableRef, ref RefData,
                                StartIndex, UserFont);
                        }
                    }


                    // -----Tables
                    if (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx + 6].Text))
                    {
                        prevTableRef = TableRef;
                        TableRef = ReportSheet.Cells[RowNdx, ColNdx + 6].Text;
                        wrkShtName = wb.Names.Item(TableRef).RefersToRange.Parent.Name;
                        CalcSheet = wb.Worksheets[wrkShtName];
                        CalcSheet.Select();
                        StartRow = CalcSheet.Range[TableRef].Row;
                        // ColNdx = CalcSheet.Range[TableRef].Column

                        Parameters = DataDict["TABLES"][TableRef]["PARAMCALC"];
                        TableParam = DataDict["TABLES"][TableRef]["PARAMTABLE"];
                        CriticalItem = DataDict["TABLES"][TableRef]["ITEMS"];
                        wrdTableTitle = DataDict["TABLES"][TableRef]["CAPTION"];

                        //StartRow = CalcSheet.Range[TableRef].Row;
                        if(prevTableRef!=TableRef)
                            DictTables = CalcDict(TableRef, CalcSheet, "Value");

                        /*if (TableParam.Contains(listSeparator + "|"))
                            TableParam = TableParam.Replace(listSeparator + "|", "|");
                        if (TableParam.Contains("|" + listSeparator))
                            TableParam = TableParam.Replace("|" + listSeparator, "|");*/

                        if (TableParam.Contains("|"))
                        {
                            TblBreakParam = TableParam.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                            TblCount = TblBreakParam.Length;
                        }
                        else
                            TblCount = 1;

                        currentRange = wrdApp.ActiveDocument.Range(wrdApp.ActiveDocument.Content.End - 1,
                            wrdApp.ActiveDocument.Content.End - 1);

                        if (TblCount == 1)
                        {
                            TblParamReq = TableParam.Split(listSeparator, StringSplitOptions.RemoveEmptyEntries);
                            Globals.ThisAddIn.Application.StatusBar = $"Preparing Word Table for the {TableRef} ...";
                            WriteTblToWord_New(wrdApp, wrdDoc, DictTables, wrdTableTitle, TblParamReq, true, false, OptionDescription,
                                DictAutoCorrect, UserFont);
                        }
                        else
                            for (int i = 0; i <= TblCount - 1; i++)
                            {
                                if (i == 1)
                                    wrdTableTitle += " Cont'd";
                                TableParam = TblBreakParam[i];
                                TblParamReq = TableParam.Split(listSeparator, StringSplitOptions.RemoveEmptyEntries);
                                WriteTblToWord_New(wrdApp, wrdDoc, DictTables, wrdTableTitle, TblParamReq, true, false,
                                    OptionDescription, DictAutoCorrect, UserFont);
                            }

                        ObjectID = wrdDoc.Tables.Count - TblCount + 1;
                        InsertCrossRefTF(ref wrdApp, currentRange, ref ObjectID, ref wrdTableTitle, "Table");
                        wrdApp.ActiveDocument.Range(wrdApp.ActiveDocument.Content.End - 1,
                            wrdApp.ActiveDocument.Content.End - 1).Select();
                    }

                    // Post Paragraph Texts
                    if (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx + 4].Text))
                    {
                        ParagraphText = ReportSheet.Cells[RowNdx, ColNdx + 4].Value;
                        HeadingText = "";
                        WriteParagraph(ref wrdApp, ListLevel, HeadingText, ParagraphText);
                    }
                    RowNdx++;
                }

                if (WriteLoA)
                {
                    WriteListofAbbr(ref LoA, wrdApp, wrdDoc, DictAutoCorrect);
                }

                wrdDoc.Fields.Update();
                wrdDoc.TablesOfContents[1].Update();
                wrdDoc.TablesOfContents[1].UpdatePageNumbers();
            }
            catch (Exception ex)
            {
                WriteStatus = false;
                MessageBox.Show(
                    $"Error Reported from the application. Please verify the imputs or contact the administrator for the support. \n{ex.ToString()}");
            }
        }


        private string[] ConvertChart2PNG(string SetID, string ObjList)
        {
            string[] FileList;
            string ShtName;
            string ChartName;
            //string listSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ListSeparator;
            string[] ChartIdentifier;
            Excel.Worksheet ChartSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string ChartPath = wb.Path + Path.DirectorySeparatorChar + "Charts";
            string FileName;
            Excel.Chart Chartpage;
            Excel.ChartObject xlChartObject;
            Excel.ChartObjects xlChartObjects;

            if (ObjList.Contains(listSeparator.ToString()))
                FileList = ObjList.Split(listSeparator);
            else
            {
                FileList = new string[1];
                FileList[0] = ObjList;
            }

            for (int i = 0; i < FileList.Length; i++)
            {
                ChartIdentifier = FileList[i].Split('|');
                ShtName = ChartIdentifier[0];
                ChartName = ChartIdentifier[1];
                ChartSheet = wb.Worksheets[ShtName];

                xlChartObjects = ChartSheet.ChartObjects();
                try
                {
                    xlChartObject = (Excel.ChartObject)ChartSheet.ChartObjects(1);
                    xlChartObject.Activate();
                }
                catch (Exception ex)
                {
                    return FileList;
                }

                Chartpage = xlChartObject.Chart;

                if (!Directory.Exists(ChartPath))
                    Directory.CreateDirectory(ChartPath);

                FileName = ChartPath + Path.DirectorySeparatorChar + SetID + "_" + ShtName + "_" +
                           ChartName.Replace(" ", "_") + ".png";
                Chartpage.Export(Filename: FileName, FilterName: "PNG");
                FileList[i] = FileName;
            }

            return FileList;
        }

        private void WriteParagraph(ref Word.Application wrdApp, int HeaderLevel, string HeadingText,
            string ParagraphText = "")
        {
            string HeadingStyle = "Heading " + HeaderLevel;
            wrdApp.ActiveDocument.Characters.Last.Select();
            wrdApp.Selection.Collapse();
            if (!string.IsNullOrEmpty(HeadingText))
            {
                HeadingListLevel(wrdApp, HeaderLevel);
                wrdApp.Selection.set_Style(wrdApp.ActiveDocument.Styles[HeadingStyle]);
                wrdApp.Selection.TypeText(HeadingText);
                wrdApp.Selection.TypeParagraph();
            }

            if (!string.IsNullOrEmpty(ParagraphText))
            {
                wrdApp.Selection.TypeText(ParagraphText);
                wrdApp.Selection.TypeParagraph();
            }
        }


        private void WriteCalculationSteps(Word.Application wrdApp, Word.Document wrdDoc,
            ref Dictionary<string, dynamic> DictTables, ref Dictionary<string, dynamic> FormulaTables,
            ref Dictionary<string, dynamic> SolSymb, ref string[] Request, ref string CritItem, ref long StartRow,
            ref Dictionary<string, string> DictAutoCorrect, ref string IDTable, ref Reference RefData, int StartIndex, string UserFont)
        {
            string FormulaString;
            string PrintEquation;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            EquationConverter ConvertFormula = new EquationConverter(IDTable);
            MathConverter ConvertMath = new MathConverter(UserFont);
            string Units;
            string Result;
            Word.Range wrdRng;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string SheetReport, bullet= DictAutoCorrect[@"\bullet"];
            //Reference RefData = new Reference();
            // Dim IDTable As String

            // Scope for future development.
            // The function ConvertFormula.DecodeFormula has been called twice which mostly repeat the same stuff with the exception of a small difference. 
            // The code logic to be updated such that the translation carried out for Math equation is done only once. 
            // As it is not computationally expensive, it has been reatained so far.


            SheetReport = GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT);

            // The below function call improves the speed of the code. Please do not remove or replace the call
            ConvertSybolic(DictTables, FormulaTables, DictAutoCorrect);

            for (int i = 0; i < Request.Length; i++) // Check of Parallel.For loop can be implemented here.
            {
                // Threading.Tasks.Parallel.For(0, UBound(Request), Sub(k) End Sub)
                if (FormulaTables[CritItem].ContainsKey(Request[i]) && FormulaTables[CritItem][Request[i]] != "")
                {
                    xlApp.StatusBar =
                        $"Progress: Translating Formula {i + 1} of {Request.Length} in {IDTable}, Percentage: {(i + 1) / (double)Request.Length:0.00%} ";

                    wrdApp.ActiveDocument.Characters.Last.Select();
                    wrdApp.Selection.Collapse();

                    FormulaString = FormulaTables[CritItem][Request[i]];
                    FormulaString = ConvertFormula.CheckUniLink(FormulaString);
                    if (FormulaString.StartsWith("="))
                    {
                        PrintEquation = ConvertFormula.DecodeFormula(ref FormulaString, DictTables, FormulaTables, SolSymb,
                            Request[i], ref CritItem, ref StartRow, "FORMULA", ref SheetReport, ref bullet/*, IDTable*/);
                        if (double.TryParse(PrintEquation, out double _))
                            FormulaString = PrintEquation;
                        else
                            ConvertMath.MathEquation(wrdApp, wrdDoc, PrintEquation, DictAutoCorrect);

                        if (FormulaTables.ContainsKey("REFERENCE"))
                        {
                            if (FormulaTables["REFERENCE"][Request[i]] != "-" ||
                                FormulaTables["REFERENCE"][Request[i]] != "")
                                RefData.InsertCrossRef(wrdApp, wrdDoc, ref FormulaTables, Request[i], StartIndex);
                        }
                    }

                    if (!FormulaString.StartsWith("=") && !double.TryParse(FormulaString, out double _))
                        PrintEquation = Request[i] + ": " + FormulaString;
                    else
                        PrintEquation = ConvertFormula.DecodeFormula(ref FormulaString, DictTables, FormulaTables, SolSymb,
                            Request[i], ref CritItem, ref StartRow, "VALUES", ref SheetReport, ref bullet/*, IDTable*/);

                    if (!PrintEquation.Contains(@"\sum_(i=1)^"))
                        ConvertMath.MathEquation(wrdApp, wrdDoc, PrintEquation, DictAutoCorrect);

                    wrdApp.ActiveDocument.Characters.Last.Select();
                    wrdApp.Selection.Collapse();
                    wrdApp.Selection.TypeText(" ");

                    if (!FormulaString.StartsWith("="))
                    {
                        Units = DictTables["UNIT"][Request[i]];
                        //Units = Units.Replace("[", "").Replace("]", "");
                        if (Units != "-")
                        {
                            /*if (Units.ToUpper() == "MPA")
                            {
                                Units = "MPa";
                            }
                            else*/
                            if (Units == "°")
                                wrdApp.Selection.TypeBackspace();

                            ConvertMath.MathEquation(wrdApp, wrdDoc, Units, DictAutoCorrect, "NO");
                            if (!FormulaString.StartsWith("=") && FormulaTables.ContainsKey("REFERENCE") &&
                                RefSheetStatus)
                            {
                                RefData.InsertCrossRef(wrdApp, wrdDoc, ref FormulaTables, Request[i], StartIndex);
                            }
                        }
                    }

                    if (FormulaString.StartsWith("="))
                    {
                        wrdRng = wrdDoc.Range();
                        wrdApp.ActiveDocument.Characters.Last.Select();
                        wrdApp.Selection.Collapse();

                        wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
                        wrdRng.MoveEnd();
                        wrdRng.InsertParagraphAfter();
                        wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);


                        wrdApp.ActiveDocument.Characters.Last.Select();
                        wrdApp.Selection.Collapse();
                        wrdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        if (DictTables.ContainsKey("DESCRIPTION"))
                            Result = DictTables["DESCRIPTION"][Request[i]];
                        else
                            Result = DictTables["TBLDESCR"][Request[i]];

                        if (PrintEquation.Contains(@"\matrix") &&
                            ConvertFormula.getDimension((object)xlApp.Application.Evaluate(FormulaString)) != 0)
                        {
                            wrdApp.Selection.TypeText($"\t{Result} has been calculated as \t\t ");
                            PrintEquation = ConvertFormula.MatrixResults("", FormulaString,
                                (DictTables["SYMBOL"][Request[i]]), "VALUES");
                            ConvertMath.MathEquation(wrdApp, wrdDoc, PrintEquation, DictAutoCorrect, "NO");
                        }
                        else
                            wrdApp.Selection.TypeText(
                                $"\t{Result} has been calculated as \t\t{double.Parse(DictTables[CritItem][Request[i]]):0.00} ");

                        Units = DictTables["UNIT"][Request[i]];
                        //Units = Units.Replace("[", "").Replace("]", "");
                        if (Units != "-")
                        {
                            /*if (Units.ToUpper() == "MPA")
                                Units = "MPa";
                            else */
                            if (Units == "°")
                                wrdApp.Selection.TypeBackspace();

                            
                            ConvertMath.MathEquation(wrdApp, wrdDoc, Units, DictAutoCorrect, "NO");
                        }
                    }
                }
            }

            wrdApp.ActiveDocument.Characters.Last.Select();
            wrdApp.Selection.Collapse();
            wrdApp.Selection.TypeParagraph();
        }


        private void ConvertSybolic(Dictionary<string, dynamic> DictTables, Dictionary<string, dynamic> FormulaTables,
            Dictionary<string, string> DictAutoCorrect)
        {
            // Dim aCorrect As Word.OMathAutoCorrectEntry
            //string ParamKey;
            string[] KeyList = new string[101];
            string[] strFormula;
            int Count = 0;
            string[] strUnit;
            int i;
            string strSymb = null;
            string strUnits = null;

            foreach (string ParamKey in DictTables["PARAMETER"].Keys)
            {
                if (ParamKey != "PARAMETER")
                {
                    KeyList[Count] = ParamKey;
                    Count++;
                    strSymb = strSymb + ";" + DictTables["SYMBOL"][ParamKey];
                    strUnits = strUnits + ";" + DictTables["UNIT"][ParamKey];
                }
            }

            if (strSymb.Contains(@"\") || strUnits.Contains(@"\"))
            {
                // For Each aCorrect In aEntries  'wrdApp.OMathAutoCorrect.Entries
                // If strSymb.Contains(aCorrect.Name) Then
                // strSymb = strSymb.Replace(aCorrect.Name, aCorrect.Value)
                // End If

                // If strUnits.Contains(aCorrect.Name) Then
                // strUnits = strUnits.Replace(aCorrect.Name, aCorrect.Value)
                // End If

                // If Not strSymb.Contains("\") And Not strUnits.Contains("\") Then Exit For
                // Next aCorrect
                foreach (string keyItem in DictAutoCorrect.Keys)
                {
                    if (strSymb.Contains(keyItem))
                        strSymb = strSymb.Replace(keyItem, DictAutoCorrect[keyItem]);

                    if (strUnits.Contains(keyItem))
                        strUnits = strUnits.Replace(keyItem, DictAutoCorrect[keyItem]);

                    if (!strSymb.Contains(@"\") && !strUnits.Contains(@"\"))
                        break;
                }
            }

            strFormula = strSymb.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            strUnit = strUnits.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            for (i = 0; i <= Count - 1; i++)
            {
                FormulaTables["SYMBOL"][KeyList[i]] = strFormula[i];
                DictTables["SYMBOL"][KeyList[i]] = strFormula[i];
                FormulaTables["UNIT"][KeyList[i]] = strUnit[i];
                DictTables["UNIT"][KeyList[i]] = strUnit[i];
            }
        }


        private void HeadingListLevel(Word.Application wrdApp, int HeadingLvl)
        {
            string wrdHeadingNr;
            int i;
            Word.ListTemplate ListTemp;


            wrdHeadingNr = "%" + 1;
            ListTemp = wrdApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[5];

            for (i = 1; i <= HeadingLvl; i++)
            {
                if (i > 1)
                    wrdHeadingNr = wrdHeadingNr + "." + "%" + 1;
            }


            switch (HeadingLvl)
            {
                case 1:
                    {
                        ListTemp.ListLevels[HeadingLvl].LinkedStyle =
                            wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading1].NameLocal;
                        break;
                    }

                case 2:
                    {
                        ListTemp.ListLevels[HeadingLvl].LinkedStyle =
                            wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading2].NameLocal;
                        break;
                    }

                case 3:
                    {
                        ListTemp.ListLevels[HeadingLvl].LinkedStyle =
                            wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading3].NameLocal;
                        break;
                    }

                case 4:
                    {
                        ListTemp.ListLevels[HeadingLvl].LinkedStyle =
                            wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading4].NameLocal;
                        break;
                    }

                case 5:
                    {
                        ListTemp.ListLevels[HeadingLvl].LinkedStyle =
                            wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading5].NameLocal;
                        break;
                    }

                case 6:
                    {
                        ListTemp.ListLevels[HeadingLvl].LinkedStyle =
                            wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading6].NameLocal;
                        break;
                    }

                case 7:
                    {
                        ListTemp.ListLevels[HeadingLvl].LinkedStyle =
                            wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading7].NameLocal;
                        break;
                    }
            }

            ListTemp.ListLevels[HeadingLvl].NumberFormat = wrdHeadingNr;
            ListTemp.ListLevels[HeadingLvl].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;


            wrdApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate: ListTemp);

            //ListTemp = null;
        }

        private void WrdInsertFigures(ref string[] FileList, Word.Application wrdApp, Word.Document wrdDoc,
            string CaptionPrefix)
        {
            string filepath;
            string filename;
            string wrdFigureTitle;
            int PicIndex;
            Word.Range wrdRng;
            wrdRng = wrdDoc.Range();
            wrdApp.ActiveDocument.Characters.Last.Select();
            wrdApp.Selection.Collapse();

            wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
            wrdRng.MoveEnd();
            wrdRng.InsertParagraphAfter();
            wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);


            wrdApp.ActiveDocument.Characters.Last.Select();
            wrdApp.Selection.Collapse();
            try
            {
                for (int i = 0; i <= FileList.Length - 1; i++)
                {
                    filepath = FileList[i];
                    filename = Path.GetFileNameWithoutExtension(filepath);

                    wrdFigureTitle = GetCaptionTitle(ref filename, ref CaptionPrefix);
                    wrdDoc.InlineShapes.AddPicture(FileName: filepath, LinkToFile: false, SaveWithDocument: true);
                    PicIndex = wrdDoc.InlineShapes.Count;
                    wrdDoc.InlineShapes[PicIndex].Select();


                    wrdApp.Selection.InsertCaption(Label: "Figure", TitleAutoText: "", Title: ": " + wrdFigureTitle,
                        Position: Word.WdCaptionPosition.wdCaptionPositionBelow, ExcludeLabel: false);
                    wrdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    wrdApp.Selection.TypeParagraph();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not insert the pictures \n{ex.ToString()}");
            }

            wrdRng = null;
        }


        private string GetCaptionTitle(ref string filename, ref string CaptionPrefix)
        {
            string ImageCaption = CaptionPrefix;
            if (!string.IsNullOrEmpty(CaptionPrefix))
                ImageCaption += " " + filename.Replace("_", " ");
            else
                ImageCaption = filename.Replace("_", " ");
            return ImageCaption;
        }

        private void WriteTblToWord(Word.Application wrdApp, Word.Document wrdDoc,
            Dictionary<string, dynamic> DictTables, string wrdTableTitle, string[] TblParamReq, bool Symbolic,
            bool DescriptionOption, bool ItemColOption, Dictionary<string, string> DictAutoCorrect, string UserFont,
            string FilterResults = "")
        {
            int RowNdx, ColNdx;
            //string ItemKey;
            int RowSize, ColumnSize;
            dynamic CellValue; // Why is this taken as dynamic?
            int i;
            Word.Table wrdTbl;
            Word.Range wrdRng;
            string Symbol, Unit;
            Word.Range objRange;
            string DecSep = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            int TableId;
            bool LandScapeLayout = false;
            //int TableWidth;
            MathConverter EqData = new MathConverter(UserFont);
            // TableWidth = EstimateTableWidth(TblParamReq)

            if (TblParamReq.Length > 12 && !DictTables.ContainsKey("DESCRIPTION"))
            {
                LandScapeLayout = true;
                // usrResponse = MessageBox.Show("The number of columns requested for the table does not fit in one page. Do you want to the program include the table in a " &
                // "Landscape page?. This does not guarantee the contents of the table in one page.", "", MessageBoxButtons.YesNo)
                // If usrResponse = DialogResult.Yes Then
                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();
                wrdApp.Selection.TypeParagraph();
                wrdApp.Selection.InsertBreak(Type: Word.WdBreakType.wdSectionBreakNextPage);
                // wrdApp.Selection.TypeParagraph
                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();
            }

            wrdRng = wrdDoc.Range();
            wrdApp.ActiveDocument.Characters.Last.Select();
            wrdApp.Selection.Collapse();

            // If Asc(wrdDoc.Characters(wrdDoc.Selection.Start-1)) = 13 And Asc(wrdDoc.Characters(Selection.Start )) = 13 Then
            // messagebox.show= "y"
            // End If


            wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
            wrdRng.MoveEnd();
            wrdRng.InsertParagraphAfter();
            wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);


            TableId = wrdDoc.Tables.Count;

            RowNdx = 0;
            ColNdx = 1;
            if (DictTables.ContainsKey("TBLDESCR"))
            {
                if (DictTables.ContainsKey("REFERENCE"))
                    RowSize = DictTables.Count - 2;
                else
                    RowSize = DictTables.Count - 1;

                if (DescriptionOption == false)
                    RowSize -= 1;
                ColumnSize = TblParamReq.Length + 1;
            }
            else if (DictTables.ContainsKey("DESCRIPTION"))
            {
                ColumnSize = DictTables.Count - 1;
                RowSize = TblParamReq.Length + 1;
            }
            else
            {
                RowSize = DictTables.Count;
                ColumnSize = TblParamReq.Length;
            }

            wrdDoc.Tables.Add(wrdRng, RowSize, ColumnSize);
            wrdTbl = wrdDoc.Tables[TableId + 1];
            wrdTbl.Borders.Enable = 1; //true

            if (DictTables.ContainsKey("TBLDESCR"))
            {
                wrdTbl.Cell(1, 1).Range.Text = "ITEM";
                for (i = 0; i < TblParamReq.Length; i++)
                {
                    if (DescriptionOption == true)
                    {
                        RowNdx++;
                        wrdTbl.Cell(RowNdx, ColNdx + i + 1).Range.Text = DictTables["TBLDESCR"][TblParamReq[i]];
                        wrdTbl.Cell(RowNdx, ColNdx + i + 1).Range.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }

                    Symbol = DictTables["SYMBOL"][TblParamReq[i]];

                    Unit = DictTables["UNIT"][TblParamReq[i]];
                    //Unit = Unit.Replace("[", "").Replace("]", "");

                    if (Symbolic == true)
                    {
                        objRange = wrdTbl.Cell(RowNdx + 1, ColNdx + i + 1).Range;
                        EqData.TableEquation(wrdApp, objRange, Symbol, DictAutoCorrect);
                        objRange = wrdTbl.Cell(RowNdx + 2, ColNdx + i + 1).Range;
                        EqData.TableEquation(wrdApp, objRange, Unit, DictAutoCorrect);
                    }
                    else
                    {
                        wrdTbl.Cell(RowNdx + 1, ColNdx + i + 1).Range.Text = Symbol;
                        wrdTbl.Cell(RowNdx + 2, ColNdx + i + 1).Range.Text = Unit;
                    }


                    wrdTbl.Cell(RowNdx + 1, ColNdx + i + 1).Range.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wrdTbl.Cell(RowNdx + 2, ColNdx + i + 1).Range.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                RowNdx += 3;

                foreach (string ItemKey in DictTables.Keys)
                {
                    if (double.TryParse(ItemKey, out double _))
                    {
                        if (RowNdx > 3 - Convert.ToInt32(!DescriptionOption)) // &IsNumeric(ItemKey)
                            wrdTbl.Cell(RowNdx, ColNdx).Range.Text = ItemKey;
                        for (i = 0; i < TblParamReq.Length; i++)
                        {
                            CellValue = DictTables[ItemKey][TblParamReq[i]];
                            if (double.TryParse(CellValue, out double ResultValue))
                            {
                                if (CellValue.ToString().IndexOf(DecSep) != -1)
                                    wrdTbl.Cell(RowNdx, ColNdx + i + 1).Range.Text =
                                        $"{Math.Round(ResultValue, 2):0.00}";
                                else
                                    wrdTbl.Cell(RowNdx, ColNdx + i + 1).Range.Text = $"{CellValue:0}";
                                wrdTbl.Cell(RowNdx, ColNdx + i + 1).Range.ParagraphFormat.Alignment =
                                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                            else
                            {
                                wrdTbl.Cell(RowNdx, ColNdx + i + 1).Range.Text = CellValue;
                                wrdTbl.Cell(RowNdx, ColNdx + i + 1).Range.ParagraphFormat.Alignment =
                                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                        }

                        RowNdx++;
                    }
                }
            }
            else if (DictTables.ContainsKey("DESCRIPTION"))
            {
                RowNdx++;
                foreach (string ItemKey in DictTables.Keys)
                {
                    if (ItemKey != "PARAMETER")
                    {
                        if (ItemKey == "SYMBOL")
                            wrdTbl.Cell(RowNdx, ColNdx).Range.Text = "PARAMETER";
                        else
                            wrdTbl.Cell(RowNdx, ColNdx).Range.Text = ItemKey;
                        wrdTbl.Cell(RowNdx, ColNdx).Range.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        for (i = 0; i < TblParamReq.Length; i++)
                        {
                            CellValue = DictTables[ItemKey][TblParamReq[i]];
                            if (double.TryParse(CellValue, out double ResultValue))
                            {
                                if (CellValue.ToString().IndexOf(DecSep) != -1)
                                    wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.Text =
                                        $"{Math.Round(ResultValue, 2):0.00}";
                                else
                                    wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.Text = $"{CellValue:0}";
                                wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.ParagraphFormat.Alignment =
                                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                            else
                            {
                                if (ItemKey == "SYMBOL")
                                {
                                    Symbol = DictTables["SYMBOL"][TblParamReq[i]];
                                    objRange = wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range;
                                    EqData.TableEquation(wrdApp, objRange, Symbol, DictAutoCorrect);
                                }
                                else if (ItemKey == "UNIT")
                                {
                                    Unit = DictTables["UNIT"][TblParamReq[i]];
                                    //Unit = Unit.Replace("[", "").Replace("]", "");
                                    objRange = wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range;
                                    EqData.TableEquation(wrdApp, objRange, Unit, DictAutoCorrect);
                                }
                                else
                                {
                                    wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.Text = CellValue;
                                }

                                wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.ParagraphFormat.Alignment =
                                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                        }

                        ColNdx++;
                    }
                }
            }
            else
            {
                RowNdx++;
                foreach (string ItemKey in DictTables.Keys)
                {
                    for (i = 0; i < TblParamReq.Length; i++)
                    {
                        CellValue = DictTables[ItemKey][TblParamReq[i]];
                        if (IsNumeric(CellValue))
                        {
                            if (CellValue.ToString().IndexOf(DecSep) != -1)
                                wrdTbl.Cell(RowNdx, ColNdx + i).Range.Text =
                                    $"{Math.Round(CellValue, 2):0.00}"; // = string.Format(Math.Round(CellValue, 2), "0.00");
                            else
                                wrdTbl.Cell(RowNdx, ColNdx + i).Range.Text =
                                    $"{CellValue:0}"; //string.Format(CellValue, "0");

                            wrdTbl.Cell(RowNdx, ColNdx + i).Range.ParagraphFormat.Alignment =
                                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        else
                        {
                            wrdTbl.Cell(RowNdx, ColNdx + i).Range.Text = CellValue;
                            wrdTbl.Cell(RowNdx, ColNdx + i).Range.ParagraphFormat.Alignment =
                                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                    }

                    RowNdx++;
                }
            }

            WrdTableFormat(ref wrdApp, ref wrdTbl, ref TableId, ref wrdTableTitle);

            if (DictTables.ContainsKey("TBLDESCR"))
            {
                if (ItemColOption == false)
                    wrdTbl.Columns[1].Delete();
                else
                {
                    wrdRng = wrdTbl.Cell(1, 1).Range;
                    wrdRng.End = wrdTbl.Cell(3 - Convert.ToInt32(!DescriptionOption), 1).Range.End;
                    wrdRng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wrdRng.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    wrdTbl.Range.Cells[1].Merge(wrdTbl.Cell(3 - Convert.ToInt32(!DescriptionOption), 1));
                }
            }

            // wrdRng = wrdDoc.Range
            // wrdRng.InsertParagraphAfter()
            wrdApp.ActiveDocument.Characters.Last.Select();
            // wrdApp.Selection.Collapse()
            // wrdApp.Selection.TypeParagraph()

            if (LandScapeLayout)
            {
                wrdApp.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();
                wrdApp.Selection.TypeParagraph();
                wrdApp.Selection.InsertBreak(Type: Word.WdBreakType.wdSectionBreakNextPage);
                wrdApp.Selection.TypeParagraph();
                wrdApp.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            }

            wrdRng = null;
            wrdTbl = null;
        }


        private void WriteTblToWord_New(Word.Application wrdApp, Word.Document wrdDoc,
    Dictionary<string, dynamic> DictTables, string wrdTableTitle, string[] TblParamReq, bool Symbolic,
    bool DescriptionOption, bool ItemColOption, Dictionary<string, string> DictAutoCorrect, string UserFont,
    string FilterResults = "")
        {
            int RowNdx, ColNdx;
            //string ItemKey;
            int RowSize, ColumnSize;
            dynamic CellValue; // Why is this taken as dynamic?
            int i;
            Word.Table wrdTbl;
            Word.Range wrdRng;
            string Symbol, Unit;
            Word.Range objRange;
            string DecSep = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            int TableId;
            bool LandScapeLayout = false;
            //int TableWidth;
            MathConverter EqData = new MathConverter(UserFont);
            // TableWidth = EstimateTableWidth(TblParamReq)

            if (TblParamReq.Length > 12 && !DictTables.ContainsKey("DESCRIPTION"))
            {
                LandScapeLayout = true;
                // usrResponse = MessageBox.Show("The number of columns requested for the table does not fit in one page. Do you want to the program include the table in a " &
                // "Landscape page?. This does not guarantee the contents of the table in one page.", "", MessageBoxButtons.YesNo)
                // If usrResponse = DialogResult.Yes Then
                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();
                wrdApp.Selection.TypeParagraph();
                wrdApp.Selection.InsertBreak(Type: Word.WdBreakType.wdSectionBreakNextPage);
                // wrdApp.Selection.TypeParagraph
                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();
            }

            wrdRng = wrdDoc.Range();
            wrdApp.ActiveDocument.Characters.Last.Select();
            wrdApp.Selection.Collapse();

            wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
            wrdRng.MoveEnd();
            wrdRng.InsertParagraphAfter();
            wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);

            TableId = wrdDoc.Tables.Count;

            RowNdx = 0;
            ColNdx = 1;
            if (DictTables.ContainsKey("TBLDESCR"))
            {
                if (DictTables.ContainsKey("REFERENCE"))
                    RowSize = DictTables.Count - 2;
                else
                    RowSize = DictTables.Count - 1;

                if (DescriptionOption == false)
                    RowSize -= 1;
                ColumnSize = TblParamReq.Length + 1;
            }
            else if (DictTables.ContainsKey("DESCRIPTION"))
            {
                ColumnSize = DictTables.Count - 1;
                RowSize = TblParamReq.Length + 1;
            }
            else
            {
                RowSize = DictTables.Count;
                ColumnSize = TblParamReq.Length;
            }

            object tab = Word.WdTableFieldSeparator.wdSeparateByTabs;
            string DescrText = "", SymbolText = "", UnitText = ""; //TblText = ""
            StringBuilder TblText = new StringBuilder();

            if (DictTables.ContainsKey("TBLDESCR"))
            {                
                TblText.Append("ITEM");           

                for (i = 0; i < TblParamReq.Length; i++)
                {
                    if (DescriptionOption == true)
                    {
                        DescrText += DictTables["TBLDESCR"][TblParamReq[i]]+ "\t" ;
                        
                    }

                    Symbol = DictTables["SYMBOL"][TblParamReq[i]];

                    Unit = DictTables["UNIT"][TblParamReq[i]];
                    //Unit = Unit.Replace("[", "").Replace("]", "");

                    /*
                    wrdTbl.Cell(RowNdx + 1, ColNdx + i + 1).Range.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wrdTbl.Cell(RowNdx + 2, ColNdx + i + 1).Range.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;*/
                    SymbolText += "\t" + Symbol;
                    UnitText += "\t" + Unit;
                }
                if (DescriptionOption == true)
                {
                    DescrText = DescrText.Substring(0, DescrText.LastIndexOf('\t'));
                    DescrText += "\n";
                }

                TblText.Append(DescrText + SymbolText+ "\n" + UnitText+ "\n");
                /*
                RowNdx += 3;*/

                foreach (string ItemKey in DictTables.Keys)
                {
                    if (double.TryParse(ItemKey, out double _))
                    {
                            TblText.Append(ItemKey + "\t");
                        for (i = 0; i < TblParamReq.Length; i++)
                        {
                            CellValue = DictTables[ItemKey][TblParamReq[i]];
                            if (double.TryParse(CellValue, out double ResultValue))
                            {
                                if (CellValue.ToString().IndexOf(DecSep) != -1)
                                    TblText.Append( $"{Math.Round(ResultValue, 2):0.00}" + "\t");
                                else
                                    TblText.Append( $"{CellValue:0}" + "\t");
                            }
                            else
                            {
                                TblText.Append( $"{CellValue}" + "\t");
                            }
                        }
                        //TblText = TblText.Substring(0, TblText.LastIndexOf('\t'));
                        TblText.Remove(TblText.Length - 1, 1);
                        TblText.Append("\n");
                    }
                }
                //TblText.TrimEnd('\n');
                wrdRng.Text = TblText.ToString();
                wrdTbl = wrdRng.ConvertToTable(Separator: ref tab);

                for (i = 0; i < TblParamReq.Length; i++)
                {
                    if (Symbolic == true)
                    {
                        if (DescriptionOption == true)
                            RowNdx++;

                        Symbol = DictTables["SYMBOL"][TblParamReq[i]];

                        Unit = DictTables["UNIT"][TblParamReq[i]];
                        //Unit = Unit.Replace("[", "").Replace("]", "");

                        objRange = wrdTbl.Cell(RowNdx + 1, ColNdx + i + 1).Range;
                        EqData.TableEquation(wrdApp, objRange, Symbol, DictAutoCorrect);

                        if (Unit.Contains(@"\") || Unit.Contains(@"^"))
                        {
                            objRange = wrdTbl.Cell(RowNdx + 2, ColNdx + i + 1).Range;
                            EqData.TableEquation(wrdApp, objRange, Unit, DictAutoCorrect);
                        }
                    }
                }
            }
            else if (DictTables.ContainsKey("DESCRIPTION"))
            {
                RowNdx++;
                wrdDoc.Tables.Add(wrdRng, RowSize, ColumnSize);
                wrdTbl = wrdDoc.Tables[TableId + 1];

                foreach (string ItemKey in DictTables.Keys)
                {
                    if (ItemKey != "PARAMETER")
                    {
                        if (ItemKey == "SYMBOL")
                            wrdTbl.Cell(RowNdx, ColNdx).Range.Text = "PARAMETER";
                            //TblText = "ITEM\t";
                        else
                            wrdTbl.Cell(RowNdx, ColNdx).Range.Text = ItemKey;
                            //TblText += ItemKey + "\t";
                        wrdTbl.Cell(RowNdx, ColNdx).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        for (i = 0; i < TblParamReq.Length; i++)
                        {
                            CellValue = DictTables[ItemKey][TblParamReq[i]];
                            if (double.TryParse(CellValue, out double ResultValue))
                            {
                                if (CellValue.ToString().IndexOf(DecSep) != -1)
                                    wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.Text = $"{Math.Round(ResultValue, 2):0.00}";
                                    //TblText += $"{Math.Round(ResultValue, 2):0.00}" + "\t";
                                else
                                    wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.Text = $"{CellValue:0}";
                                    //TblText += $"{CellValue:0}" + "\t";
                                wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                            else
                            {
                                if (ItemKey == "SYMBOL")
                                {
                                    Symbol = DictTables["SYMBOL"][TblParamReq[i]];
                                    objRange = wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range;
                                    EqData.TableEquation(wrdApp, objRange, Symbol, DictAutoCorrect);
                                    //TblText += $"{Symbol}" + "\t";
                                }
                                else if (ItemKey == "UNIT")
                                {
                                    Unit = DictTables["UNIT"][TblParamReq[i]];
                                    //Unit = Unit.Replace("[", "").Replace("]", "");
                                   objRange = wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range;
                                    EqData.TableEquation(wrdApp, objRange, Unit, DictAutoCorrect);
                                    //TblText += $"{Unit}" + "\t";
                                }
                                else
                                {
                                    wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.Text = CellValue;
                                    //TblText += $"{CellValue}" + "\t";
                                }

                                wrdTbl.Cell(RowNdx + i + 1, ColNdx).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                        }

                        ColNdx++;
                    }
                }
                /*TblText.TrimEnd('\n');
                TblRange.Text = TblText;
                wrdTbl = TblRange.ConvertToTable(Separator: ref tab);*/
            }
            else
            {
                //RowNdx++;
                foreach (string ItemKey in DictTables.Keys)
                {
                    for (i = 0; i < TblParamReq.Length; i++)
                    {
                        CellValue = DictTables[ItemKey][TblParamReq[i]];
                        if (IsNumeric(CellValue))
                        {
                            if (CellValue.ToString().IndexOf(DecSep) != -1)
                                //wrdTbl.Cell(RowNdx, ColNdx + i).Range.Text = $"{Math.Round(CellValue, 2):0.00}"; 
                                TblText.Append($"{Math.Round(CellValue, 2):0.00}" + "\t");
                            else
                                //wrdTbl.Cell(RowNdx, ColNdx + i).Range.Text = $"{CellValue:0}"; 
                                TblText.Append( $"{CellValue:0}" + "\t");

                            //wrdTbl.Cell(RowNdx, ColNdx + i).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        else
                        {
                            TblText.Append( $"{CellValue}" + "\t");
                            //wrdTbl.Cell(RowNdx, ColNdx + i).Range.Text = CellValue;
                            //wrdTbl.Cell(RowNdx, ColNdx + i).Range.ParagraphFormat.Alignment =  Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                    }
                    //TblText = TblText.Substring(0, TblText.LastIndexOf('\t'));
                    TblText.Remove(TblText.Length - 1, 1);
                    TblText.Append( "\n");
                    //RowNdx++;
                }
                //TblText.TrimEnd('\n');
                wrdRng.Text = TblText.ToString();
                wrdTbl = wrdRng.ConvertToTable(Separator: ref tab);
            }
            
            wrdTbl.Borders.Enable = 1;

            WrdTableFormat(ref wrdApp, ref wrdTbl, ref TableId, ref wrdTableTitle);

            if (DictTables.ContainsKey("TBLDESCR"))
            {
                if (ItemColOption == false)
                    wrdTbl.Columns[1].Delete();
                else
                {
                    wrdRng = wrdTbl.Cell(1, 1).Range;
                    wrdRng.End = wrdTbl.Cell(3 - Convert.ToInt32(!DescriptionOption), 1).Range.End;
                    wrdRng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wrdRng.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    wrdTbl.Range.Cells[1].Merge(wrdTbl.Cell(3 - Convert.ToInt32(!DescriptionOption), 1));
                }
            }

            // wrdRng = wrdDoc.Range
            // wrdRng.InsertParagraphAfter()
            wrdApp.ActiveDocument.Characters.Last.Select();
            // wrdApp.Selection.Collapse()
            // wrdApp.Selection.TypeParagraph()

            if (LandScapeLayout)
            {
                wrdApp.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();
                wrdApp.Selection.TypeParagraph();
                wrdApp.Selection.InsertBreak(Type: Word.WdBreakType.wdSectionBreakNextPage);
                wrdApp.Selection.TypeParagraph();
                wrdApp.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            }

            wrdRng = wrdTbl.Range;
            //wrdRng.End = wrdTbl.Cell(3 - Convert.ToInt32(!DescriptionOption), 1).Range.End;
            wrdRng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wrdRng.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            
            //wrdTbl = null;
        }

        private void WrdTableFormat(ref Word.Application wrdApp, ref Word.Table wrdTbl, ref int TableId,
            ref string wrdTableTitle)
        {
            wrdTbl.Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            wrdTbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
            wrdTbl.Columns.AutoFit();
            wrdTbl.Range.Font.Size = 10;
            wrdTbl.Rows.Height = 14;
            wrdTbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;

            wrdApp.ActiveDocument.Tables[TableId + 1].Rows[1].Select();

            wrdApp.Selection.Rows.HeadingFormat = -1; //true
            wrdApp.Selection.Font.Size = 8;
            wrdApp.Selection.Rows.Height = 24;
            wrdApp.Selection.Font.Name = "Arial Black";
            wrdApp.Selection.Font.Italic = 2; //true
            wrdApp.Selection.InsertCaption(Label: "Table", TitleAutoText: "", Title: ": " + wrdTableTitle,
                Position: Word.WdCaptionPosition.wdCaptionPositionAbove, ExcludeLabel: false);
            wrdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        }

        private void InsertCrossRefTF(ref Word.Application wrdApp, Word.Range currentRange, ref int ObjectID,
            ref string ObjectTitle, string ObjectType)
        {
            currentRange.InsertBefore($"The {ObjectTitle} is presented in ");
            currentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            currentRange.Select();
            wrdApp.Selection.InsertCrossReference(ReferenceType: ObjectType,
                ReferenceKind: Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem: ObjectID,
                InsertAsHyperlink: true, IncludePosition: false, SeparateNumbers: false, SeparatorString: " ");

            wrdApp.Selection.InsertAfter(".");

        }

        //private void TableEquation(Word.Application wrdApp, Word.Range objRange, string PrintEquation)
        //{
        //    Word.OMath objEq;
        //    wrdApp.OMathAutoCorrect.UseOutsideOMath = true;

        //    objRange.Text = PrintEquation;
        //    if (PrintEquation.Contains(@"\"))
        //    {
        //        foreach (Word.OMathAutoCorrectEntry aCorrect in wrdApp.OMathAutoCorrect.Entries)
        //        {
        //            {
        //                if (objRange.Text.Contains(aCorrect.Name))
        //                    objRange.Text = objRange.Text.Replace(aCorrect.Name, aCorrect.Value);
        //                if (!objRange.Text.Contains(@"\"))
        //                    break;
        //            }
        //        }
        //    }
        //    objRange = objRange.OMaths.Add(objRange);
        //    objEq = objRange.OMaths[1];
        //    objEq.BuildUp();
        //}


        private bool IsNumeric(dynamic Value)
        {
            return double.TryParse(Value, out double _);
        }

        private Dictionary<string, dynamic> CalcDict(string TableRef, Excel.Worksheet CalcSheet, string RequestType)
        {
            long RowNdx, ColNdx, StartRow, StartCol;
            int CountColumn,i, CountRow;
            string CPkey, ParamKey, Unit;
            Dictionary<string, dynamic> ParamDict;

            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            // Dim strFormula As String
            CalcSheet.Select();
            CalcSheet.Range[TableRef].Select();
            RowNdx = xlApp.Selection.Row;
            ColNdx = xlApp.Selection.Column;
            StartRow = RowNdx;
            StartCol = ColNdx;

            CountColumn = CalcSheet.Range[CalcSheet.Range[TableRef],
                CalcSheet.Range[TableRef].End[Excel.XlDirection.xlToRight]].Count;
            CountRow = CalcSheet
                .Range[CalcSheet.Range[TableRef], CalcSheet.Range[TableRef].End[Excel.XlDirection.xlDown]].Count;

            Dictionary<string, dynamic> ReturnDict = new Dictionary<string, dynamic>();

            switch (CalcSheet.Range[TableRef].Text)
            {
                case "TBLDESCR":
                    {
                        while (!string.IsNullOrEmpty(CalcSheet.Cells[RowNdx, ColNdx].formula))
                        {
                            ParamDict = new Dictionary<string, dynamic>();
                            CPkey = CalcSheet.Cells[RowNdx, ColNdx].Text;

                            for (i = 0; i <= CountColumn - 1; i++)
                            {
                                ParamKey = CalcSheet.Cells[StartRow + 1, ColNdx + i].Text;
                                if (CPkey == "UNIT")
                                {
                                    Unit = CalcSheet.Cells[RowNdx, ColNdx + i].Text;
                                    Unit = Unit.Replace("[", "").Replace("]", "");
                                    if (Unit.ToUpper() == "MPA")
                                    {
                                        Unit = "MPa";
                                    }
                                    ParamDict.Add(ParamKey, Unit);
                                }
                                else if (RequestType == "Value")
                                    ParamDict.Add(ParamKey, CalcSheet.Cells[RowNdx, ColNdx + i].Text);
                                else
                                    ParamDict.Add(ParamKey, CalcSheet.Cells[RowNdx, ColNdx + i].Formula);
                            }

                            ReturnDict.Add(CPkey, ParamDict);
                            RowNdx++;
                        }

                        break;
                    }

                case "DESCRIPTION":
                    {
                        while (!string.IsNullOrEmpty(CalcSheet.Cells[RowNdx, ColNdx].formula))
                        {
                            ParamDict = new Dictionary<string, dynamic>();

                            // For j = 3 To CountColumn - 2
                            // If CalcSheet.Cells[StartRow, ColNdx].Value <> "" Then
                            CPkey = CalcSheet.Cells[StartRow, ColNdx].Text;
                            // End If

                            for (i = 1; i <= CountRow - 1; i++)
                            {
                                ParamKey = CalcSheet.Cells[StartRow + i, StartCol + 1].Text;
                                if (CPkey == "UNIT")
                                {
                                    Unit = CalcSheet.Cells[RowNdx + i, ColNdx].Text;
                                    Unit = Unit.Replace("[", "").Replace("]", "");
                                    if (Unit.ToUpper() == "MPA")
                                    {
                                        Unit = "MPa";
                                    }
                                    ParamDict.Add(ParamKey, Unit);
                                }
                                else if (RequestType == "Value")
                                    ParamDict.Add(ParamKey, CalcSheet.Cells[RowNdx + i, ColNdx].Text);
                                else
                                    ParamDict.Add(ParamKey, CalcSheet.Cells[RowNdx + i, ColNdx].Formula);
                            }

                            ReturnDict.Add(CPkey, ParamDict);
                            ColNdx++;
                        }

                        break;
                    }

                default:
                    {
                        // RowNdx += 1
                        while (!string.IsNullOrEmpty(CalcSheet.Cells[RowNdx, ColNdx].formula))
                        {
                            ParamDict = new Dictionary<string, dynamic>();
                            CPkey = CalcSheet.Cells[RowNdx, ColNdx].Text;
                            for (i = 0; i <= CountColumn - 1; i++)
                            {
                                ParamKey = CalcSheet.Cells[StartRow, ColNdx + i].Text;
                                if (CPkey == "UNIT")
                                {
                                    Unit = CalcSheet.Cells[RowNdx, ColNdx + i].Text;
                                    Unit = Unit.Replace("[", "").Replace("]", "");
                                    if (Unit.ToUpper() == "MPA")
                                    {
                                        Unit = "MPa";
                                    }
                                    ParamDict.Add(ParamKey, Unit);
                                }
                                else if (RequestType == "Value")
                                    ParamDict.Add(ParamKey, CalcSheet.Cells[RowNdx, ColNdx + i].Text);
                                else
                                    ParamDict.Add(ParamKey, CalcSheet.Cells[RowNdx, ColNdx + i].Formula);
                            }

                            ReturnDict.Add(CPkey, ParamDict);
                            RowNdx++;
                        }

                        break;
                    }
            }

            return ReturnDict;
        }


        private Dictionary<string, dynamic> ColMapDict(string TableRef, ref Excel.Worksheet CalcSheet)
        {
            long RowNdx;
            long ColNdx;
            long StartRow;
            int CountColumn;
            string ColKey;
            string ParamKey;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlapp = Globals.ThisAddIn.Application;

            CalcSheet.Select();

            CalcSheet.Range[TableRef].Select();
            RowNdx = xlapp.Selection.Row;
            ColNdx = xlapp.Selection.Column;
            StartRow = RowNdx;

            CountColumn = CalcSheet.Range[CalcSheet.Range[TableRef],
                CalcSheet.Range[TableRef].End[Excel.XlDirection.xlToRight]].Count;

            Dictionary<string, dynamic> ReturnDict = new Dictionary<string, dynamic>();

            while (!string.IsNullOrEmpty(CalcSheet.Cells[RowNdx, ColNdx].formula))
            {
                ColKey = CalcSheet.Cells[RowNdx, ColNdx].Address.Split('$')[1];
                ParamKey = CalcSheet.Cells[StartRow + 1, ColNdx].Text;
                ReturnDict.Add(ColKey, ParamKey);
                ColNdx++;
            }

            return ReturnDict;
        }

        private void CheckReportSheet()
        {
            string RptSheetName;
            Excel.Worksheet wrkSheet;
            bool SheetExistChk;
            string[] CellNames = new[] { "ReportPath", "ReportName", "IDTable", "DefaultFont", "ListAbbr" };
            long RowNdx;
            long ColNdx;
            int i;
            Excel.Worksheet CurrentSheet;
            Excel.Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Range Selection;
            Excel.Range StartRange;

            CurrentSheet = wb.Worksheets[ActiveSheet.Name];

            RptSheetName = GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT);
            SheetExistChk = false;
            foreach (Excel.Worksheet Sheet in wb.Worksheets)
            {
                // if (wb.Application.Proper(Sheet.Name) == wb.Application.Proper(RptSheetName))
                if (Sheet.Name == RptSheetName)
                    SheetExistChk = true;
            }

            if (SheetExistChk == false)
            {
                {
                    wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]).Name = RptSheetName;
                }
                wrkSheet = wb.Worksheets[RptSheetName];

                StartRange = wrkSheet.Range["B2"];

                StartRange.Value = "OPTIONS FOR THE REPORT PREPARATION";
                StartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                StartRange.Offset[1, 0].Value = "Report Path";
                StartRange.Offset[2, 0].Value = "Report File Name";
                StartRange.Offset[3, 0].Value = "Selected Table";
                StartRange.Offset[4, 0].Value = "General Font";
                StartRange.Offset[5, 0].Value = "Write Abbreviations?";
                StartRange.Offset[1, 1].Value = @"C:\Temp";
                StartRange.Offset[2, 1].Value = "FileName";
                StartRange.Offset[3, 1].Value = "Table1";
                StartRange.Offset[4, 1].Value = "Arial";
                StartRange.Offset[5, 1].Value = "Yes";

                Selection = wrkSheet.Range[StartRange.Offset[1, 0], StartRange.Offset[5, 1]];
                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                Selection = wrkSheet.Range[StartRange.Offset[1, 0], StartRange.Offset[5, 0]];
                Selection.Font.Bold = FontStyle.Bold;

                // ---------
                StartRange = wrkSheet.Range["B9"];
                StartRange.Value = "PLEASE DO NOT CHANGE THE FORMAT OF THE TABLE (Insert rows as needed)";
                StartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                StartRange.Offset[1, 0].Value = "TABLENAME";
                StartRange.Offset[1, 0].Name = "ListTables";
                StartRange.Offset[1, 1].Value = "PARAMCALC";
                StartRange.Offset[1, 2].Value = "PARAMTABLE";
                StartRange.Offset[1, 3].Value = "ITEMS";
                StartRange.Offset[1, 4].Value = "CAPTION";

                Selection = wrkSheet.Range[StartRange.Offset[1, 0], StartRange.Offset[20, 4]];
                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                Selection = wrkSheet.Range[StartRange.Offset[1, 0], StartRange.Offset[1, 4]];
                Selection.Font.Bold = FontStyle.Bold;

                StartRange = wrkSheet.Range["J9"];
                StartRange.Value = "PLEASE DO NOT CHANGE THE FORMAT OF THE TABLE (Insert rows as needed)";
                StartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                StartRange.Offset[1, 0].Value = "SETID";
                StartRange.Offset[1, 0].Name = "ListImages";
                StartRange.Offset[1, 1].Value = "TYPE";
                StartRange.Offset[1, 2].Value = "FILELIST";
                StartRange.Offset[1, 3].Value = "CAPTION";

                Selection = wrkSheet.Range[StartRange.Offset[1, 0], StartRange.Offset[20, 3]];
                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                Selection = wrkSheet.Range[StartRange.Offset[1, 0], StartRange.Offset[1, 3]];
                Selection.Font.Bold = FontStyle.Bold;

                StartRange = wrkSheet.Range["B31"];

                StartRange.Value =
                    "PLEASE DO NOT CHANGE THE FORMAT OF THE TABLE (Insert rows as needed, Leave Cells empty when not needed. Header Level is mandatory)";
                StartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                StartRange.Offset[1, 0].Value = "Header Level";
                StartRange.Offset[1, 0].Name = "ReportContents";
                StartRange.Offset[1, 1].Value = "Heading Text";
                StartRange.Offset[1, 2].Value = "Paragraph Text";
                StartRange.Offset[1, 3].Value = "Charts/Figures Set ID";
                StartRange.Offset[1, 4].Value = "Paragraph Text (Post Image/Table)";
                StartRange.Offset[1, 5].Value = "Calculation Source (Table Set ID)";
                StartRange.Offset[1, 6].Value = "Table Source (Table Set ID)";

                Selection = wrkSheet.Range[StartRange.Offset[1, 0], StartRange.Offset[50, 6]];
                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                Selection = wrkSheet.Range[StartRange.Offset[1, 0], StartRange.Offset[1, 6]];
                Selection.Font.Bold = FontStyle.Bold;

                // End With

                wrkSheet.Columns["B"].ColumnWidth = 25;
                wrkSheet.Columns["C:D"].ColumnWidth = 70;
                // wrkSheet.Columns["D"].ColumnWidth = 70
                wrkSheet.Columns["E"].ColumnWidth = 25;
                wrkSheet.Columns["F"].ColumnWidth = 40;
                wrkSheet.Columns["G:H"].ColumnWidth = 25;
                // wrkSheet.Columns["H"].ColumnWidth = 25
                wrkSheet.Columns["J:K"].ColumnWidth = 25;
                wrkSheet.Columns["L:M"].ColumnWidth = 70;
                wrkSheet.Range["C3"].Select();
                RowNdx = wrkSheet.Range["C3"].Row;
                ColNdx = wrkSheet.Range["C3"].Column;

                for (i = 0; i < CellNames.Length; i++)
                    wrkSheet.Cells[RowNdx + i, ColNdx].Name = CellNames[i];
                wrkSheet.Select();
                wb.Windows[1].DisplayGridlines = false;
                CurrentSheet.Select();
            }
        }



        public static string GetTableListData()
        {
            //Name TblName;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            string TableListData = "";

            foreach (Excel.Name TblName in wb.Names)
            {
                if (TblName.Name.ToUpper().StartsWith("TABLEC"))
                    TableListData = TableListData + ";" + TblName.Name;
                else if (TblName.Name.ToUpper().Contains("!TABLEC"))
                    TableListData = TableListData + ";" + TblName.Name.Split('!')[1];
            }

            if (TableListData != "")
                TableListData = TableListData.Substring(1, TableListData.Length - 1);
            else
                TableListData = "No Table Data";
            return TableListData;

        }

        public static void DeleteTableListData()
        {
            //Name TblName;
            string wrkShtName;
            Excel.Worksheet CalcSheet;
            DialogResult Response;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            Response = MessageBox.Show(
                @"You are about to delete the Range names beginning with ""Table"". Are you sure?.", "Confirmation",
                MessageBoxButtons.YesNo);

            if (Response == DialogResult.Yes)
            {
                foreach (Excel.Name TblName in wb.Names)
                {
                    if (xlApp.WorksheetFunction.IsErr(TblName.Name)) //(TblName.Name.Contains(XlCVError.xlErrRef))
                        TblName.Delete();
                    else if (TblName.Name.ToUpper().StartsWith("TABLEC"))
                    {
                        try
                        {
                            wrkShtName = wb.Names.Item(TblName.Name).RefersToRange.Parent.Name;
                            CalcSheet = wb.Worksheets[wrkShtName];
                            CalcSheet.Range[TblName].Name.Delete();
                        }
                        catch (Exception ex)
                        {
                            TblName.Delete();
                        }
                    }
                }

                MessageBox.Show(
                    "The Table identifiers have been deleted. You must create them again in the first cell of the table for the tool to function.");
            }
        }

        public void AutoTableNames()
        {
            //Name TblName;
            string wrkShtName;
            Excel.Worksheet CalcSheet;
            DialogResult Response;

            //Excel.Worksheet wrkSht;
            Excel.Range CellRng;
            int i = 1;
            string StartAddress;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            Response = MessageBox.Show(
                "You are about to create the table identifiers on the tables with the first cell containing the key \"TBLDESCR\". This will renumber all the existing names. Are you sure?",
                "Warning!!!", MessageBoxButtons.YesNo);

            if (Response == DialogResult.Yes)
            {
                foreach (Excel.Name TblName in wb.Names)
                {
                    if (TblName.Name.ToUpper().StartsWith("TABLEC") || TblName.Name.ToUpper().Contains("!TABLEC"))
                    {
                        if (xlApp.WorksheetFunction.IsErr(TblName.Name)) //(TblName.Name.Contains(XlCVError.xlErrRef))
                            TblName.Delete();
                        else
                            try
                            {
                                wrkShtName = wb.Names.Item(TblName.Name).RefersToRange.Parent.Name;
                                CalcSheet = wb.Worksheets[wrkShtName];
                                CalcSheet.Range[TblName].Name.Delete();
                            }
                            catch (Exception ex)
                            {
                                TblName.Delete();
                            }
                    }
                }

                foreach (Excel.Worksheet wrkSht in wb.Worksheets)
                {
                    CellRng = wrkSht.UsedRange.Find("TBLDESCR", LookIn: Excel.XlFindLookIn.xlValues);
                    if (CellRng != null)
                    {
                        StartAddress = CellRng.Address;
                        do
                        {
                            CellRng.Name = "TableC" + i;
                            CellRng = wrkSht.UsedRange.FindNext(CellRng);
                            i++;
                        } while (CellRng != null && CellRng.Address != StartAddress);
                    }

                    CellRng = wrkSht.UsedRange.Find("DESCRIPTION", LookIn: Excel.XlFindLookIn.xlValues);
                    if (CellRng != null)
                    {
                        StartAddress = CellRng.Address;
                        do
                        {
                            if (CellRng.Value.ToUpper() == "DESCRIPTION" &&
                                CellRng.Offset[0, 1].Value.ToUpper() == "PARAMETER")
                            {
                                CellRng.Name = "TableC" + i;
                                CellRng = wrkSht.UsedRange.FindNext(CellRng);
                                i++;
                            }
                            else
                            {
                                CellRng = wrkSht.UsedRange.FindNext(CellRng);
                            }
                        } while (CellRng != null && CellRng.Address != StartAddress);
                    }

                }
            }
        }


        public void AddCustomTable()
        {
            Excel.Range Rng = null;
            Excel.Worksheet TblSheet;
            bool WriteTbl;
            WriteTbl = true;
            long RowNdx;
            long ColNdx;
            long i;
            long j;
            string SheetName;
            string CellName;
            //Name TblName;
            int TblCount;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Range Selection;


            try
            {
                Rng = wb.Application.InputBox("Select the Start Cell of the Table.", "Obtain Range Object",
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type: 8);
            }
            catch (Exception ex)
            {
                if (Rng == null)
                {
                    MessageBox.Show("Cancelled by the user. No table added");
                    return;
                }
            }

            TblSheet = wb.Worksheets[Rng.Parent.Name];

            TblSheet.Select();
            RowNdx = Rng[1, 1].Row;
            ColNdx = Rng[1, 1].Column;

            SheetName = Rng.Parent.Name;
            CellName = Rng.Address.Replace("$", "");

            for (i = RowNdx; i <= RowNdx + 6; i++)
            {
                for (j = ColNdx; j <= ColNdx + 4; j++)
                {
                    if (TblSheet.Cells[i, j].Value != null)
                        WriteTbl = false;
                }
            }

            if (WriteTbl == true)
            {
                TblSheet.Cells[RowNdx, ColNdx] = "TBLDESCR";
                TblSheet.Cells[RowNdx + 1, ColNdx] = "PARAMETER";
                TblSheet.Cells[RowNdx + 2, ColNdx] = "SYMBOL";
                TblSheet.Cells[RowNdx + 3, ColNdx] = "UNIT";
                TblSheet.Cells[RowNdx + 4, ColNdx] = "REFERENCE";
                TblSheet.Cells[RowNdx + 5, ColNdx] = 1;
                TblSheet.Cells[RowNdx + 6, ColNdx] = 2;

                for (i = 1; i <= 4; i++)
                {
                    TblSheet.Cells[RowNdx, ColNdx + i] = "Description" + i;
                    TblSheet.Cells[RowNdx + 1, ColNdx + i] = "Param" + i;
                    TblSheet.Cells[RowNdx + 2, ColNdx + i] = "Symbol" + i;
                    TblSheet.Cells[RowNdx + 3, ColNdx + i] = "Unit" + i;
                    TblSheet.Cells[RowNdx + 4, ColNdx + i] = "-";
                }

                Selection = TblSheet.Range[TblSheet.Cells[RowNdx, ColNdx],
                    TblSheet.Cells[RowNdx + 6, ColNdx + 4]]; // .Select

                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                Selection.HorizontalAlignment = Excel.Constants.xlCenter;

                TblCount = 0;
                foreach (Excel.Name TblName in wb.Names)
                {
                    if (TblName.Name.ToUpper().StartsWith("TABLEC"))
                        TblCount++;
                }

                Rng.Name = "TableC" + TblCount + 1;

                MessageBox.Show($"The calculation table created in the sheet '{SheetName}' at the cell '{CellName}'.");
            }
            else
                MessageBox.Show(
                    "Error! The Range selected overwrites an existing data. Please select an empty area and insert the table again.");
        }

        public void AddCustomTableSingle()
        {
            Excel.Range Rng = null;
            Excel.Worksheet TblSheet;
            bool WriteTbl;
            WriteTbl = true;
            long RowNdx;
            long ColNdx;
            long i;
            long j;
            string SheetName;
            string CellName;
            //Name TblName;
            int TblCount;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Range Selection;

            try
            {
                Rng = wb.Application.InputBox("Select the Start Cell of the Table.", "Obtain Range Object",
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type: 8);
            }
            catch (Exception ex)
            {
                if (Rng == null)
                {
                    MessageBox.Show("Cancelled by the user. No table added");
                    return;
                }
            }

            TblSheet = wb.Worksheets[Rng.Parent.Name];

            TblSheet.Select();
            RowNdx = Rng[1, 1].Row;
            ColNdx = Rng[1, 1].Column;

            SheetName = Rng.Parent.Name;
            CellName = Rng.Address.Replace("$", "");

            for (i = RowNdx; i <= RowNdx + 3; i++)
            {
                for (j = ColNdx; j <= ColNdx + 5; j++)
                {
                    if (TblSheet.Cells[i, j].Value != null)
                        WriteTbl = false;
                }
            }

            if (WriteTbl == true)
            {
                TblSheet.Cells[RowNdx, ColNdx] = "DESCRIPTION";
                TblSheet.Cells[RowNdx, ColNdx + 1] = "PARAMETER";
                TblSheet.Cells[RowNdx, ColNdx + 2] = "SYMBOL";
                TblSheet.Cells[RowNdx, ColNdx + 3] = "VALUE";
                TblSheet.Cells[RowNdx, ColNdx + 4] = "UNIT";
                TblSheet.Cells[RowNdx, ColNdx + 5] = "REFERENCE";

                for (i = 1; i <= 2; i++)
                {
                    TblSheet.Cells[RowNdx + i, ColNdx] = "Description" + i;
                    TblSheet.Cells[RowNdx + i, ColNdx + 1] = "Param" + i;
                    TblSheet.Cells[RowNdx + i, ColNdx + 2] = "Symbol" + i;
                    TblSheet.Cells[RowNdx + i, ColNdx + 3] = "Value" + i;
                    TblSheet.Cells[RowNdx + i, ColNdx + 4] = "Unit" + i;
                    TblSheet.Cells[RowNdx + i, ColNdx + 5] = "-";
                }

                Selection = TblSheet.Range[TblSheet.Cells[RowNdx, ColNdx],
                    TblSheet.Cells[RowNdx + 2, ColNdx + 5]]; // .Select

                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                Selection.HorizontalAlignment = Excel.Constants.xlCenter;

                TblCount = 0;
                foreach (Excel.Name TblName in wb.Names)
                {
                    if (TblName.Name.ToUpper().StartsWith("TABLEC"))
                        TblCount++;
                }

                Rng.Name = "TableC" + TblCount + 1;

                MessageBox.Show($"The calculation table created in the sheet '{SheetName}' at the cell '{CellName}'.");
            }
            else
                MessageBox.Show(
                    "The Range selected overwrites an existing data. Please select an empty area and insert the table again.",
                    "Error!");
        }


        public static string GetParameterList(string TableId, string CaptionName = "")
        {
            string wrkShtName;
            Excel.Worksheet CalcSheet;
            Excel.Range ParamRng;
            int count = 0;
            //Excel.Range Rng=null;
            long StartRow;
            long StartColumn;
            string listSeparator;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ActiveWorksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            string ParameterList;

            listSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ListSeparator;

            try
            {
                wrkShtName = wb.Names.Item(TableId).RefersToRange.Parent.Name;
                CalcSheet = wb.Worksheets[wrkShtName];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show(
                    $"The Table Selected refers to a deleted table or deleted cell. Please Renumber the tables to clean the list.\n{ex.Message}");
                return "";
            }

            ParameterList = "";
            CalcSheet.Select();
            CalcSheet.Range[TableId].Select();

            if (CalcSheet.Range[TableId].Value == "TBLDESCR")
            {
                StartRow = CalcSheet.Range[TableId].Row + 1;
                StartColumn = CalcSheet.Range[TableId].Column + 1;
                ParamRng = CalcSheet.Range[CalcSheet.Range[TableId].Offset[1, 1],
                    CalcSheet.Range[TableId].Offset[1, 1].End[Excel.XlDirection.xlToRight]];
            }
            else if (CalcSheet.Range[TableId].Value == "DESCRIPTION")
            {
                StartRow = CalcSheet.Range[TableId].Row + 1;
                StartColumn = CalcSheet.Range[TableId].Column + 1;
                ParamRng = CalcSheet.Range[CalcSheet.Range[TableId].Offset[1, 1],
                    CalcSheet.Range[TableId].Offset[1, 1].End[Excel.XlDirection.xlDown]];
            }
            else
            {
                StartRow = CalcSheet.Range[TableId].Row;
                StartColumn = CalcSheet.Range[TableId].Column;
                ParamRng = CalcSheet.Range[CalcSheet.Range[TableId],
                    CalcSheet.Range[TableId].End[Excel.XlDirection.xlToRight]];
            }

            if (CalcSheet.Range[TableId].Row > 1)
                CaptionName = CalcSheet.Range[TableId].Offset[-1, 0].Value;

            foreach (Excel.Range Rng in ParamRng)
                ParameterList = ParameterList + listSeparator + Rng.Value;
            ParameterList = ParameterList.Substring(1, ParameterList.Length - 1);

            return ParameterList;
        }


        public bool IsFileOpen(FileInfo file)
        {
            FileStream myStream = null;
            bool StatusFileOpen = false;
            try
            {
                myStream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                myStream.Close();
            }
            catch (Exception ex)
            {
                if (!file.Exists)
                    StatusFileOpen = false;
                else
                    StatusFileOpen = true;
            }

            return StatusFileOpen;
        }


        private void TblofContents(Word.Application wrdApp, Word.Document wrdDoc)
        {
            Word.Range rngDocStart;


            wrdApp.Selection.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

            wrdApp.Selection.TypeParagraph();
            wrdApp.Selection.MoveLeft(Word.WdUnits.wdCharacter, 1);
            wrdApp.Selection.set_Style(wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleNormal].NameLocal);

            wrdApp.Selection.TypeParagraph();
            wrdApp.Selection.InsertBreak(Type: Word.WdBreakType.wdSectionBreakNextPage);

            wrdApp.Selection.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
            wrdApp.Selection.TypeParagraph();
            wrdApp.Selection.set_Style(wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleTocHeading].NameLocal);
            //wrdApp.Selection.Range.ListFormat.RemoveNumbers(NumberType: Word.WdNumberType.wdNumberParagraph);
            wrdApp.Selection.TypeText("TABLE OF CONTENTS");

            wrdApp.Selection.TypeParagraph();
            wrdApp.Selection.Collapse();

            rngDocStart = wrdApp.Selection.Range;
            rngDocStart.Collapse();

            wrdDoc.TablesOfContents.Add(Range: rngDocStart, RightAlignPageNumbers: true, UseHeadingStyles: true,
                IncludePageNumbers: true, AddedStyles: "styleSection", UseHyperlinks: false, HidePageNumbersInWeb: true,
                UseOutlineLevels: true);
            wrdDoc.TablesOfContents[1].Range.Font.Name = "Arial Narrow";
            wrdDoc.TablesOfContents[1].Range.Font.Size = 11;
            wrdDoc.TablesOfContents[1].TabLeader = Word.WdTabLeader.wdTabLeaderDots;
            wrdDoc.TablesOfContents.Format = Word.WdTocFormat.wdTOCSimple;


            //wrdApp.Selection.Collapse();
            //wrdApp.Selection.InsertBreak(Type: Word.WdBreakType.wdSectionBreakNextPage);
        }

        public static List<string> GetChartList()
        {
            //Excel.ChartObject xlChartObject;
            Excel.ChartObject xlChart;
            Excel.Chart Chartpage;
            //Excel.Worksheet WrkSheet;
            int Count;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            List<string> PictureList = new List<string>();

            foreach (Excel.Worksheet WrkSheet in wb.Worksheets)
            {
                Count = 1;
                // WrkSheet.Activate()
                foreach (Excel.ChartObject xlChartObject in WrkSheet.ChartObjects())
                {
                    xlChart = (Excel.ChartObject)WrkSheet.ChartObjects(Count);
                    Chartpage = xlChart.Chart;
                    PictureList.Add(WrkSheet.Name + "|" + xlChart.Name);
                    Count++;
                }
            }

            return PictureList;
        }


        private Dictionary<string, dynamic> InputDataTable(Excel.Worksheet ReportSheet)
        {
            string TableRef;
            string ParamKey;
            Dictionary<string, dynamic> TableDict = new Dictionary<string, dynamic>();
            Dictionary<string, string> ParamDict;
            Dictionary<string, dynamic> DataDict = new Dictionary<string, dynamic>();
            long StartRow;
            long RowNdx;
            long ColNdx;
            string StartCell;

            StartCell = GetCellNameReport(CellNameReport.NAME_CELL_START_TABLE);
            StartRow = ReportSheet.Range[StartCell].Row;
            RowNdx = StartRow + 1;
            ColNdx = ReportSheet.Range[StartCell].Column;
            while (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx].Value))
            {
                ParamDict = new Dictionary<string, string>();

                for (int i = 0; i <= 4; i++)
                {
                    ParamKey = ReportSheet.Cells[StartRow, ColNdx + i].Text;
                    ParamDict.Add(ParamKey, ReportSheet.Cells[RowNdx, ColNdx + i].Text);
                }

                TableRef = ReportSheet.Cells[RowNdx, ColNdx].Value;
                TableDict.Add(TableRef, ParamDict);
                RowNdx++;
            }

            DataDict.Add("TABLES", TableDict);

            StartCell = GetCellNameReport(CellNameReport.NAME_CELL_START_FIGURE);
            StartRow = ReportSheet.Range[StartCell].Row;
            RowNdx = StartRow + 1;
            ColNdx = ReportSheet.Range[StartCell].Column;
            TableDict = new Dictionary<string, dynamic>();
            while (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx].Value))
            {
                ParamDict = new Dictionary<string, string>();
                for (int i = 0; i <= 3; i++)
                {
                    ParamKey = ReportSheet.Cells[StartRow, ColNdx + i].Text;
                    ParamDict.Add(ParamKey, ReportSheet.Cells[RowNdx, ColNdx + i].Value);
                }
                TableRef = ReportSheet.Cells[RowNdx, ColNdx].Value;
                TableDict.Add(TableRef, ParamDict);
                RowNdx++;
            }
            DataDict.Add("FIGURES", TableDict);
            return DataDict;
        }

        public bool WorksheetExists(string WorksheetName, Excel.Workbook wb)
        {
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            //bool sheetExists = false;
            //Worksheet Sht;
            foreach (Excel.Worksheet Sheet in wb.Worksheets)
            {
                //if (xlApp.Application.Proper(Sht.Name) == xlApp.Application.Proper(WorksheetName))
                if (Sheet.Name == WorksheetName)
                {
                    return true;
                }
            }

            return false;
        }

        private bool CheckSummaryTable(ref string TableList)
        {
            long RowNDx, ColNdx;
            string[] Request;
            string ChkParameter;
            bool CkeckTable;
            string CalcParameters, TblParameters, ParamMissingList = "";
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            //string listSeparator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            Excel.Worksheet ReportSheet;
            Excel.Worksheet CalcSheet;
            string TableRef, wrkShtName;
            bool ParamCheck = true;
            string[] TableRequest;
            string Item;
            bool ItemCheck = true, RequestIDCheck, CalcTableFormat = true;

            ReportSheet = wb.Worksheets[GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT)];
            RowNDx = ReportSheet.Range[GetCellNameReport(CellNameReport.NAME_CELL_START_TABLE)].Row + 1;
            ColNdx = ReportSheet.Range[GetCellNameReport(CellNameReport.NAME_CELL_START_TABLE)].Column;

            if (string.IsNullOrEmpty(ReportSheet.Cells[RowNDx, ColNdx].Value))
            {
                MessageBox.Show("The Report Summary Table cannot be empty.");
                return false;
            }

            while (!string.IsNullOrEmpty(ReportSheet.Cells[RowNDx, ColNdx].Value))
            {
                TableRef = ReportSheet.Cells[RowNDx, ColNdx].Value;

                try
                {
                    wrkShtName = wb.Names.Item(TableRef).RefersToRange.Parent.Name;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Please verify the Table ID: {0}. It is not found in the Excel Workbook" , TableRef));
                    CkeckTable = false;
                    RowNDx++;
                    continue;
                }

                
                CalcSheet = wb.Worksheets[wrkShtName];
                CalcSheet.Select();

                if (CalcSheet.Range[TableRef].Value != "TBLDESCR" || CalcSheet.Range[TableRef].Value != "DESCRIPTION")
                    CalcTableFormat = false;

                CalcParameters = ReportSheet.Cells[RowNDx, ColNdx + 1].Value;

                if (CalcParameters != null)
                {
                    Request = CalcParameters.Split(listSeparator);
                    if (CalcParameters.Contains("|"))
                    {
                        ParamCheck = false;
                        MessageBox.Show($"{TableRef} is not allowed in the calculation parameter list for the | ");
                    }
                }
                else
                {
                    CalcParameters = ReportSheet.Cells[RowNDx, ColNdx + 2].Value;
                    Request = CalcParameters.Split(listSeparator);
                }

                TblParameters = ReportSheet.Cells[RowNDx, ColNdx + 2].Value;
                TblParameters = TblParameters.Replace("|", "");
                TableRequest = TblParameters.Split(listSeparator);

                ChkParameter = GetParameterList(TableRef);
                CkeckTable = true;
                for (int i = 0; i < Request.Length; i++)
                {
                    if (!ChkParameter.Contains(Request[i]))
                    {
                        ParamCheck = false;
                        CkeckTable = false;
                    }
                }

                for (int i = 0; i < TableRequest.Length; i++)
                {
                    if (!ChkParameter.Contains(TableRequest[i]))
                    {
                        ParamCheck = false;
                        CkeckTable = false;
                    }
                }

                Item = Convert.ToString(ReportSheet.Cells[RowNDx, ColNdx + 3].Value);
                if (CalcTableFormat)
                {
                    if (string.IsNullOrEmpty(Item))
                    {
                        ItemCheck = false;
                        CkeckTable = false;
                    }
                    else
                    {
                        RequestIDCheck = GetItemList(TableRef, Item);
                        if (!RequestIDCheck)
                        {
                            MessageBox.Show($"The requested ID \"{Item}\" could not be found in {TableRef}",
                                "Input Error!");
                            CkeckTable = false;
                        }
                    }
                }

                if (!CkeckTable)
                {
                    if (TableList == "")
                        TableList = TableRef;
                    else
                        TableList = TableList + "," + TableRef;
                }

                RowNDx++;
                CalcTableFormat = true;
            }

            if (!ItemCheck || TableList != "")
            {
                MessageBox.Show(
                    $"The Requested parameters/Item for the below tables do not match with the table.\n\n{TableList}\n\n Please check the parameters and rerun the program.",
                    "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return ParamCheck;
        }

        private bool ValidateInputs()
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string ReportSheetName = GetSheetNameReport(SheetNameReport.SHEET_NAME_REPORT);
            Excel.Worksheet ReportSheet;
            Dictionary<string, dynamic> DataDict;
            bool Status = true;
            long RowNdx;
            long ColNdx;
            string GroupID;
            string InvalidSets = "";

            ReportSheet = wb.Worksheets[ReportSheetName];
            DataDict = InputDataTable(ReportSheet);

            RowNdx = ReportSheet.Range[GetCellNameReport(CellNameReport.NAME_CELL_START_CONTENT)].Row + 1;
            ColNdx = ReportSheet.Range[GetCellNameReport(CellNameReport.NAME_CELL_START_CONTENT)].Column;

            while (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx].Text))
            {
                if (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx + 3].Text))
                {
                    GroupID = ReportSheet.Cells[RowNdx, ColNdx + 3].Text;
                    if (DataDict == null)
                    {
                        InvalidSets = InvalidSets + "," + GroupID;
                        Status = false;
                    }
                    else if (!((Dictionary<string, dynamic>)DataDict["FIGURES"]).ContainsKey(GroupID))
                    {
                        InvalidSets = InvalidSets + "," + GroupID;
                        Status = false;
                    }
                }

                if (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx + 5].Text))
                {
                    GroupID = ReportSheet.Cells[RowNdx, ColNdx + 5].Text;
                    if (DataDict == null)
                    {
                        InvalidSets = InvalidSets + "," + GroupID;
                        Status = false;
                    }
                    else if (!((Dictionary<string, dynamic>)DataDict["TABLES"]).ContainsKey(GroupID))
                    {
                        InvalidSets = InvalidSets + "," + GroupID;
                        Status = false;
                    }
                }

                // -----Tables
                if (!string.IsNullOrEmpty(ReportSheet.Cells[RowNdx, ColNdx + 6].Text))
                {
                    GroupID = ReportSheet.Cells[RowNdx, ColNdx + 6].Text;
                    if (DataDict == null)
                    {
                        InvalidSets = InvalidSets + "," + GroupID;
                        Status = false;
                    }
                    else if (!((Dictionary<string, dynamic>)DataDict["TABLES"]).ContainsKey(GroupID))
                    {
                        InvalidSets = InvalidSets + "," + GroupID;
                        Status = false;
                    }
                }



                RowNdx++;
            }

            if (!Status)
                MessageBox.Show(
                    $"The following Group IDs in the Report Content Table are not matching with the Input Group IDs.\n{InvalidSets}",
                    "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return Status;
        }

        private bool GetItemList(string TableId, string RequestItem)
        {
            string wrkShtName;
            Excel.Worksheet CalcSheet;
            Excel.Range ItemRng;
            int count = 0;
            Excel.Range Rng;
            long StartRow;
            long StartColumn;
            string listSeparator;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ActiveWorksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            string ItemList;
            bool ItemCheck = false;

            listSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ListSeparator;

            try
            {
                wrkShtName = wb.Names.Item(TableId).RefersToRange.Parent.Name;
                CalcSheet = wb.Worksheets[wrkShtName];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show(
                    $"The Table Selected refers to a deleted table or deleted cell. Please Renumber the tables to clean the list.\n{ex.Message}");
                return ItemCheck;
            }

            ItemList = "";
            CalcSheet.Select();
            CalcSheet.Range[TableId].Select();

            if (CalcSheet.Range[TableId].Value == "TBLDESCR")
            {
                StartRow = CalcSheet.Range[TableId].Row + 4;
                StartColumn = CalcSheet.Range[TableId].Column;
                if (CalcSheet.Cells[StartRow, StartColumn].Value == "REFERENCE")
                    StartRow++;
                Rng = CalcSheet.Cells[StartRow, StartColumn];
                if (!string.IsNullOrEmpty(Rng.Offset[1, 0].Value))
                    ItemRng = CalcSheet.Range[Rng, Rng.End[Excel.XlDirection.xlDown]];
                else
                    ItemRng = Rng;
            }
            else if (CalcSheet.Range[TableId].Value == "DESCRIPTION")
            {
                StartRow = CalcSheet.Range[TableId].Row;
                StartColumn = CalcSheet.Range[TableId].Column + 3;
                count = CalcSheet.Range[CalcSheet.Cells[StartRow, StartColumn],
                    CalcSheet.Cells[StartRow, StartColumn].End[Excel.XlDirection.xlToRight]].Count - 1;
                ItemRng = CalcSheet.Range[CalcSheet.Cells[StartRow, StartColumn],
                    CalcSheet.Cells[StartRow, StartColumn + count]];
            }
            else
                return ItemCheck;

            foreach (Excel.Range Rng1 in ItemRng)
            {
                if (IsNumeric(RequestItem))
                {
                    if (Rng1.Value == RequestItem)
                    {
                        ItemCheck = true;
                        break;
                    }
                }
                else if (CalcSheet.Range[TableId].Value == "TBLDESCR")
                {
                    if (Rng1.Offset[0, 1].Value == RequestItem)
                    {
                        ItemCheck = true;
                        break;
                    }
                }
                else if (CalcSheet.Range[TableId].Value == "DESCRIPTION")
                {
                    if (Rng1.Value == RequestItem)
                    {
                        ItemCheck = true;
                        break;
                    }
                }
            }

            return ItemCheck;
        }


        private void WriteListofAbbr(ref SortedDictionary<string, Dictionary<string, string>> LoADict, Word.Application wrdApp, Word.Document wrdDoc, Dictionary<string, string> DictAutoCorrect)
        {
            //string ParamKey;
            string Symb;
            string Unit;
            string Descr;
            Word.Range objRange;
            Excel.Application xlapp = Globals.ThisAddIn.Application;
            MathConverter ConvertMath = new MathConverter();

            xlapp.StatusBar = @"Preparing the List of Abbreviation...";


            /*objRange = wrdApp.ActiveDocument.Content;
            objRange.Find.Execute(FindText: "The list of abbreviations are presented below.", Forward: false);
            if (objRange.Find.Found)
            {
                objRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
                MessageBox.Show($"{objRange.Start}:{ objRange.End}");

                objRange.Select();
                //myRange.InsertParagraphAfter();
                wrdApp.Selection.TypeParagraph();
            }
            else
            {
                HeadingListLevel(wrdApp, 2);
                wrdApp.Selection.ParagraphFormat.set_Style(wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading2].NameLocal);
                wrdApp.Selection.TypeText(@"List of Abbreviations");
                wrdApp.Selection.TypeParagraph();
                wrdApp.Selection.TypeText(@"The list of abbreviations are presented below.");

                wrdApp.Selection.TypeParagraph();
            }*/

            HeadingListLevel(wrdApp, 2);
            wrdApp.Selection.ParagraphFormat.set_Style(wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading2].NameLocal);
            wrdApp.Selection.TypeText(@"List of Abbreviations");
            wrdApp.Selection.TypeParagraph();
            wrdApp.Selection.TypeText(@"The list of abbreviations are presented below.");

            wrdApp.Selection.TypeParagraph();

            wrdApp.Selection.Font.Bold = 1;
            wrdApp.Selection.Font.Italic = 1;
            wrdApp.Selection.Font.Underline = Word.WdUnderline.wdUnderlineWords;
            wrdApp.Selection.TypeText("PARAMETER \t\t UNIT \t\t\t DESCRIPTION");
            wrdApp.Selection.TypeParagraph();
            wrdApp.Selection.Font.Underline = 0;
            wrdApp.Selection.Font.Bold = 0;
            wrdApp.Selection.Font.Italic = 0;

            //int Count = 0;
            foreach (string ParamKey in LoADict.Keys)
            {
                if (ParamKey != "PARAMETER")
                {
                    Symb = LoADict[ParamKey]["SYMBOL"];
                    Unit = LoADict[ParamKey]["UNIT"];
                    Descr = LoADict[ParamKey]["DESCRIPTION"];

                    wrdApp.ActiveDocument.Characters.Last.Select();
                    wrdApp.Selection.Collapse();
                    wrdApp.Selection.TypeParagraph();


                    
                    if (ApplySuffix(wrdApp,Symb)==false)
                        ConvertMath.MathEquation(wrdApp, wrdDoc, Symb, DictAutoCorrect, "NO", true);
                    /*if (Symb.Contains(@"\") || Symb.Contains(@"^") || Symb.Contains(@"_"))
                    {

                    }
                    else
                    {
                        wrdApp.ActiveDocument.Characters.Last.Select();
                        wrdApp.Selection.Collapse();
                        wrdApp.Selection.TypeText(Symb);
                    }*/
                    

                    objRange = wrdDoc.Range();
                    wrdApp.ActiveDocument.Characters.Last.Select();

                    wrdApp.Selection.Collapse();

                    objRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
                    objRange.MoveEnd();
                    objRange.InsertParagraphAfter();
                    objRange.MoveEnd(Unit: Word.WdUnits.wdCharacter, Count: -1);
                    if (Symb.Length <= 8)
                        objRange.InsertAfter("\t\t\t\t");
                    else
                        objRange.InsertAfter("\t\t\t");

                    objRange.MoveEnd(Unit: Word.WdUnits.wdCharacter, Count: -4);
                    objRange.Delete(Unit: Word.WdUnits.wdCharacter, Count: 1);
                    objRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);


                    wrdApp.ActiveDocument.Characters.Last.Select();
                    wrdApp.Selection.Collapse();

                    if (string.IsNullOrEmpty(Unit))
                        Unit = "-";
                    /*else
                    {
                        Unit = Unit.Replace("[", "").Replace("]", "");
                        if (Unit.ToUpper() == "MPA")
                            Unit = "MPa";
                    }*/

                    if (ApplySuffix(wrdApp, Unit) == false)
                    {
                        ConvertMath.MathEquation(wrdApp, wrdDoc, Unit, DictAutoCorrect, "NO", true);
                    }
                    
                    /*if (Unit.Contains(@"\") || Unit.Contains(@"^"))
                    {
                        ConvertMath.MathEquation(wrdApp, wrdDoc, Unit, DictAutoCorrect, "NO", true);
                    }
                    else
                    {
                        wrdApp.ActiveDocument.Characters.Last.Select();
                        wrdApp.Selection.Collapse();
                        wrdApp.Selection.TypeText(Unit);
                    }*/

                    objRange = wrdDoc.Range();
                    wrdApp.ActiveDocument.Characters.Last.Select();

                    wrdApp.Selection.Collapse();


                    objRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
                    objRange.MoveEnd(Unit: Word.WdUnits.wdCharacter, Count: 1);
                    objRange.InsertParagraphAfter();
                    objRange.InsertAfter("\t\t\t");
                    objRange.MoveEnd(Unit: Word.WdUnits.wdCharacter, Count: -4);
                    objRange.Delete(Unit: Word.WdUnits.wdCharacter, Count: 1);
                    objRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);

                    wrdApp.ActiveDocument.Characters.Last.Select();
                    wrdApp.Selection.Collapse();

                    wrdApp.Selection.TypeText(Descr);
                }
            }

            wrdApp.Selection.TypeParagraph();

            wrdApp.Selection.TypeParagraph();
            //wrdApp.Selection.InsertBreak(Type: Word.WdBreakType.wdSectionBreakNextPage);
        }

        private void writeTableofAbbr(ref SortedDictionary<string, Dictionary<string, string>> LoADict, Word.Application wrdApp, Word.Document wrdDoc, Dictionary<string, string> DictAutoCorrect)
        {
            //string ParamKey;
            string Symb;
            string Unit;
            string Descr;
            Word.Range objRange;
            Excel.Application xlapp = Globals.ThisAddIn.Application;
            //MathConverter ConvertMath = new MathConverter();
            StringBuilder TblText = new StringBuilder();
            Word.Table wrdTbl;
            Word.Range TblRange;
            object tab = Word.WdTableFieldSeparator.wdSeparateByTabs;

            xlapp.StatusBar = @"Preparing the List of Abbreviation...";

            HeadingListLevel(wrdApp, 2);
            wrdApp.Selection.ParagraphFormat.set_Style(wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleHeading2].NameLocal);
            wrdApp.Selection.TypeText(@"List of Abbreviations");
            wrdApp.Selection.TypeParagraph();
            wrdApp.Selection.TypeText(@"The list of abbreviations are presented below.");

            wrdApp.Selection.TypeParagraph();

            TblRange = wrdDoc.Range();
            wrdApp.ActiveDocument.Characters.Last.Select();
            wrdApp.Selection.Collapse();

            TblRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
            TblRange.MoveEnd();
         

            TblText.Append("PARAMETER \t UNIT \t DESCRIPTION\n");

            foreach (string ParamKey in LoADict.Keys)
            {
                if (ParamKey != "PARAMETER")
                {
                    Symb = LoADict[ParamKey]["SYMBOL"];
                    Unit = LoADict[ParamKey]["UNIT"];
                    Descr = LoADict[ParamKey]["DESCRIPTION"];

                    if (string.IsNullOrEmpty(Unit))
                        Unit = "-";

                    TblText.Append($"{Symb} \t {Unit} \t {Descr}\n");

                    
                }
            }
            TblRange.Text = TblText.ToString();
            wrdTbl = TblRange.ConvertToTable(Separator: ref tab);

            wrdTbl.Borders.Enable = 1;

            wrdTbl.Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            wrdTbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
            wrdTbl.Columns.AutoFit();
            wrdTbl.Range.Font.Size = 10;
            wrdTbl.Rows.Height = 14;
            wrdTbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;


            wrdApp.ActiveDocument.Characters.Last.Select();


            TblRange = wrdTbl.Range;

            TblRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            TblRange.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        }

        private bool ApplySuffix(Word.Application wrdApp, string Parameter)
        {
            bool success = true;
            string[] Params;

            if (!(Parameter.Contains(@"\") || Parameter.Contains(@"^") || Parameter.Contains(@"_")))
            {
                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();
                wrdApp.Selection.TypeText(Parameter);
                return success;
            }

            if (Parameter.Contains(@"\"))
                    return !success;

            if (Parameter.Contains(@"_") && Parameter.Contains(@"^"))
                return !success;

            if (Parameter.Contains(@"_"))
            {
                Params = Parameter.Split('_');
                if(Params.Length>2)
                    return !success;

                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();
                wrdApp.Selection.TypeText(Params[0]);
                wrdApp.Selection.Font.Subscript = 1;
                if (Params[1].StartsWith("(") && Params[1].EndsWith(")"))
                    Params[1] = Params[1].Substring(1, Params[1].Length - 2);
                wrdApp.Selection.TypeText(Params[1]);
                wrdApp.Selection.Font.Subscript = 0;
                return success;
            }

            if (Parameter.Contains(@"^"))
            {
                Params = Parameter.Split('^');
                if (Params.Length > 2)
                    return !success;

                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();
                wrdApp.Selection.TypeText(Params[0]);
                wrdApp.Selection.Font.Superscript = 1;
                if (Params[1].StartsWith("(") && Params[1].EndsWith(")"))
                    Params[1] = Params[1].Substring(1, Params[1].Length - 2);
                wrdApp.Selection.TypeText(Params[1]);
                wrdApp.Selection.Font.Superscript = 0;
                return success;
            }

            return !success;
        }
    }
}
