using StressUtilities;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

/*
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

namespace Report
{
    class Reference
    {
        public Reference()
        {

        }

        public void InsertRefTable()
        {
            string RefSheetName;
            Excel.Worksheet wrkSheet;
            bool SheetExistChk;
            Excel.Worksheet CurrentSheet;
            Excel.Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            Excel.Range Selection;
            CurrentSheet = wb.Worksheets[ActiveSheet.Name];
            //WriteReport ReportData = new WriteReport();

            RefSheetName = WriteReport.GetSheetNameReport(SheetNameReport.SHEET_NAME_REFERENCE);
            SheetExistChk = false;
            foreach (Excel.Worksheet Sheet in wb.Worksheets)
            {
                //if (xlApp.Application.Proper(Sheet.Name) == xlApp.Application.Proper(RefSheetName))
                if (Sheet.Name == RefSheetName)
                    SheetExistChk = true;
            }

            if (SheetExistChk == false)
            {

                wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]).Name = RefSheetName;

                wrkSheet = wb.Worksheets[RefSheetName];

                wrkSheet.Range["B2"].Value = "LIST OF REFERENCES";
                wrkSheet.Range["B2"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                wrkSheet.Range["B2"].Font.Name = "Arial";
                wrkSheet.Range["B3"].Value = "ID";
                wrkSheet.Range["B4"].Value = 1;
                wrkSheet.Range["B5"].Value = 2;
                wrkSheet.Range["C3"].Value = "TITLE";
                wrkSheet.Range["D3"].Value = "REFERENCE";
                wrkSheet.Range["E3"].Value = "ISSUE";
                wrkSheet.Range["F3"].Value = "DATE";
                wrkSheet.Range["G3"].Value = "AUTHOR";
                Selection = wrkSheet.Range["B3:G40"];
                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                Selection.Font.Name = "Arial";

                Selection = wrkSheet.Range["B3:G3"];
                Selection.Font.Name = "Arial Black";
                Selection.Font.Size = 8;
                Selection.Font.Italic = true;

                Selection = wrkSheet.Range["F4:F40"];
                Selection.NumberFormat = "dd mmm yyyy";


                wrkSheet.Columns["B"].ColumnWidth = 5;
                wrkSheet.Columns["C"].ColumnWidth = 45;
                wrkSheet.Columns["D"].ColumnWidth = 25;
                wrkSheet.Columns["E"].ColumnWidth = 10;
                wrkSheet.Columns["F"].ColumnWidth = 10;
                wrkSheet.Columns["F"].ColumnWidth = 20;

                wrkSheet.Range["B3"].Name = "REFTABLE";
                wrkSheet.Select();
                wb.Windows[1].DisplayGridlines = false;
                CurrentSheet.Select();
            }
            else
                MessageBox.Show("Reference Table already exists. No Table Created");
        }


        public void wrdRefTable(Word.Application wrdApp, Word.Document wrdDoc, Dictionary<string, dynamic> RefDict)
        {
            string ReferenceText;

            wrdApp.ActiveDocument.Characters.Last.Select();
            wrdApp.Selection.Collapse();

            //wrdApp.Selection.set_Style = wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleTocHeading].NameLocal;
            wrdApp.Selection.set_Style(wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleTocHeading].NameLocal);
            //wrdApp.Selection.Range.ListFormat.RemoveNumbers(NumberType: Word.WdNumberType.wdNumberParagraph);
            wrdApp.Selection.TypeText("LIST OF REFERENCES");
            
            wrdApp.Selection.TypeParagraph();
            
            wrdApp.Selection.set_Style(wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleNormal].NameLocal);

            wrdApp.Selection.Collapse();

            wrdReferenceList(wrdApp);

            foreach (string Item in RefDict.Keys)
            {
                if (RefDict[Item]["TITLE"] != "TITLE")
                {
                    ReferenceText = $"{RefDict[Item]["TITLE"]}, {RefDict[Item]["REFERENCE"]}, Issue {RefDict[Item]["ISSUE"]}, {RefDict[Item]["DATE"]}.";
                    wrdApp.Selection.TypeText(ReferenceText);
                    wrdApp.Selection.TypeParagraph();
                }
            }

            wrdApp.Selection.set_Style(wrdApp.ActiveDocument.Styles[Word.WdBuiltinStyle.wdStyleNormal].NameLocal);
            wrdApp.Selection.TypeParagraph();
            wrdApp.Selection.InsertBreak(Type: Word.WdBreakType.wdSectionBreakNextPage);
        }


        public void wrdReferenceList(Word.Application wrdApp)
        {
            Word.ListTemplate ListTemp;

            ListTemp = wrdApp.ListGalleries[Word.WdListGalleryType.wdNumberGallery].ListTemplates[1];

            ListTemp.ListLevels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            ListTemp.ListLevels[1].NumberFormat = "[%1].";
            ListTemp.ListLevels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;

            ListTemp.ListLevels[1].NumberPosition = wrdApp.CentimetersToPoints((float)0.63);
            ListTemp.ListLevels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            ListTemp.ListLevels[1].TextPosition = wrdApp.CentimetersToPoints((float)1.5);
            ListTemp.ListLevels[1].ResetOnHigher = 0; //false
            ListTemp.ListLevels[1].StartAt = 1;

            wrdApp.ListGalleries[Word.WdListGalleryType.wdNumberGallery].ListTemplates[1].Name = "";
            wrdApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate: ListTemp);

            //ListTemp = null;
        }

        public void InsertCrossRef(Word.Application wrdApp, Word.Document wrdDoc, ref Dictionary<string, dynamic> FormulaTables, string Request,int StartIndex)
        {

            Word.Range wrdRng;
            /*= wrdApp.ActiveDocument.Content;
            wrdRng.Find.Execute(FindText: "LIST OF REFERENCES", Forward: false);
            wrdRng.Select();*/

            string RefItem = FormulaTables["REFERENCE"][Request];

            RefItem = RefItem.Replace("[", "").Replace("]", "");
            if (RefItem != "-" && !string.IsNullOrEmpty(RefItem))
            {
                RefItem = $"{int.Parse(RefItem)+ StartIndex}";
                wrdRng = wrdDoc.Range();
                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();

                wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
                wrdRng.MoveEnd();
                wrdRng.InsertParagraphAfter();
                wrdRng.MoveEnd(Unit: Word.WdUnits.wdCharacter, Count: -1);
                wrdRng.InsertAfter("\t Ref. ");
                wrdRng.MoveEnd(Unit: Word.WdUnits.wdCharacter, Count: -1 * ("\t Ref. ").Length - 1);
                wrdRng.Delete(Unit: Word.WdUnits.wdCharacter, Count: 1);
                wrdRng.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);

                wrdApp.ActiveDocument.Characters.Last.Select();
                wrdApp.Selection.Collapse();

                wrdApp.Selection.InsertCrossReference(ReferenceType: Word.WdReferenceType.wdRefTypeNumberedItem, ReferenceKind: Word.WdReferenceKind.wdNumberNoContext, ReferenceItem: RefItem, InsertAsHyperlink: true, IncludePosition: false, SeparateNumbers: false, SeparatorString: " ");
                wrdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }
    }
}
