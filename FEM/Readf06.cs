using StressUtilities;
using StressUtilities.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/** 
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

namespace FEM
{
    class Readf06
    {
        bool StatusFlag = false;
        //bool OverallStatus = false;

        static Regex regExHeaderGen;
        static Regex ElmResults = new Regex(@"^\s+(\d+)\s+[\w\W\s\S\d\D]*$");
        static Regex ElmResultsZero = new Regex(@"^0\s+(\d+)\s+CEN[\w\W\s\S\d\D]*$");
        static Regex ElmResultsLayered = new Regex(@"0\s+(\d+)\s+(\d+)\s+(-?\d+\.\d+)[\w\W\s\S\d\D]*$");
        static Regex regExSubCase = new Regex(@"^0\s+[\w\W\s\S\d\D]*SUBCASE\s([A-Z 0-9 a-z]*)$");
        static Regex regExCROD = new Regex(@"\s+(\d+)([\s]{13}|\s+-?\d+\.\d+E[+-]?\d+\s+)(-?\d+\.\d+E[+-]?\d+)", RegexOptions.IgnorePatternWhitespace);
        static Regex regExCBEAMElem = new Regex(@"^0\s+(\d+)$");
        static Regex regExBeamResults = new Regex(@"^\s+\d+\s+\d+\.d+\s+(-?\d+\.\d+E[+-]?\d+)\s+(-?\d+\.\d+E[+-]?\d+)\s+(-?\d+\.\d+E[+-]?\d+)");
        static Regex regExHeader;
        static Regex ColumnHeaders;
        static Regex regExSolidElem;
        static Regex regExSolidResults;

        public void LaunchF06Form()
        {
            IEnumerable<ImportF06Form> FrmCollection = Application.OpenForms.OfType<ImportF06Form>();

            if (FrmCollection.Any())
                FrmCollection.First().Focus();
            else
            {
                ImportF06Form F06form = new ImportF06Form();
                F06form.Show();
            }
        }


        public void f06Read(List<string> f06FileList, string Request, string ElemList)
        {
            // AnalysisType string

            Excel.Range Rng;
            // f06FileList() string
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            //Dictionary<string, object> f06ElemDict;
            List<long> ElementList;
            //General CommonData = new General();
            // StatusFlag bool = false;

            ElementList = General.GetEntityList(ElemList);

            if (ElementList.Count == 0)
            {
                MessageBox.Show(@"The Entity List is empty.");
                return;
            }

            try
            {
                Rng = wb.Application.InputBox("Select the Start Cell for populating the results.", "Obtain Range Object", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            }
            catch (Exception Ex)
            {
                Rng = null;
            }


            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            xlApp.ScreenUpdating = false; 

            if (Rng == null)
                MessageBox.Show(@"Cancelled by the user. File not imported");
            else
            {
                //ElementList
                Importf06filesGeneral(Rng, f06FileList, Request, ref ElementList);// , StatusFlag
            }

            // f06ElemDict = Nothing
            ElementList = null;
            xlApp.StatusBar = null;
            xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            xlApp.ScreenUpdating = true; 

            if (StatusFlag == true)
                MessageBox.Show(@"Nastran .f06 File(s) Imported Successfully.");
            else
                MessageBox.Show(@"Could not import the results from Nastran .f06 File(s). Entity ID is incorrect or the element type or requested type is not supported.");
        }

        public void Importf06filesGeneral(Excel.Range Rng, List<string> f06FileList, string Request, ref List<long> ElementList)
        {
            string TextLine = "", PreviousLine, ElemType = "", ElementID;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            //bool ElPosIDChk = true;

            MatchCollection matches;
            string StringSearch; //, SearchHeader = "", PreviousHeader = "";
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            bool ExtractData;
            string Subcase = "", SCDescription = "";
            string[] ColumnHeadings = { }; 
            List<string> Results = new List<string>(), ResultsLayer2 = new List<string>();
            int StartIndex, Count;
            //Excel.Worksheet F06Sheet;
            //long RowNdx, ColNdx, StartCol, StartRow;
            bool ZeroLock = false, Layer2 = false;
            int FileCount;

            //string[] CRODResults;
            int delta = 0;
            bool ZeroStart = false;
            string ElemTypeFixed = "";
            bool streamBlock = true;
            //ulong LineCount;

            Dictionary<string, string> f06CompDict = new Dictionary<string, string>();
            Dictionary<string, dynamic> f06ElemDict = new Dictionary<string, dynamic>();

            regExSolidElem = new Regex(@"^0\s+(\d+)\s+0GRID\s+CS\s+\d+\s+GP");
            regExSolidResults = new Regex(@"[XYZ]\s+(-?\d+\.\d+E[+-]?\d+)\s+[XYZ]+\s+(-?\d+\.\d+E[+-]?\d+)\s+[ABC]\s+(-?\d+\.\d+E[+-]?\d+)\s+[LXYZ]+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+\s+(-?\d+\.\d+E[+-]?\d+)\s+(-?\d+\.\d+E[+-]?\d+)|[XYZ]\s+(-?\d+\.\d+E[+-]?\d+)\s+[XYZ]+\s+(-?\d+\.\d+E[+-]?\d+)\s+[ABC]\s+(-?\d+\.\d+E[+-]?\d+)\s+[LXYZ]+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+", RegexOptions.IgnorePatternWhitespace);

            switch (Request)
            {
                case "STRESSES":
                    StringSearch = "S T R E S S E S";
                    regExHeader = new Regex($@"^\s+{StringSearch}\s+[\w\W\s\S\d\D]+\(([\w\W\s\S\d\D]*)\)");
                    regExHeaderGen = new Regex(@"^\s+(?:\S\s){3}[\w\W\s\S\d\D]*$"); //(@"^\s+\S\s\S\s\S\s[\w\W\s\S\d\D]*$")
                    ColumnHeaders = new Regex(@"^\s+ID[\w\W\s\S\d\D]*$");
                    regExSolidElem = new Regex(@"^0\s+(\d+)\s+0GRID\s+CS\s+\d+\s+GP");
                    regExSolidResults = new Regex(@"[XYZ]\s+(-?\d+\.\d+E[+-]?\d+)\s+[XYZ]+\s+(-?\d+\.\d+E[+-]?\d+)\s+[ABC]\s+(-?\d+\.\d+E[+-]?\d+)\s+[LXYZ]+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+\s+(-?\d+\.\d+E[+-]?\d+)\s+(-?\d+\.\d+E[+-]?\d+)|[XYZ]\s+(-?\d+\.\d+E[+-]?\d+)\s+[XYZ]+\s+(-?\d+\.\d+E[+-]?\d+)\s+[ABC]\s+(-?\d+\.\d+E[+-]?\d+)\s+[LXYZ]+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+", RegexOptions.IgnorePatternWhitespace);
                    break;
                case "STRAINS":
                    StringSearch = "S T R A I N S";
                    regExHeader = new Regex($@"^\s+{StringSearch}\s+[\w\W\s\S\d\D]+\(([\w\W\s\S\d\D]*)\)");
                    regExHeaderGen = new Regex(@"^\s+(?:\S\s){3}[\w\W\s\S\d\D]*$"); //(@"^\s+\S\s\S\s\S\s[\w\W\s\S\d\D]*$")
                    ColumnHeaders = new Regex(@"^\s+ID[\w\W\s\S\d\D]*$");
                    regExSolidElem = new Regex(@"^0\s+(\d+)\s+0GRID\s+CS\s+\d+\s+GP");
                    regExSolidResults = new Regex(@"[XYZ]\s+(-?\d+\.\d+E[+-]?\d+)\s+[XYZ]+\s+(-?\d+\.\d+E[+-]?\d+)\s+[ABC]\s+(-?\d+\.\d+E[+-]?\d+)\s+[LXYZ]+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+\s+(-?\d+\.\d+E[+-]?\d+)\s+(-?\d+\.\d+E[+-]?\d+)|[XYZ]\s+(-?\d+\.\d+E[+-]?\d+)\s+[XYZ]+\s+(-?\d+\.\d+E[+-]?\d+)\s+[ABC]\s+(-?\d+\.\d+E[+-]?\d+)\s+[LXYZ]+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+[\s]?-?\d+\.\d+", RegexOptions.IgnorePatternWhitespace);
                    break;
                case "FORCES":
                    StringSearch = "F O R C E S   I N";
                    regExHeader = new Regex($@"^\s+{StringSearch}\s+[\w\W\s\S\d\D]+\(([\w\W\s\S\d\D]*)\)");
                    regExHeaderGen = new Regex(@"^\s+(?:\S\s){3}[\w\W\s\S\d\D]*$"); //(@"^\s+\S\s\S\s\S\s[\w\W\s\S\d\D]*$")
                    ColumnHeaders = new Regex(@"^\s+ID[\w\W\s\S\d\D]*$");
                    regExCROD = new Regex(@"(\d+)\s+(-?\d+\.\d+E[+-]?\d+)\s+((-?\d+\.\d+)|(-?\d+\.\d+E[+-]?\d+))", RegexOptions.IgnorePatternWhitespace);
                    break;
                case "DISPLACEMENTS":
                    StringSearch = "D I S P L A C E M E N T";
                    regExHeader = new Regex($@"^\s+{StringSearch}\s+[\w\W\s\S\d\D]*$");
                    regExHeaderGen = new Regex(@"^\s+(?:\S\s){3}[\w\W\s\S\d\D]*$"); //(@"^\s+\S\s\S\s\S\s[\w\W\s\S\d\D]*$")
                    ColumnHeaders = new Regex(@"^\s+POINT\sID[\w\W\s\S\d\D]*$");
                    ElmResults = new Regex(@"^\s+(\d+)\s+\S\s+(-?\d+\.\d+)[\w\W\s\S\d\D]*$");
                    break;
                case "SPC FORCES":
                    StringSearch = "F O R C E S   O F   S I N G L E - P O I N T";
                    regExHeader = new Regex($@"^\s+{StringSearch}\s+[\w\W\s\S\d\D]*$");
                    regExHeaderGen = new Regex(@"^\s+(?:\S\s){3}[\w\W\s\S\d\D]*$"); //(@"^\s+\S\s\S\s\S\s[\w\W\s\S\d\D]*$")
                    ColumnHeaders = new Regex(@"^\s+POINT\sID[\w\W\s\S\d\D]*$");
                    break;
                case "LOADS":
                    StringSearch = "L O A D";
                    regExHeader = new Regex($@"^\s+{StringSearch}\s+[\w\W\s\S\d\D]*$");
                    regExHeaderGen = new Regex(@"^\s+(?:\S\s){3}[\w\W\s\S\d\D]*$"); //(@"^\s+\S\s\S\s\S\s[\w\W\s\S\d\D]*$")
                    ColumnHeaders = new Regex(@"^\s+POINT\sID[\w\W\s\S\d\D]*$");
                    break;
                default:
                    MessageBox.Show("Please select the Request Type");
                    return;
                    break;
            }

            Excel.Worksheet F06Sheet = wb.Worksheets[Rng.Parent.Name];
            F06Sheet.Select();
            long StartRow = Rng[1, 1].Row + 1;
            long StartCol = Rng[1, 1].Column;
            long ColNdx = StartCol;
            long RowNdx = StartRow;

            FileCount = f06FileList.Count();

            foreach (string f06File in f06FileList)
            {
                using (StreamReader ainp = new StreamReader(f06File))
                {
                    ExtractData = false;
                    //LineCount = 0
                    xlApp.StatusBar = "Processing File:" + new FileInfo(f06File).Name;
                    do
                    {
                        PreviousLine = TextLine;
                        TextLine = ainp.ReadLine();
                        //LineCount++;

                        if (regExSubCase.IsMatch(TextLine))
                        {
                            matches = regExSubCase.Matches(TextLine);
                            Subcase = matches[0].Groups[1].Value.Trim();
                            SCDescription = PreviousLine.Trim();
                        }

                        if (regExHeader.IsMatch(TextLine))
                        {
                            matches = regExHeader.Matches(TextLine);
                            ElemType = matches[0].Groups[1].Value;
                            ElemType = ElemType.Replace(" ", "");
                            if (TextLine.Trim().StartsWith(StringSearch))
                            {
                                ExtractData = true;
                            }
                            else
                            {
                                ExtractData = false;
                            }
                            ZeroStart = false;

                            if (ElemType != ElemTypeFixed && ElemTypeFixed != "")
                            {
                                if (FileCount == 1)
                                {
                                    WriteColumnHeadings(ColumnHeadings, Request, SCDescription, Layer2, ElemTypeFixed, F06Sheet, StartRow, StartCol, ColNdx);

                                    StartRow += 6;

                                    RowNdx = StartRow;

                                    ElemTypeFixed = ElemType;

                                    StatusFlag = false;

                                    ZeroLock = false;

                                    ZeroStart = false;
                                }

                                else
                                {
                                    streamBlock = false;

                                }
                            }

                            else
                            {
                                streamBlock = true;

                            }

                        }
                        else if (regExHeaderGen.IsMatch(TextLine))
                        {
                            ExtractData = false;

                        }


                        if (streamBlock == false)
                        {
                            continue; //continue do
                        }

                        if (ColumnHeaders.IsMatch(TextLine) && ExtractData == true)
                        {
                            matches = ColumnHeaders.Matches(TextLine);


                            ColumnHeadings = GetColumnHeadings(ref TextLine, ref ElemType, ref PreviousLine);

                        }

                        //  Read Stress && Strain data for the Shell Elements

                        if (regExSolidElem.IsMatch(TextLine) && ExtractData == true)
                        {
                            matches = regExSolidElem.Matches(TextLine);

                            ElementID = matches[0].Groups[1].Value;

                            if (ElementList.Contains(long.Parse(ElementID)))
                            {
                                if (ElemTypeFixed == "")
                                {
                                    ElemTypeFixed = ElemType;
                                }
                                f06ElemDict = new Dictionary<string, dynamic>();
                                f06CompDict = new Dictionary<string, string>();
                                int i = 1;
                                while (i <= 3)
                                {
                                    TextLine = ainp.ReadLine();
                                    if (regExSolidResults.IsMatch(TextLine))
                                    {
                                        matches = regExSolidResults.Matches(TextLine);
                                        switch (i)
                                        {
                                            case 1:
                                                f06CompDict.Add("X", matches[0].Groups[1].Value);
                                                f06CompDict.Add("XY", matches[0].Groups[2].Value);
                                                f06CompDict.Add("PRINCIPAL1", matches[0].Groups[3].Value);
                                                f06CompDict.Add("MEAN_PRESSURE", matches[0].Groups[4].Value);
                                                f06CompDict.Add("VON_MISES", matches[0].Groups[5].Value);
                                                break;
                                            case 2:
                                                f06CompDict.Add("Y", matches[0].Groups[6].Value);
                                                f06CompDict.Add("YZ", matches[0].Groups[7].Value);
                                                f06CompDict.Add("PRINCIPAL2", matches[0].Groups[8].Value);
                                                break;
                                            case 3:
                                                f06CompDict.Add("Z", matches[0].Groups[6].Value);
                                                f06CompDict.Add("ZX", matches[0].Groups[7].Value);
                                                f06CompDict.Add("PRINCIPAL3", matches[0].Groups[8].Value);
                                                break;
                                        }
                                        i++;
                                    }
                                }

                                f06ElemDict.Add(ElementID, f06CompDict);
                                F06Sheet.Cells[RowNdx + 1, ColNdx] = ElementID;
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 1] = f06ElemDict[ElementID]["X"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 2] = f06ElemDict[ElementID]["Y"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 3] = f06ElemDict[ElementID]["Z"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 4] = f06ElemDict[ElementID]["XY"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 5] = f06ElemDict[ElementID]["YZ"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 6] = f06ElemDict[ElementID]["ZX"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 7] = f06ElemDict[ElementID]["PRINCIPAL1"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 8] = f06ElemDict[ElementID]["PRINCIPAL2"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 9] = f06ElemDict[ElementID]["PRINCIPAL3"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 10] = f06ElemDict[ElementID]["MEAN_PRESSURE"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 11] = f06ElemDict[ElementID]["VON_MISES"];
                                F06Sheet.Cells[RowNdx + 1, ColNdx + 12] = Subcase;
                                if (SCDescription != "")
                                {
                                    F06Sheet.Cells[RowNdx + 1, ColNdx + 12] = SCDescription;
                                }
                                RowNdx++;
                                //f06CompDict = Nothing
                                //f06ElemDict = Nothing

                                ZeroLock = true;
                                StatusFlag = true;

                            }
                            ZeroStart = true;
                        }
                        else if (ElmResultsZero.IsMatch(TextLine) && ExtractData == true)
                        {
                            matches = ElmResultsZero.Matches(TextLine);
                            ElementID = matches[0].Groups[1].Value;
                            if (ElementList.Contains(long.Parse(ElementID)))
                            {
                                if (ElemTypeFixed == "")
                                {
                                    ElemTypeFixed = ElemType;
                                }
                                Results = TextLine.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                                TextLine = ainp.ReadLine();
                                if (ElmResultsZero.IsMatch(TextLine))
                                {
                                    if (ElementList.Contains(long.Parse(ElementID)))
                                    {
                                        Results = TextLine.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                                    }
                                }
                                else
                                {
                                    ResultsLayer2 = TextLine.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                                    if (ResultsLayer2.Count - 1 == Results.Count - 1 - 3)
                                    {
                                        StartIndex = Results.Count;
                                        Count = ResultsLayer2.Count;
                                        //ReDim Preserve Results(Results.Length + ResultsLayer2.Length - 1);
                                        //System.Array.Copy(ResultsLayer2, 0, Results, StartIndex, Count);
                                        Results.AddRange(ResultsLayer2);
                                    }

                                    if ((ElemType.Contains("QUAD") || ElemType.Contains("TRIA")) && Request != "FORCES")
                                    {
                                        Layer2 = true;

                                    }

                                }


                                for (int k = 1; k < Results.Count; k++)
                                {
                                    if (Layer2 == true)
                                    {
                                        F06Sheet.Cells[RowNdx + 1, ColNdx].Value = Results[k];
                                        ColNdx++;
                                    }
                                    else if (ColumnHeadings[k - 1] != "GRID-ID")
                                    {
                                        F06Sheet.Cells[RowNdx + 1, ColNdx].Value = Results[k];
                                        ColNdx++;
                                    }
                                }
                                F06Sheet.Cells[RowNdx + 1, ColNdx] = Subcase;
                                if (SCDescription != "")
                                {
                                    F06Sheet.Cells[RowNdx + 1, ColNdx + 1] = SCDescription;
                                }
                                ColNdx = StartCol;
                                RowNdx++;
                                StatusFlag = true;
                                ZeroLock = true;
                            }
                            ZeroStart = true;
                        }
                        else if (regExCBEAMElem.IsMatch(TextLine) && ExtractData == true)  //To be verified.
                        {
                            matches = regExCBEAMElem.Matches(TextLine);
                            ElementID = matches[0].Groups[1].Value;
                            if (ElementList.Contains(long.Parse(ElementID)))
                            {
                                if (ElemTypeFixed == "")
                                {
                                    ElemTypeFixed = ElemType;
                                }
                                F06Sheet.Cells[RowNdx + 1, ColNdx].Value = ElementID;
                                ColNdx++;
                                for (int k = 0; k <= 1; k++)
                                {
                                    TextLine = ainp.ReadLine();
                                    Results = TextLine.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();

                                    if (Request == "FORCES")
                                    {
                                        delta = 1;
                                    }
                                    else if (Request == "STRESSES" || Request == "STRAINS")
                                    {
                                        delta = 0;
                                    }
                                    /*for (int l = 0; l <= 7 + delta; l++)
                                    {
                                        F06Sheet.Cells[RowNdx + 1, ColNdx] = Results[l];
                                        ColNdx++;
                                    }*/
                                    F06Sheet.Range[F06Sheet.Cells[RowNdx + 1, ColNdx], F06Sheet.Cells[RowNdx + 1, ColNdx + 7 + delta]].Value = Results.GetRange(0, 7 + delta).ToArray();
                                    ColNdx += (7 + delta);
                                }
                                F06Sheet.Cells[RowNdx + 1, ColNdx].Value = Subcase;
                                if (SCDescription != "")
                                {
                                    F06Sheet.Cells[RowNdx + 1, ColNdx + 1].Value = SCDescription;
                                }
                                ColNdx = StartCol;
                                RowNdx++;
                                StatusFlag = true;
                                ZeroLock = true;
                            }
                            ZeroStart = true;
                        }
                        else if (ElmResultsLayered.IsMatch(TextLine) && ExtractData == true)
                        {
                            matches = ElmResultsLayered.Matches(TextLine);
                            ElementID = matches[0].Groups[1].Value;
                            //LayerID = matches[0].Groups[2].Value
                            if (ElementList.Contains(long.Parse(ElementID)))
                            {
                                if (ElemTypeFixed == "")
                                {
                                    ElemTypeFixed = ElemType;
                                }
                                Results = TextLine.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                                for (int k = 1; k < Results.Count; k++)
                                {
                                    F06Sheet.Cells[RowNdx + 1, ColNdx].Value = Results[k];
                                    ColNdx++;
                                }
                                F06Sheet.Cells[RowNdx + 1, ColNdx] = Subcase;
                                if (SCDescription != "")
                                {
                                    F06Sheet.Cells[RowNdx + 1, ColNdx + 1].Value = SCDescription;
                                }
                                ColNdx = StartCol;
                                RowNdx++;
                                StatusFlag = true;
                                ZeroLock = true;
                            }
                            ZeroStart = true;
                        }
                        else if (ElmResults.IsMatch(TextLine) && ExtractData == true && ZeroLock == false && ZeroStart == false)
                        {
                            if (ElemType == "CROD")
                            {
                                foreach (Match match in regExCROD.Matches(TextLine)) //Regex.Matches(TextLine.Trim(), patternCROD)
                                {
                                    ElementID = match.Groups[1].Value;
                                    if (ElementList.Contains(long.Parse(ElementID)))
                                    {
                                        if (ElemTypeFixed == "")
                                        {
                                            ElemTypeFixed = ElemType;
                                        }
                                        F06Sheet.Cells[RowNdx + 1, ColNdx].Value = ElementID; //Element ID
                                        F06Sheet.Cells[RowNdx + 1, ColNdx + 1].Value = match.Groups[2].Value; //Axial Force
                                        F06Sheet.Cells[RowNdx + 1, ColNdx + 2] = match.Groups[3].Value;  //Torque
                                                                                                         //F06Sheet.Cells[RowNdx + 1, ColNdx + 3) = match.Groups[4].Value  //Torque
                                        F06Sheet.Cells[RowNdx + 1, ColNdx + 3] = Subcase;
                                        if (SCDescription != "")
                                        {
                                            F06Sheet.Cells[RowNdx + 1, ColNdx + 4] = SCDescription;
                                        }
                                        RowNdx++;
                                        StatusFlag = true;
                                    }
                                }
                            }
                            else
                            {
                                matches = ElmResults.Matches(TextLine);
                                ElementID = matches[0].Groups[1].Value;
                                if (ElementList.Contains(long.Parse(ElementID)))
                                {
                                    if (ElemTypeFixed == "")
                                    {
                                        ElemTypeFixed = ElemType;
                                    }
                                    Results = TextLine.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                                    for (int k = 0; k < Results.Count; k++)
                                    {
                                        if (Layer2 == true)
                                        {
                                            F06Sheet.Cells[RowNdx + 1, ColNdx] = Results[k];
                                            ColNdx++;
                                        }
                                        else if (ColumnHeadings[k] != "GRID-ID")
                                        {
                                            F06Sheet.Cells[RowNdx + 1, ColNdx] = Results[k];
                                            ColNdx++;
                                        }
                                    }
                                    F06Sheet.Cells[RowNdx + 1, ColNdx].Value = Subcase;
                                    if (SCDescription != "")
                                    {
                                        F06Sheet.Cells[RowNdx + 1, ColNdx + 1].Value = SCDescription;
                                    }
                                    ColNdx = StartCol;
                                    RowNdx++;
                                    StatusFlag = true;
                                }
                            }
                        }
                    } while (ainp.Peek() >= 0);
                }
                ExtractData = false;
            }

            WriteColumnHeadings(ColumnHeadings, Request, SCDescription, Layer2, ElemTypeFixed, F06Sheet, StartRow, StartCol, ColNdx);


            xlApp.StatusBar = false;

        }

        private void WriteColumnHeadings(string[] ColumnHeadings, string Request, string SCDescription, bool Layer2, string ElemTypeFixed, Excel.Worksheet F06Sheet, long StartRow, long StartCol, long ColNdx)
        {

            if (StatusFlag == true)
            {
                ColumnHeadingsFinal(ref ColumnHeadings, ref Request, ref SCDescription, ref Layer2, ref ElemTypeFixed);

                for (int i = 0; i < ColumnHeadings.Length; i++)
                {
                    if (Layer2 == true)
                    {
                        F06Sheet.Cells[StartRow, ColNdx] = ColumnHeadings[i];
                        ColNdx++;
                    }
                    else if (ColumnHeadings[i] != "GRID-ID")
                    {
                        F06Sheet.Cells[StartRow, ColNdx] = ColumnHeadings[i];
                        ColNdx++;
                    }
                }
                if (ElemTypeFixed == "")
                    F06Sheet.Cells[StartRow - 1, StartCol] = Request.ToUpper();
                else
                    F06Sheet.Cells[StartRow - 1, StartCol] = Request.ToUpper() + "-" + ElemTypeFixed;



                long EndCol = StartCol - 1 + F06Sheet.Range[F06Sheet.Cells[StartRow, StartCol], F06Sheet.Cells[StartRow, StartCol].End(Excel.XlDirection.xlToRight)].Count;
                long EndRow = StartRow - 1 + F06Sheet.Range[F06Sheet.Cells[StartRow, StartCol], F06Sheet.Cells[StartRow, StartCol].End(Excel.XlDirection.xlDown)].Count;

                Excel.Range Selection = F06Sheet.Range[F06Sheet.Cells[StartRow, StartCol], F06Sheet.Cells[EndRow, EndCol]];

                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                Selection.HorizontalAlignment = Excel.Constants.xlCenter;
                Selection.EntireColumn.AutoFit();
            }
        }

        public string[] GetColumnHeadings(ref string TextLine, ref string ElemType, ref string PreviousLine)
        {
            //string[] ColumnHeadings;

            if (TextLine.Contains("VON MISES"))
                TextLine = TextLine.Replace("VON MISES", "VON-MISES");
            if (TextLine.Contains("SHEAR XZ-MAT"))
                TextLine = TextLine.Replace("SHEAR XZ-MAT", "SHEAR_XZ-MAT");
            if (TextLine.Contains("SHEAR YZ-MAT"))
                TextLine = TextLine.Replace("SHEAR YZ-MAT", "SHEAR_YZ-MAT");
            if (TextLine.Contains("PLANE "))
                TextLine = TextLine.Replace("PLANE ", "PLANE");
            if (TextLine.Contains("POINT ID"))
                TextLine = TextLine.Replace("POINT ID", "POINT_ID");
            
            string[]  ColumnHeadings = TextLine.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            return ColumnHeadings;
        }

        private void ColumnHeadingsFinal(ref string[] ColumnHeadings, ref string Request, ref string SCDescription, ref bool Layer2, ref string ElemType) // As String()
        {
            // Dim Layer2 As Boolean
            //int ColCount;

            if (ElemType.Contains("QUAD") || ElemType.Contains("TRIA"))
            {
                ColumnHeadings[0] = "Element_ID";

                if (ColumnHeadings[1] == "ID")
                    ColumnHeadings[1] = "PLY_ID";
            }
            else if (ElemType.Contains("CBAR") && Request == "FORCES")
            {
                ColumnHeadings[0] = "Element_ID";
                ColumnHeadings[1] = "BM_EA_P1";
                ColumnHeadings[2] = "BM_EA_P2";
                ColumnHeadings[3] = "BM_EB_P1";
                ColumnHeadings[4] = "BM_EB_P2";
                ColumnHeadings[5] = "SHEAR_P1";
                ColumnHeadings[6] = "SHEAR_P2";
                ColumnHeadings[7] = "AXIAL_FORCE";
            }
            else if (ElemType.Contains("CROD"))
            {
                if (Request == "FORCES")
                {
                    ColumnHeadings = new string[3];
                    ColumnHeadings[0] = "Element_ID";
                    ColumnHeadings[1] = "AXIAL_FORCE";
                    ColumnHeadings[2] = "TORQUE";
                }
                else if (Request == "STRESSES")
                {
                    ColumnHeadings = new string[3];
                    ColumnHeadings[0] = "Element_ID";
                    ColumnHeadings[1] = "AXIAL_STRESS";
                    ColumnHeadings[2] = "TORSIONAL_STRESS";
                }
                else if (Request == "STRAINS")
                {
                    ColumnHeadings = new string[3];
                    ColumnHeadings[0] = "Element_ID";
                    ColumnHeadings[1] = "AXIAL_STRAIN";
                    ColumnHeadings[2] = "TORSIONAL_STRAIN";
                }
            }
            else if (ElemType.Contains("HEXA") || ElemType.Contains("PENTA") || ElemType.Contains("CTETRA"))
            {
                ColumnHeadings = new string[12];
                ColumnHeadings[0] = "Element_ID";
                ColumnHeadings[1] = "X";
                ColumnHeadings[2] = "Y";
                ColumnHeadings[3] = "Z";
                ColumnHeadings[4] = "XY";
                ColumnHeadings[5] = "YZ";
                ColumnHeadings[6] = "ZX";
                ColumnHeadings[7] = "PRINCIPAL_1";
                ColumnHeadings[8] = "PRINCIPAL_2";
                ColumnHeadings[9] = "PRINCIPAL_3";
                ColumnHeadings[10] = "MEAN_PRESSURE";
                ColumnHeadings[11] = "VON_MISES";
            }
            else if (ElemType.Contains("CBEAM"))
            {
                if (Request == "FORCES")
                {
                    ColumnHeadings = new string[19];
                    ColumnHeadings[0] = "Element_ID";
                    ColumnHeadings[1] = "GRID_1";
                    ColumnHeadings[2] = "STARTDIST_LENGTH_1";
                    ColumnHeadings[3] = "BM_P1_ND1";
                    ColumnHeadings[4] = "BM_P2_ND1";
                    ColumnHeadings[5] = "SHEAR_P1_ND1";
                    ColumnHeadings[6] = "SHEAR_P2_ND1";
                    ColumnHeadings[7] = "AXIAL_FORCE_ND1";
                    ColumnHeadings[8] = "TOTAL_TORQUE_ND1";
                    ColumnHeadings[9] = "WARPING_TORQUE_ND1";
                    ColumnHeadings[10] = "GRID_2";
                    ColumnHeadings[11] = "STARTDIST_LENGTH_2";
                    ColumnHeadings[12] = "BM_P1_ND2";
                    ColumnHeadings[13] = "BM_P2_ND2";
                    ColumnHeadings[14] = "SHEAR_P1_ND2";
                    ColumnHeadings[15] = "SHEAR_P2_ND2";
                    ColumnHeadings[16] = "AXIAL_FORCE_ND2";
                    ColumnHeadings[17] = "TOTAL_TORQUE_ND2";
                    ColumnHeadings[18] = "WARPING_TORQUE_ND2";
                }
                else if (Request == "STRAINS" || Request == "STRESSES")
                {
                    ColumnHeadings = new string[17];
                    ColumnHeadings[0] = "Element_ID";
                    ColumnHeadings[1] = "GRID_1";
                    ColumnHeadings[2] = "STARTDIST_LENGTH_1";
                    ColumnHeadings[3] = "SXC_1";
                    ColumnHeadings[4] = "SXD_1";
                    ColumnHeadings[5] = "SXE_1";
                    ColumnHeadings[6] = "SXF_1";
                    ColumnHeadings[7] = "S-MAX_1";
                    ColumnHeadings[8] = "SMIN_1";
                    ColumnHeadings[9] = "GRID_2";
                    ColumnHeadings[10] = "STARTDIST_LENGTH_2";
                    ColumnHeadings[11] = "SXC_2";
                    ColumnHeadings[12] = "SXD_2";
                    ColumnHeadings[13] = "SXE_2";
                    ColumnHeadings[14] = "SXF_2";
                    ColumnHeadings[15] = "S-MAX_2";
                    ColumnHeadings[16] = "SMIN_2";
                }
            }

            if (Layer2 == true)
            {
                int ColCount = ColumnHeadings.Length;
                if (SCDescription != "")
                {
                    string[] oldColumnHeadings = ColumnHeadings;
                    ColumnHeadings = new string[ColumnHeadings.Length + ColumnHeadings.Length - 1 + 1];
                    if (oldColumnHeadings != null)
                        Array.Copy(oldColumnHeadings, ColumnHeadings, Math.Min(ColumnHeadings.Length + ColumnHeadings.Length - 1 + 1, oldColumnHeadings.Length));
                }
                else
                {
                    string[] oldColumnHeadings = ColumnHeadings;
                    ColumnHeadings = new string[ColumnHeadings.Length + ColumnHeadings.Length - 2 + 1];
                    if (oldColumnHeadings != null)
                        Array.Copy(oldColumnHeadings, ColumnHeadings, Math.Min(ColumnHeadings.Length + ColumnHeadings.Length - 2 + 1, oldColumnHeadings.Length));
                }
                for (int i = ColCount; i < ColumnHeadings.Length; i++)
                    ColumnHeadings[i] = ColumnHeadings[i - ColCount + 2] + "2";
                for (int i = 2; i <= ColCount - 1; i++)
                    ColumnHeadings[i] = ColumnHeadings[i] + "1";
            }
            else if (SCDescription != "")
            {
                string[] oldColumnHeadings = ColumnHeadings;
                ColumnHeadings = new string[ColumnHeadings.Length + 1 + 1];
                if (oldColumnHeadings != null)
                    Array.Copy(oldColumnHeadings, ColumnHeadings, Math.Min(ColumnHeadings.Length + 1 + 1, oldColumnHeadings.Length));
            }
            else
            {
                string[] oldColumnHeadings = ColumnHeadings;
                ColumnHeadings = new string[ColumnHeadings.Length + 1];
                if (oldColumnHeadings != null)
                    Array.Copy(oldColumnHeadings, ColumnHeadings, Math.Min(ColumnHeadings.Length + 1, oldColumnHeadings.Length));
            }

            if (SCDescription != "")
            {
                ColumnHeadings[ColumnHeadings.Length - 2] = "SUBCASE";
                ColumnHeadings[ColumnHeadings.Length - 1] = "DESCRIPTION";
            }
            else
                ColumnHeadings[ColumnHeadings.Length - 1] = "SUBCASE";
        }



    }
}
