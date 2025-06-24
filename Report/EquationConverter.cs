using Microsoft.Office.Interop.Excel;
using StressUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Report
{
    class EquationConverter
    {

        static Regex RangeAddress = new Regex(@"([A-Z]{1,2})([0-9]+)");

        private string _TableID { get; set; }
        private string _ReportSheet { get; set; }
        private Dictionary<string, object> _DictTables { get; set; }
        private bool OptionUnits { get; set; }
        
        public string TableID
        {
            get
            {
                return _TableID;
            }
            set
            {
                _TableID = value;
            }
        }
        public string ReportSheet
        {
            get
            {
                return _ReportSheet;
            }
            set
            {
                _ReportSheet = value;
            }
        }

        public EquationConverter()
        {
            OptionUnits = StressUtilities.Properties.Settings.Default.OptionUnits;
        }

        public EquationConverter(string IDTable)
        {
            TableID = IDTable;
            OptionUnits = StressUtilities.Properties.Settings.Default.OptionUnits;

        }

        private void ReadTableData()
        {
            //The function is intended for reading only the table ID into the class so that the rest of the data can be read by the class method itself
        }

        public string DecodeFormula(ref string FormulaString, Dictionary<string, dynamic> DictTables, Dictionary<string, dynamic> FormulaTables, Dictionary<string, dynamic> SolSymb, string Request, ref string Item, ref long StartRow, string EquationType, ref string SheetReport, ref string bullet/*, string IDTable*/)
        {
            int Position, StartPos, i, offst = 0;
            bool Chkspl = false;
            string Prefix, Units, MathFormula = "", TempString = "";
            //string ColKey;
            string[] xlFunctSymb;
            long RowNdx = 1;
            string[] xlFunctReplace;

            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlapp = Globals.ThisAddIn.Application;

            xlFunctSymb = new[] { "SQRT", "SUM", "PI()", "ASINH", "ACOSH", "ATANH", "ACOTH", "ASECH", "ACSCH", "SINH", "COSH", "TANH", "COTH", "SECH", "CSCH", "ASIN", "ACOS", "ATAN", "ACOT", "ASEC", "ACSC", "SIN", "COS", "TAN", "COT", "SEC", "CSC", "LN", "LOG10", "<=", ">=", "<>" };
            xlFunctReplace = new[] { @"\sqrt", @"\sum", @"\pi", "sinh^-1 ", "cosh^-1 ", "tanh^-1 ", "coth^-1 ", "sech^-1 ", "cosech^-1 ", "sinh", "cosh", "tanh", "coth", "sech", "cosech", "sin^-1 ", "cos^-1 ", "tan^-1 ", "cot^-1 ", "sec^-1 ", "cosec^-1 ", "sin", "cos", "tan", "cot", "sec", "cosec", "ln", "log_10 ", @"\le ", @"\ge ", @"\ne " };

            //TableID = IDTable;  // Value assigned to the class property. Please do not remove this code. This can be moved to different method(function) in order to reduce the function complexity
            ReportSheet = SheetReport;

            if (FormulaTables.ContainsKey("REFERENCE"))
                offst = 4;
            else
                offst = 3;
            if (double.TryParse(Item, out double _))
                RowNdx = StartRow + offst + Convert.ToInt32(Item);
            // To check if excel functions are included in the equation. Refer ChkSupportFn for the list of functions supported

            if (FormulaString.StartsWith("="))
            {
                MathFormula = FormulaString.Replace("$", "");
                MathFormula = MathFormula.TrimStart('=');
            }
            else if (double.TryParse(FormulaString, out double ResultData))
                MathFormula = FormulaTables["SYMBOL"][Request] + "=" + Convert.ToString(Math.Round(ResultData, 2));  //double.Parse(DictTables[Item][Request]), 2));
            else
                MathFormula = FormulaTables["SYMBOL"][Request] + "=" + FormulaString;

            if (!double.TryParse(FormulaString, out double _))
            {
                MathFormula = ChkSupportFn(MathFormula, EquationType);
                if (!double.TryParse(MathFormula, out double _))
                {
                    if (double.TryParse(Item, out double _))
                        MathFormula = ReplaceRows(MathFormula, RowNdx);

                    for (i = 0; i < xlFunctSymb.Length; i++)
                        MathFormula = MathFormula.Replace(xlFunctSymb[i], xlFunctReplace[i]);

                    if (!MathFormula.StartsWith("="))
                    {
                        if (MathFormula.Contains(@"\matrix") && getDimension((object)xlapp.Application.Evaluate(FormulaString)) != 0 && !FormulaString.StartsWith("=MUNIT"))
                            MathFormula = MatrixResults(MathFormula, FormulaString, (FormulaTables["SYMBOL"][Request]), "FORMULA");
                        else
                            MathFormula = FormulaTables["SYMBOL"][Request] + "=" + MathFormula;
                    }
                    else
                        MathFormula = FormulaTables["SYMBOL"][Request] + MathFormula;



                    foreach (string ColKey in SolSymb.Keys)
                    {
                        StartPos = 0;
                        while (MathFormula.IndexOf(ColKey + "$", StartPos) != -1)
                        {
                            Position = MathFormula.IndexOf(ColKey + "$", StartPos);
                            if (MathFormula.IndexOf(@"\naryand", StartPos) != -1)
                                Chkspl = IsSeperator(MathFormula.Substring(Position - @"\naryand".Length, @"\naryand".Length));
                            if ((IsSeperator(MathFormula.Substring(Position - 1, 1)) || Chkspl == true) && IsSeperator(MathFormula.Substring(Position + ColKey.Length, 1)))
                            {
                                Prefix = MathFormula.Substring(Position - 1, 1);
                                if (EquationType == "FORMULA")
                                {
                                    MathFormula = ReplaceFirst(MathFormula, Prefix + ColKey + "$", Prefix + FormulaTables["SYMBOL"][SolSymb[ColKey]]);
                                    StartPos = Position;
                                }
                                else
                                {
                                    TempString = DictTables[Item][SolSymb[ColKey]];
                                    if (OptionUnits)
                                    {
                                        Units = DictTables["UNIT"][SolSymb[ColKey]];

                                        if (double.TryParse(TempString, out double ResultData))
                                        {
                                            MathFormula = ReplaceWithUnits(MathFormula, Prefix + ColKey + "$", Prefix + Convert.ToString(Math.Round(ResultData, 2)), ref Units,ref bullet);
                                        }
                                        else
                                        {
                                            MathFormula = ReplaceWithUnits(MathFormula, Prefix + ColKey + "$", Prefix + DictTables[Item][SolSymb[ColKey]], ref Units, ref bullet);
                                        }

                                    }
                                    else
                                    {
                                        if (double.TryParse(TempString, out double ResultData))
                                        {
                                            MathFormula = ReplaceFirst(MathFormula, Prefix + ColKey + "$", Prefix + Convert.ToString(Math.Round(ResultData, 2)));
                                        }
                                        else
                                        {
                                            MathFormula = ReplaceFirst(MathFormula, Prefix + ColKey + "$", Prefix + DictTables[Item][SolSymb[ColKey]]);
                                        }
                                    }

                                    StartPos = Position;
                                }
                            }
                            else
                                StartPos = Position + 1;
                        }
                    }
                    MathFormula = ResidueRanges(FormulaString, MathFormula, RowNdx, SolSymb, ref EquationType, ref bullet);
                    MathFormula = MathFormula.Replace("*", @"\times ");
                }
            }

            return MathFormula;
        }

        private string ReplaceFirst(string text, string search, string replace)
        {
            int pos = text.IndexOf(search);
            if (pos < 0)
            {
                return text;
            }
            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }

        private string ReplaceWithUnits(string text, string search, string replace, ref string Unit, ref string bullet)
        {
            string ReturnParam;
            int pos = text.IndexOf(search), posNext;

            if (pos < 0)
            {
                return text;
            }

            posNext = pos + search.Length;
            if (Unit != "-")
            {
                ReturnParam = text.Substring(0, pos);
               
                if (posNext < text.Length - 1)
                {
                    if (text.Substring(posNext, 1) == "^" || text.Substring(posNext, 1) == "/" || replace.Substring(0, 1) == "/")
                    {
                        return ReturnParam + replace.Substring(0, 1) + "(" + replace.Substring(1) + bullet + Unit + ")" + text.Substring(posNext);
                    }
                    else
                    {
                        return ReturnParam + replace + bullet + Unit + text.Substring(pos + search.Length);
                    }
                }
                else
                {
                    if (replace.Substring(0, 1) == "/")
                    {
                        return ReturnParam + replace.Substring(0, 1) + "(" + replace.Substring(1) + bullet + Unit + ")" + text.Substring(posNext);
                    }
                    else
                    {
                        return ReturnParam + replace + bullet + Unit + text.Substring(posNext);  //
                    }
                }
                

                //return ReturnParam;
            }
            else
            {
                return text.Substring(0, pos) + replace + text.Substring(posNext);
            }
        }

        private string ReplaceRows(string MathFormula, long RowNdx)
        {
            string[] Parameters;
            string[] Separators = new[] { "+", "-", "*", "/", "^", "=", "(", ")", ",", ":", ";", "&", "@", "|", ">", "<" };
            int i;
            string ReplaceText;
            long CellRow;
            MatchCollection matches;

            Parameters = MathFormula.Split(Separators, StringSplitOptions.RemoveEmptyEntries);

            for (i = 0; i < Parameters.Length; i++)
            {
                if (!double.TryParse(Parameters[i], out double _) && Parameters[i].Contains(RowNdx.ToString()))
                {
                    if (RangeAddress.IsMatch(Parameters[i]))
                    {
                        matches = RangeAddress.Matches(Parameters[i]);
                        CellRow = long.Parse(matches[0].Groups[2].Value);
                        if (CellRow == RowNdx)
                        {
                            ReplaceText = Parameters[i].Replace(RowNdx.ToString(), "$");
                            MathFormula = MathFormula.Replace(Parameters[i], ReplaceText);
                        }
                    }
                }
            }

            return MathFormula;
        }

        public int getDimension(object MatResult)
        {
            int MatrixRank;

            if (MatResult.GetType() == typeof(double) || MatResult.GetType() == typeof(string))
                return 0;

            Array CastedMat = MatResult as Array;
            MatrixRank = CastedMat.Rank;
            return MatrixRank;
        }

        public string MatrixResults(string MathFormula, string FormulaString, string FunctName, string Request)
        {
            long MatRowSize;
            long MatColSize;
            object EvalMatResult;
            string BuildEq = "", CellAddress, StartAddress = "", EndAddress = "", Symbol;
            int ArrayDim;
            int i, j;
            Excel.Range StartRng = null, EndRng;
            long[] MatSize;
            //Excel.Name TblName;
            string wrkShtName;
            Excel.Worksheet wrkSheet = null;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlapp = Globals.ThisAddIn.Application;


            EvalMatResult = xlapp.Application.Evaluate(FormulaString);
            ArrayDim = getDimension(EvalMatResult);
            Array CastedMat = EvalMatResult as Array;
            System.Collections.IList ResultCollection = CastedMat;
            if (ArrayDim == 0)
            {
                MathFormula = FunctName + "=" + MathFormula;
                BuildEq = MathFormula;
            }
            else
            {
                MatRowSize = ((Array)EvalMatResult).GetLength(0);  //Information.UBound(EvalMatResult, 1)
                if (ArrayDim == 2)
                    MatColSize = ((Array)EvalMatResult).GetLength(1); //Information.UBound(EvalMatResult, 2);
                else
                    MatColSize = 1;

                if (MatRowSize <= 12 && MatColSize <= 12)
                {
                    if (Request == "FORMULA")
                    {
                        foreach (Excel.Name TblName in wb.Names)
                        {
                            if (TblName.Name == TableID)
                            {
                                wrkShtName = wb.Names.Item(TblName.Name).RefersToRange.Parent.Name;
                                wrkSheet = wb.Worksheets[wrkShtName];
                                break;
                            }
                        }

                        wrkSheet.Activate();
                        int RowStart = wrkSheet.Range[TableID].Row;
                        int ColStart = wrkSheet.Range[TableID].Column;
                        int RowEnd = wrkSheet.Range[TableID].End[Excel.XlDirection.xlDown].Row;
                        int ColEnd = wrkSheet.Range[TableID].End[Excel.XlDirection.xlToRight].Column;
                        wrkSheet.Range[wrkSheet.Cells[RowStart, ColStart], wrkSheet.Cells[RowEnd, ColEnd]].Select();

                        Excel.Range rngData = wrkSheet.Range[wrkSheet.Cells[RowStart, ColStart], wrkSheet.Cells[RowEnd, ColEnd]];

                        string FormulaToFind = FormulaGenericToLocal(FormulaString);

                        StartRng = rngData.Find(FormulaToFind, LookIn: Excel.XlFindLookIn.xlFormulas, LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlNext);
                        if (StartRng != null)
                        {
                            EndRng = StartRng;
                            StartAddress = StartRng.Address;
                            do
                            {
                                EndAddress = EndRng.Address;
                                EndRng = rngData.Find(FormulaToFind, After: EndRng, LookIn: Excel.XlFindLookIn.xlFormulas, LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlNext);
                            }
                            while (EndRng != null && EndRng.Address != StartAddress);
                        }

                        MatSize = MatrixSize(StartAddress, EndAddress);

                        for (i = 0; i < MatRowSize; i++)
                        {
                            for (j = 0; j < MatColSize; j++)
                            {
                                if (MatColSize != MatSize[1])
                                {
                                    CellAddress = StartRng.Offset[j, i].Address;
                                    Symbol = GetEqSymbol(wrkSheet, CellAddress, Request);
                                    BuildEq += Symbol;
                                }
                                else
                                {
                                    CellAddress = StartRng.Offset[i, j].Address;
                                    Symbol = GetEqSymbol(wrkSheet, CellAddress, Request);
                                    BuildEq += Symbol;
                                }
                                if (j < MatColSize - 1)
                                    BuildEq += "&";
                            }
                            if (i < MatRowSize - 1)
                                BuildEq += "@";
                        }
                    }
                    else
                    {
                        foreach (object RowItem in ResultCollection/*CastedMat*/)
                        {
                            if (RowItem is System.Collections.IEnumerable)
                            {
                                foreach (object ColItem in (System.Collections.IEnumerable)RowItem)
                                {
                                    BuildEq += Math.Round(Convert.ToDouble(ColItem), 2).ToString();
                                    BuildEq += "&";
                                }
                                BuildEq = BuildEq.TrimEnd('&');
                            }
                            else
                            {
                                BuildEq += Math.Round(Convert.ToDouble(RowItem), 2).ToString();
                            }
                            BuildEq += "@";
                        }
                        BuildEq = BuildEq.TrimEnd('@');
                    }
                    /*for (i = 0; i < MatRowSize; i++)
                    {
                        for (j = 0; j < MatColSize; j++)
                        {
                            if (MatColSize == 1)

                                BuildEq += ResultCollection.[j]; //Convert.ToString(Math.Round(StartRng.Offset[j, i].Value2,2));//((System.Collections.IList)CastedMat)[i].ToString();

                            else
                                //BuildEq += ResultCollection[i].indexAt(j);//Convert.ToString(Math.Round(StartRng.Offset[i, j].Value2, 2)); //((List<string>)CastedMat)[i][j];  // Error needs to be fixed.
                            if (j < MatColSize-1)
                                BuildEq += "&";
                        }
                        if (i < MatRowSize-1)
                            BuildEq += "@";
                    }*/
                    BuildEq = @"[\matrix(" + BuildEq + ")] ";
                    BuildEq = BuildEq;
                }
                else
                    BuildEq = FunctName;
            }

            if (MathFormula != "")
                BuildEq += "=" + MathFormula;

            return BuildEq;
        }


        public dynamic CheckUniLink(string FormulaString)
        {
            bool UnaryChk = false;
            string tmpFormula = FormulaString;
            string Reference;
            Excel.Worksheet wrkSheet;
            string CellAddress;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            dynamic UniLink;

            if (tmpFormula.StartsWith("="))
                tmpFormula = tmpFormula.TrimStart('=');

            if (tmpFormula.StartsWith("-"))
            {
                UnaryChk = true;
                tmpFormula = tmpFormula.TrimStart('-');
            }

            tmpFormula = tmpFormula.Replace("$", "");

            if (IsEquation(tmpFormula) || double.TryParse(tmpFormula, out double _))
                UniLink = FormulaString;
            else if (!UnaryChk && !FormulaString.StartsWith("="))
                UniLink = FormulaString;
            else
            {
                if (tmpFormula.Contains("!"))
                {
                    Reference = tmpFormula.Split('!')[0];
                    CellAddress = tmpFormula.Split('!')[1];
                }
                else
                {
                    Reference = wb.ActiveSheet.Name;
                    CellAddress = tmpFormula;
                }
                wrkSheet = wb.Worksheets[Reference];
                try
                {
                    UniLink = wrkSheet.Range[CellAddress].Value;
                }
                catch (Exception ex)
                {
                    UniLink = tmpFormula;
                }
                if (UniLink.ToString().IndexOf(xlApp.Application.DecimalSeparator) != -1)
                    UniLink = Math.Round(UniLink, 2);
                if (UnaryChk == true)
                    UniLink = -1 * UniLink;
            }

            return UniLink;
        }


        private string FormulaLocalToGeneric(string iFormula)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            wb.Names.Add("tempFormula", RefersToLocal: iFormula);
            string GenericFormula = wb.Names.Item("tempFormula").RefersTo;
            wb.Names.Item("tempFormula").Delete();
            return GenericFormula;
        }
        private string FormulaGenericToLocal(string iFormula)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            wb.Worksheets[ReportSheet].Range["AZ1"].Formula = iFormula;
            string FormulaLocal = wb.Worksheets[ReportSheet].Range["AZ1"].Formulalocal;
            wb.Worksheets[ReportSheet].Range["AZ1"].Value = "";
            return FormulaLocal;
        }




        private string ResidueRanges(string FormulaString, string MathFormula, long RowNdx, Dictionary<string, object> SolSymb, ref string Request, ref string bullet)
        {
            string TruncEqu, Unit;
            int i;
            string tmpFormula; tmpFormula = FormulaString;
            string[] ResidueCells;
            int count; count = 0;
            string[] ListRanges = new string[51];
            string ReplaceVal;
            //string ColKey;
            string Reference;
            string CellAddress;
            Excel.Worksheet wrkSheet;
            Excel.Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            string[] xlFunctSymb = new[] { "SQRT", "SUM", "PI()", "PRODUCT", "ASINH", "ACOSH", "ATANH", "ACOTH", "ASECH", "ACSCH", "SINH", "COSH", "TANH", "COTH", "SECH", "CSCH", "ASIN", "ACOS", "ATAN", "ACOT", "ASEC", "ACSC", "SIN", "COS", "TAN", "COT", "SEC", "CSC", "LN", "LOG10", "<=", ">=", "<>" };
            string[] xlFunctionEval = new[] { "MINVERSE", "MDETERM", "MMULT", "MUNIT", "VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "ADDRESS", "DGET", "SIGN", "COUNTA", "COUNTBLANK", "COUNTIFS", "COUNTIF", "COUNT", "FORECAST.LINEAR", "FORECAST", "ROW", "COLUMN" };
            string[] xlFunct = new[] { "SUMSQ", "SUM", "ABS", "DEGREES", "RADIANS", "IF", "AVERAGE", "ROUNDUP", "ROUNDDOWN", "MROUND", "ROUND", "TRUNC", "MAX", "MIN", "EXP", "LOG", "POWER", "SIGN", "AND", "OR" };

            string[] Separators = new[] { "+", "-", "*", "/", "^", "=", "(", ")", ":", ";", "&", "@", "|", ">", "<" };

            for (i = 0; i < xlFunctionEval.Length; i++)
            {
                if (tmpFormula.Contains(xlFunctionEval[i]))
                {
                    TruncEqu = TrimFunction(tmpFormula, System.Convert.ToString(xlFunctionEval[i]));
                    tmpFormula = tmpFormula.Replace(TruncEqu, " ");
                }
            }

            for (i = 0; i < xlFunctSymb.Length; i++)
            {
                if (tmpFormula.Contains(xlFunctSymb[i]))
                    tmpFormula = tmpFormula.Replace(xlFunctSymb[i], ",");
            }

            for (i = 0; i < xlFunct.Length; i++)
            {
                if (tmpFormula.Contains(xlFunct[i]))
                    tmpFormula = tmpFormula.Replace(xlFunct[i], ",");
            }

            for (i = 0; i < Separators.Length; i++)
            {
                if (tmpFormula.Contains(Separators[i]))
                    tmpFormula = tmpFormula.Replace(Separators[i], ",");
            }

            foreach (string ColKey in SolSymb.Keys)
            {
                if (tmpFormula.Contains("," + ColKey + RowNdx))
                    tmpFormula = tmpFormula.Replace("," + ColKey + RowNdx, ",");
            }


            while (tmpFormula.Contains(",,"))
                tmpFormula = tmpFormula.Replace(",,", ",");

            tmpFormula = tmpFormula.Replace("$", "");
            ResidueCells = tmpFormula.Split(new char[] {','},StringSplitOptions.RemoveEmptyEntries);

            for (i = 0; i < ResidueCells.Length; i++)
            {
                if (!double.TryParse(ResidueCells[i], out double _) && !string.IsNullOrEmpty(ResidueCells[i]))
                {
                    if (double.TryParse(ResidueCells[i].Substring(ResidueCells[i].Length - 1, 1), out double _))
                    {
                        ListRanges[count] = ResidueCells[i];
                        count++;
                    }
                }
            }

            for (i = 0; i <= count - 1; i++)
            {
                if (tmpFormula.Contains(ListRanges[i]))
                {
                    if (ListRanges[i].Contains("!"))
                    {
                        Reference = ListRanges[i].Split('!')[0];
                        CellAddress = ListRanges[i].Split('!')[1];
                    }
                    else
                    {
                        Reference = ActiveSheet.Name;
                        CellAddress = ListRanges[i];
                    }
                    wrkSheet = wb.Worksheets[Reference];
                    try
                    {
                        if (Request == "FORMULA")
                            ReplaceVal = GetEqSymbol(wrkSheet, CellAddress, Request);
                        else
                        {
                            if (OptionUnits)
                            {
                                Unit = GetEqSymbol(wrkSheet, CellAddress, Request);
                                if (Unit != "-")
                                {
                                    Unit = Unit.Replace("[", "").Replace("]", "");
                                    if (Unit.ToUpper() == "MPA")
                                    {
                                        Unit = "MPa";
                                    }
                                    ReplaceVal = $"({wrkSheet.Range[CellAddress].Value:0.00}{bullet}{Unit})";
                                }
                                else
                                    ReplaceVal = $"{wrkSheet.Range[CellAddress].Value:0.00}";
                            }
                            else
                                ReplaceVal = $"{wrkSheet.Range[CellAddress].Value:0.00}";
                        }
                    }

                    catch (Exception ex)
                    {
                        ReplaceVal = CellAddress;
                    }

                    try
                    {
                        if (wrkSheet.Range[ListRanges[i]].Row == RowNdx)
                        {
                            ListRanges[i] = ListRanges[i].Replace("$", "");
                            ListRanges[i] = ListRanges[i].Replace(RowNdx.ToString(), "$");
                        }
                    }
                    catch (Exception ex)
                    {
                    }

                    if (ReplaceVal.ToString().IndexOf(xlApp.Application.DecimalSeparator) != -1 && double.TryParse(ReplaceVal, out double _))
                        ReplaceVal = Math.Round(double.Parse(ReplaceVal), 2).ToString();


                    MathFormula = MathFormula.Replace(ListRanges[i], ReplaceVal);
                    MathFormula = CheckCustomNames(MathFormula, Request);
                }
            }

            return MathFormula;
        }


        private string CheckCustomNames(string MathFormula, string Request)
        {
            //Name UDFName;
            string wrkSheet, Unit;
            Excel.Worksheet CalcSheet;
            string CellAddress;
            string ReplaceVal;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            foreach (Name UDFName in wb.Names)
            {
                if (MathFormula.Contains(UDFName.Name))
                {
                    CellAddress = UDFName.Name;
                    wrkSheet = wb.Names.Item(UDFName.Name).RefersToRange.Parent.Name;
                    CalcSheet = wb.Worksheets[wrkSheet];
                    if (Request == "FORMULA")
                        ReplaceVal = GetEqSymbol(CalcSheet, CellAddress, Request);
                    else
                    {
                        if (OptionUnits)
                        {
                            Unit = GetEqSymbol(CalcSheet, CellAddress, Request);
                            if (Unit != "-")
                            {
                                Unit = Unit.Replace("[", "").Replace("]", "");
                                if (Unit.ToUpper() == "MPA")
                                {
                                    Unit = "MPa";
                                }
                                ReplaceVal = $"({CalcSheet.Range[CellAddress].Value:0.00}\\bullet{Unit})";
                            }
                            else
                            ReplaceVal = CalcSheet.Range[CellAddress].Value;
                        }
                        else
                        {
                            ReplaceVal = CalcSheet.Range[CellAddress].Value;
                        }
                    }
                    MathFormula = MathFormula.Replace(CellAddress, ReplaceVal);
                }
            }
            return MathFormula;
        }

        private string GetEqSymbol(Excel.Worksheet wrkSheet, string CellAddress, string Request)
        {
            string AnchorAddress, SymbCellAddr, EqSymbol;
            long RowNdx, ColNdx, tmpColNdx, offsetCol;


            AnchorAddress = wrkSheet.Range[CellAddress].End[Excel.XlDirection.xlUp].End[Excel.XlDirection.xlToLeft].Address;
            if (wrkSheet.Range[AnchorAddress].Text == "TBLDESCR")
            {
                SymbCellAddr = wrkSheet.Range[CellAddress].End[Excel.XlDirection.xlUp].Address;
                RowNdx = wrkSheet.Range[SymbCellAddr].Row;
                ColNdx = wrkSheet.Range[SymbCellAddr].Column;
                if (Request == "FORMULA")
                    EqSymbol = wrkSheet.Cells[RowNdx + 2, ColNdx].Value2;
                else
                {
                    EqSymbol = wrkSheet.Cells[RowNdx + 3, ColNdx].Value2;
                }
            }
            else
            {
                RowNdx = wrkSheet.Range[CellAddress].Row;
                ColNdx = wrkSheet.Range[CellAddress].Column;
                //tmpColNdx = ColNdx;
                tmpColNdx = wrkSheet.Range[CellAddress].End[XlDirection.xlToLeft].Column+2;
                /*if (tmpColNdx != 0)
                {
                    while (double.TryParse(wrkSheet.Cells[RowNdx, tmpColNdx].Text, out double _) && !string.IsNullOrEmpty(wrkSheet.Cells[RowNdx, tmpColNdx].formula))
                        tmpColNdx -= 1;
                }*/

                if (string.IsNullOrEmpty(wrkSheet.Cells[RowNdx, tmpColNdx].formula))
                    EqSymbol = Convert.ToString(Math.Round(wrkSheet.Range[CellAddress].Value, 3));
                else
                {
                    if (Request == "FORMULA")
                    {
                        //tmpColNdx = wrkSheet.Range[CellAddress].End[XlDirection.xlToLeft].Column + 2;
                        EqSymbol = wrkSheet.Cells[RowNdx, tmpColNdx].Text;
                    }
                    else
                    {
                        tmpColNdx = wrkSheet.Range[CellAddress].End[XlDirection.xlToRight].Column;
                        EqSymbol = wrkSheet.Cells[RowNdx, tmpColNdx].Text;
                    }
                }
            }

            return EqSymbol;
        }



        private bool IsSeperator(string CellAddress)
        {
            int i; i = 0;
            string[] Separators = new[] { "$", "+", "-", "*", "/", "^", "=", "(", ")", ",", ":", ";", "&", "@", "|", ">", "<", @"\naryand", " " };
            bool Status = false;

            while (Status == false && i < Separators.Length)
            {
                if (CellAddress == Separators[i])
                    Status = true;
                i++;
            }
            return Status;
        }

        private bool IsEquation(string Equation)
        {
            string[] Separators = new[] { "+", "-", "*", "/", "^", "=", "(", ")", ",", ":", ";", "&", "@", "|", ">", "<" };
            int i; i = 0;
            bool Status = false;

            while (Status == false && i < Separators.Length)
            {
                if (Equation.Contains(Separators[i]))
                    Status = true;
                i++;
            }
            return Status;
        }

        private string ChkSupportFn(string FormulaString, string Request)
        {
            int i;
            string ReplaceEquation; ReplaceEquation = FormulaString;
            string xlFunction;
            string listSeparator;
            string[] xlFunct = new[] { "MINVERSE", "MDETERM", "MMULT", "TRANSPOSE", "MUNIT", "SUMSQ", "SUM", "ABS", "DEGREES", "RADIANS", "IF", "AVERAGE", "ROUNDUP", "ROUNDDOWN", "MROUND", "ROUND", "TRUNC", "MAX", "MIN", "EXP", "LOG", "POWER", "SIGN", "VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "ADDRESS", "DGET", "COUNTA", "COUNTBLANK", "COUNTIFS", "COUNTIF", "COUNT", "FORECAST.LINEAR", "FORECAST", "ROW", "COLUMN", "PRODUCT" };

            listSeparator = ",";

            for (i = 0; i < xlFunct.Length; i++)
            {
                if (FormulaString.Contains(xlFunct[i]))
                {
                    xlFunction = xlFunct[i];
                    switch (xlFunction)
                    {
                        case "IF":
                            {
                                FormulaString = convertif(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "SUMSQ":
                            {
                                FormulaString = convertSumSq(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "ABS":
                            {
                                FormulaString = convertAbs(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "DEGREES":
                            {
                                FormulaString = DegRad(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "RADIANS":
                            {
                                FormulaString = DegRad(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "EXP":
                            {
                                FormulaString = convertEXP(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "POWER":
                            {
                                FormulaString = convertPOWER(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "AVERAGE":
                        case "SUM":
                        case "MAX":
                        case "MIN":
                        case "PRODUCT":
                            {
                                FormulaString = convertList(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "LOG":
                            {
                                FormulaString = convertLOG(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "ROUNDUP":
                        case "ROUNDDOWN":
                        case "MROUND":
                        case "ROUND":
                        case "TRUNC":
                            {
                                FormulaString = convertRound(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "VLOOKUP":
                        case "HLOOKUP":
                        case "XLOOKUP":
                        case "INDEX":
                        case "MATCH":
                        case "ADDRESS":
                        case "DGET":
                        case "SIGN":
                        case "COUNTA":
                        case "COUNTBLANK":
                        case "COUNTIFS":
                        case "COUNTIF":
                        case "COUNT":
                        case "FORECAST":
                        case "FORECAST.LINEAR":
                        case "ROW":
                        case "COLUMN":
                            {
                                FormulaString = EvalEquations(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "MUNIT":
                            {
                                FormulaString = ConvertMunit(FormulaString, xlFunction, listSeparator);
                                break;
                            }

                        case "MINVERSE":
                        case "MDETERM":
                        case "TRANSPOSE":
                            {
                                FormulaString = ConvertMInverseDet(FormulaString, xlFunction, listSeparator, Request);
                                break;
                            }

                        case "MMULT":
                            {
                                FormulaString = ConvertMMult(FormulaString, xlFunction, listSeparator, Request);
                                break;
                            }
                    }
                }
            }
            if (Request == "FORMULA")
            {
                if (FormulaString.Contains(@"\degree "))
                    FormulaString = FormulaString.Replace(@"\degree", "");
            }

            return FormulaString;
        }


        private string convertList(string Equation, string xlFunction, string listSeparator)
        {
            int count;
            int i;
            int j;
            string tempequation;
            string TruncEqu;
            string Mathstr; Mathstr = "";
            string[] splitfn;
            string[] splitRng;
            string BuildEq; BuildEq = "";
            int countSep;
            int CountRange;
            string StartAddress;
            string EndAddress;
            long StartRow;
            long StartColumn;
            long[] MatSize = { };
            string Addr;
            string[] splitaddr;
            long RowNdx;
            long ColNdx;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;

            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));
                countSep = Mathstr.Length - Mathstr.Replace(listSeparator, "").Length;
                CountRange = Mathstr.Length - Mathstr.Replace(":", "").Length;
                splitfn = Mathstr.Split(listSeparator.ToCharArray(), StringSplitOptions.None);

                if (CountRange != 0)
                {
                    for (j = 0; j < splitfn.Length; j++)
                    {
                        if (splitfn[j].Contains(":"))
                        {
                            splitRng = splitfn[j].Split(':');
                            StartAddress = splitRng[0];
                            EndAddress = splitRng[1];
                            StartRow = wb.ActiveSheet.Range[StartAddress].Row;
                            StartColumn = wb.ActiveSheet.Range[StartAddress].Column;

                            MatSize = MatrixSize(StartAddress, EndAddress);
                            if (MatSize[0] == 1)
                            {
                                for (ColNdx = StartColumn; ColNdx <= StartColumn + MatSize[1] - 1; ColNdx++)
                                {
                                    Addr = wb.ActiveSheet.Cells[StartRow, ColNdx].Address.Replace("$", "");
                                    BuildEq = BuildEq + Addr + listSeparator;
                                }
                            }
                            else if (MatSize[1] == 1)
                            {
                                splitaddr = wb.ActiveSheet.Cells[StartRow, StartColumn].Address.Split("$");
                                BuildEq = splitaddr[1] + "$_i ";
                            }
                            else
                                for (RowNdx = StartRow; RowNdx <= StartRow + MatSize[0] - 1; RowNdx++)
                                {
                                    for (ColNdx = StartColumn; ColNdx <= StartColumn + MatSize[1] - 1; ColNdx++)
                                    {
                                        splitaddr = wb.ActiveSheet.Cells[RowNdx, ColNdx].Address.Split("$");
                                        BuildEq = BuildEq + splitaddr[1] + "$_" + (RowNdx - StartRow + 1).ToString() + "+";   // Technic to cell reference
                                    }
                                }
                        }
                        else
                            BuildEq = BuildEq + splitfn[j] + listSeparator;
                    }
                }
                else
                    for (j = 0; j < splitfn.Length; j++)
                        BuildEq = BuildEq + splitfn[j] + listSeparator;

                BuildEq = BuildEq.Substring(0, BuildEq.Length - 1);

                if (xlFunction == "AVERAGE")
                {
                    if (MatSize[0] == 1)
                    {
                        BuildEq = BuildEq.Replace(listSeparator, "+");
                        BuildEq = "(" + BuildEq + ")" + "/" + (MatSize[0] * MatSize[1]);
                    }
                    else if (MatSize[1] == 1)
                        BuildEq = @"\sum_(i=1)^" + MatSize[0] + @"\naryand" + BuildEq + "/" + (MatSize[0] * MatSize[1]);
                    else
                        BuildEq = @"\sum_(1<=i<=" + MatSize[0] + "@1<=j<=" + MatSize[1] + ")" + @"\naryand" + BuildEq + "/" + (MatSize[0] * MatSize[1]);
                }
                else if (xlFunction == "SUM")
                {
                    if (MatSize[0] == 1)
                        BuildEq = BuildEq.Replace(listSeparator, "+");
                    else if (MatSize[1] == 1)
                        BuildEq = @"\sum_(i=1)^" + MatSize[0] + @"\naryand" + BuildEq;
                    else
                        BuildEq = @"\sum_(1<=i<=" + MatSize[0] + "@1<=j<=" + MatSize[1] + ")" + @"\naryand" + BuildEq;
                }
                else if (xlFunction == "PRODUCT")
                {
                    if (MatSize[0] == 1)
                    {
                        BuildEq = BuildEq.Replace(listSeparator, "*");
                        BuildEq = "(" + BuildEq + ")";
                    }
                    else if (MatSize[1] == 1)
                        BuildEq = @"\prod" + MatSize[0] + @"\naryand" + BuildEq;
                    else
                        BuildEq = @"\prod_(1<=i<=" + MatSize[0] + "@1<=j<=" + MatSize[1] + ")" + @"\naryand" + BuildEq;
                }
                else
                    BuildEq = xlFunction.ToLower() + "(" + BuildEq + ")";

                tempequation = tempequation.Replace(TruncEqu, BuildEq);
            }
            return tempequation;
        }


        private string convertRound(string Equation, string xlFunction, string listSeparator)
        {
            int count;
            int i; // , j As Integer
            string tempequation;
            string TruncEqu;
            string Mathstr;
            string[] splitfunction;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";
            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));
                if (!Mathstr.Contains(listSeparator))
                    Mathstr = Mathstr + listSeparator + "0";
                splitfunction = Mathstr.Split(new[] { listSeparator }, StringSplitOptions.None);

                if (IsEquation(splitfunction[0]) && !splitfunction[0].Contains("(") && !splitfunction[0].EndsWith(")"))
                    splitfunction[0] = string.Concat("(", splitfunction[0], ")");

                tempequation = tempequation.Replace(TruncEqu, splitfunction[0]);
            }
            return tempequation;
        }

        private string convertPOWER(string Equation, string xlFunction, string listSeparator)
        {
            int count;
            int i;
            string tempequation;
            string TruncEqu;
            string Mathstr;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";
            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));

                Mathstr = Mathstr.Replace(listSeparator, "^");
                tempequation = tempequation.Replace(TruncEqu, Mathstr);
            }
            return tempequation;
        }

        private string convertLOG(string Equation, string xlFunction, string listSeparator)
        {
            int count;
            int i;
            string tempequation;
            string TruncEqu;
            string Mathstr;
            string[] splitfunction;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";
            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));
                if (!Mathstr.Contains("LOG10"))
                {
                    splitfunction = Mathstr.Split(new[] { listSeparator }, StringSplitOptions.None);
                    Mathstr = "log_" + splitfunction[1] + " (" + splitfunction[0] + ")";
                    tempequation = tempequation.Replace(TruncEqu, Mathstr);
                }
            }
            return tempequation;
        }

        private string convertEXP(string Equation, string xlFunction, string listSeparator)
        {
            int count;
            int i;
            string tempequation;
            string TruncEqu;
            string Mathstr;


            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";
            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));
                Mathstr = "e^" + Mathstr;

                tempequation = tempequation.Replace(TruncEqu, Mathstr);
            }
            return tempequation;
        }

        private string convertAbs(string Equation, string xlFunction, string listSeparator)
        {
            int count;
            int i;
            string tempequation;
            string TruncEqu;
            string Mathstr;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";
            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                Mathstr = "|" + TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2)) + "|";
                tempequation = tempequation.Replace(TruncEqu, Mathstr);
            }
            return tempequation;
        }

        private string convertSumSq(string Equation, string xlFunction, string listSeparator)
        {
            int count;
            int i;
            int j;
            string[] SumSqSplit;
            string tempequation;
            string TruncEqu;
            string Mathstr;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";
            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                SumSqSplit = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2)).Split(new[] { listSeparator }, StringSplitOptions.None);
                for (j = 0; j < SumSqSplit.Length; j++)
                    Mathstr = Mathstr + SumSqSplit[j] + "^2+";
                Mathstr = Mathstr.Substring(0, Mathstr.Length - 1);
                tempequation = tempequation.Replace(TruncEqu, Mathstr);
            }
            return tempequation;
        }

        private string convertif(string Equation, string xlFunction, string listSeparator)
        {
            int ifcount;
            int ncases;
            int i;
            int j;
            string tempequation;
            string[] IfSplit;
            string Mathstr;
            string TruncEqu;
            string[] xlsubFunct = new[] { "AND", "OR" };
            string Condition = "", TrimPart = "";
            int ofst; ofst = 2;
            string equality;
            string rightParam;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            ifcount = CountOccurance(Equation, xlFunction);
            ncases = ifcount + 1;
            tempequation = Equation;
            Mathstr = "";

            for (i = 1; i <= ifcount; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);

                for (j = 0; j < xlsubFunct.Length; j++)
                {
                    if (TruncEqu.Contains(xlsubFunct[j]))
                    {
                        switch (xlsubFunct[j])
                        {
                            case "AND":
                                {
                                    Condition = ConvertAND(TruncEqu, System.Convert.ToString(xlsubFunct[j]), listSeparator);
                                    TrimPart = TrimFunction(tempequation, System.Convert.ToString(xlsubFunct[j]));
                                    break;
                                }

                            case "OR":
                                {
                                    Condition = ConvertOR(TruncEqu, System.Convert.ToString(xlsubFunct[j]), listSeparator);
                                    TrimPart = TrimFunction(tempequation, System.Convert.ToString(xlsubFunct[j]));
                                    break;
                                }
                        }
                        tempequation = tempequation.Replace(TrimPart, Condition);
                        TruncEqu = TruncEqu.Replace(TrimPart, Condition);
                    }
                }

                IfSplit = TruncEqu.Substring(xlFunction.Length + ofst - 1, TruncEqu.Length - (xlFunction.Length + ofst)).Split(new[] { listSeparator }, StringSplitOptions.None);

                if (IfSplit[2] != "DUMMY")
                {
                    if (IfSplit[2].Contains("\"\""))
                    {
                        equality = EqualityCheck(IfSplit[0]);
                        rightParam = IfSplit[0].Split(new[] { equality }, StringSplitOptions.None)[1];
                        if (rightParam.Contains("\"\""))
                            tempequation = tempequation.Replace(TruncEqu, IfSplit[1]);
                        else
                        {
                            Mathstr = IfSplit[1] + ",  &" + IfSplit[0];
                            tempequation = tempequation.Replace(TruncEqu, "DUMMY");
                        }
                    }
                    else if (IfSplit[1].Contains("\"\""))
                    {
                        equality = EqualityCheck(IfSplit[0]);
                        rightParam = IfSplit[0].Split(new[] { equality }, StringSplitOptions.None)[1];
                        if (rightParam.Contains("\"\""))
                            tempequation = tempequation.Replace(TruncEqu, IfSplit[2]);
                        else
                        {
                            Mathstr = IfSplit[2] + ",  &" + IfSplit[0];
                            tempequation = tempequation.Replace(TruncEqu, "DUMMY");
                        }
                    }
                    else
                    {
                        Mathstr = IfSplit[2] + ",  &Otherwise" + Mathstr;
                        Mathstr = IfSplit[1] + ",  &" + IfSplit[0] + " @" + Mathstr;
                        tempequation = tempequation.Replace(TruncEqu, "DUMMY");
                    }
                }
                else
                {
                    Mathstr = IfSplit[1] + ",  &" + IfSplit[0] + " @" + Mathstr;
                    tempequation = tempequation.Replace(TruncEqu, "DUMMY");
                }
            }
            if (tempequation.Contains("DUMMY"))
                return @"{\matrix(" + Mathstr + @")\right  ";
            else
                return tempequation;

        }

        private string ConvertAND(string Equation, string xlFunction, string listSeparator)
        {
            int countAND;
            int ncases;
            int i;
            string tempequation;
            string[] ANDSplit;
            string Mathstr;
            string TruncEqu;
            string[] leftParam = new string[2];
            string[] rightParam = new string[2];
            string[] equality = new string[2];

            countAND = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";


            for (i = 1; i <= countAND; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                ANDSplit = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2)).Split(new[] { listSeparator }, StringSplitOptions.None);
                ncases = ANDSplit.Length;

                if (ANDSplit.Length - 1 == 1)
                {
                    equality[0] = EqualityCheck(ANDSplit[0]);
                    leftParam[0] = ANDSplit[0].Split(new[] { equality[0] }, StringSplitOptions.None)[0];
                    rightParam[0] = ANDSplit[0].Split(new[] { equality[0] }, StringSplitOptions.None)[1];
                    equality[1] = EqualityCheck(ANDSplit[1]);
                    leftParam[1] = ANDSplit[1].Split(new[] { equality[1] }, StringSplitOptions.None)[0];
                    rightParam[1] = ANDSplit[1].Split(new[] { equality[1] }, StringSplitOptions.None)[1];

                    if (leftParam[0] == leftParam[1])
                    {
                        if (equality[0].StartsWith(">"))
                            Mathstr = rightParam[0] + equality[0].Replace(">", "<") + leftParam[0] + equality[1] + rightParam[1];
                        else if (equality[0].StartsWith("<") && equality[0].StartsWith("<>"))
                            Mathstr = rightParam[0] + equality[0].Replace("<", ">") + leftParam[0] + equality[1] + rightParam[1];
                        else
                            Mathstr = rightParam[0] + equality[0] + leftParam[0] + equality[1] + rightParam[1];
                    }
                    else if (rightParam[0] == leftParam[1])
                        Mathstr = leftParam[0] + equality[0] + rightParam[0] + equality[1] + rightParam[1];
                    else
                    {
                        Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));
                        Mathstr = Mathstr.Replace(",", " and ");
                    }
                }
                else
                {
                    Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));
                    Mathstr = Mathstr.Replace(",", " and ");
                }
            }
            return Mathstr;
        }

        private string EqualityCheck(string Equation)
        {
            int i;
            string[] equality = new[] { ">=", "<=", "<>", "=", ">", "<" };

            string ReturnValue = "NONE";


            for (i = 0; i < equality.Length; i++)
            {
                if (Equation.Contains(equality[i]))
                {
                    ReturnValue = equality[i];
                    break;
                }
            }
            return ReturnValue;
        }


        private string ConvertOR(string Equation, string xlFunction, string listSeparator)
        {
            int countOR;
            int i;
            string tempequation;
            string Mathstr;
            string TruncEqu;


            countOR = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";

            for (i = 1; i <= countOR; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));
                Mathstr = Mathstr.Replace(listSeparator, " or ");
                tempequation = tempequation.Replace(TruncEqu, Mathstr);
            }
            return Mathstr;
        }



        private string DegRad(string Equation, string xlFunction, string listSeparator)
        {
            string Mathstr;
            int count;
            string tempequation;
            string TruncEqu;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";
            while (count != 0)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                if (xlFunction == "DEGREES")
                    Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));
                else if (xlFunction == "RADIANS")
                    Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2)) + @"\degree ";
                tempequation = tempequation.Replace(TruncEqu, Mathstr);
                count = CountOccurance(tempequation, xlFunction);
            }
            return tempequation;
        }

        private int CountOccurance(string Equation, string xlFunction)
        {
            return (Equation.Length - Equation.Replace(xlFunction, "").Length) / xlFunction.Length;
        }

        private string TrimFunction(string Equation, string xlFunction)
        {
            int ParOpen;
            int ParClose; ParClose = 0;
            int StartPos;
            int EndPos;
            string TruncEqu;

            StartPos = Equation.LastIndexOf(xlFunction);
            EndPos = Equation.IndexOf(")", StartPos);
            TruncEqu = Equation.Substring(StartPos, EndPos - StartPos + 1);
            ParOpen = TruncEqu.Length - TruncEqu.Replace("(", "").Length;
            ParClose = TruncEqu.Length - TruncEqu.Replace(")", "").Length;
            while (ParOpen != ParClose)
            {
                EndPos = Equation.IndexOf(")", EndPos + 1);
                TruncEqu = Equation.Substring(StartPos, EndPos - StartPos + 1);
                ParOpen = TruncEqu.Length - TruncEqu.Replace("(", "").Length;
                ParClose = TruncEqu.Length - TruncEqu.Replace(")", "").Length;
            }
            return TruncEqu;
        }

        private string EvalEquations(string Equation, string xlFunction, string listSeparator)
        {
            string Mathstr;
            int count;
            string tempequation;
            int i;
            string TruncEqu;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";
            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                if (TruncEqu.StartsWith("="))
                {
                }
                Mathstr = Math.Round(xlApp.Application.Evaluate(TruncEqu), 2);
                tempequation = tempequation.Replace(TruncEqu, Mathstr);
            }
            return tempequation;
        }


        private string ConvertMunit(string Equation, string xlFunction, string listSeparator)
        {
            string Mathstr;
            int count;
            string tempequation;
            int i;
            int j;
            int k;
            int MatSize;
            string TruncEqu;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";
            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                MatSize = System.Convert.ToInt32(TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2)));
                if (MatSize <= 12)
                {
                    for (j = 0; j <= MatSize - 1; j++)
                    {
                        for (k = 0; k <= MatSize - 1; k++)
                        {
                            if (j == k)
                                Mathstr += 1;
                            else
                                Mathstr += 0;
                            if (k < MatSize - 1)
                                Mathstr += "&";
                        }
                        if (j < MatSize - 1)
                            Mathstr += "@";
                    }

                    Mathstr = @"[\matrix(" + Mathstr + ")] ";
                }
                else
                    Mathstr = "I_" + MatSize;
                tempequation = tempequation.Replace(TruncEqu, Mathstr);
            }
            return tempequation;
        }

        private string ConvertMInverseDet(string Equation, string xlFunction, string listSeparator, string Request)
        {
            string Mathstr;
            int count;
            string tempequation;
            int i;
            int j;
            int k;
            int RowSize;
            int ColSize;
            string TruncEqu;
            string BuildEq = "";
            dynamic tmpMatrix;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";

            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));

                if (Mathstr.Contains("["))
                    BuildEq = Mathstr.Replace(" ", "");
                else if (Mathstr.Contains("("))
                {
                    tmpMatrix = xlApp.Evaluate(Mathstr);
                    RowSize = tmpMatrix.GetLength(0); //Information.UBound(tmpMatrix, 1);
                    ColSize = tmpMatrix.GetLength(1);//Information.UBound(tmpMatrix, 2);
                    if (RowSize <= 12 && ColSize <= 12)
                    {
                        for (j = 1; j <= RowSize; j++)
                        {
                            for (k = 1; k <= ColSize; k++)
                            {
                                BuildEq += tmpMatrix(j, k);
                                if (k < ColSize)
                                    BuildEq += "&";
                            }
                            if (j < RowSize)
                                BuildEq += "@";
                        }
                    }
                    else
                        BuildEq = "Mat_ij";
                }
                else
                    BuildEq = MatElements(Mathstr, Request);

                if (xlFunction == "MINVERSE")
                    BuildEq = @"[\matrix(" + BuildEq + ")]^-1 ";
                else if (xlFunction == "MDETERM")
                    BuildEq = @"|\matrix(" + BuildEq + ")| ";
                else if (xlFunction == "TRANSPOSE")
                {
                    if (!BuildEq.Contains(@"[\matrix"))
                        BuildEq = @"[\matrix(" + BuildEq + ")]^-T ";
                }

                tempequation = tempequation.Replace(TruncEqu, BuildEq);
            }
            return tempequation;
        }

        private string ConvertMMult(string Equation, string xlFunction, string listSeparator, string Request)
        {
            int count;
            string tempequation;
            int i;
            int j;
            int k;
            int l;
            int RowSize;
            int ColSize;
            string TruncEqu;
            string[] BuildEq = new string[3];
            dynamic tmpMatrix;
            string[] Params;
            string Mathstr;
            string MmultEqn;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            count = CountOccurance(Equation, xlFunction);
            tempequation = Equation;
            Mathstr = "";

            for (i = 1; i <= count; i++)
            {
                TruncEqu = TrimFunction(tempequation, xlFunction);
                Mathstr = TruncEqu.Substring(xlFunction.Length + 1, TruncEqu.Length - (xlFunction.Length + 2));

                Params = Mathstr.Split(new[] { listSeparator }, StringSplitOptions.None);
                for (j = 0; j < Params.Length; j++)
                {
                    Mathstr = Params[j];
                    if (Mathstr.Contains("["))
                        BuildEq[j] = Mathstr.Replace(" ", "");
                    else if (Mathstr.Contains("("))
                    {
                        tmpMatrix = xlApp.Evaluate(Mathstr);
                        RowSize = tmpMatrix.GetLength(0); //Information.UBound(tmpMatrix, 1);
                        ColSize = tmpMatrix.GetLength(1);//Information.UBound(tmpMatrix, 2);
                        if (RowSize <= 12 && ColSize <= 12)
                        {
                            if (Mathstr.Contains("TRANSPOSE"))
                                BuildEq[j] = ConvertMInverseDet(Mathstr, "TRANSPOSE", listSeparator, Request);
                            else
                            {
                                for (k = 1; k <= RowSize; k++)
                                {
                                    for (l = 1; l <= ColSize; l++)
                                    {
                                        BuildEq[j] = BuildEq[j] + tmpMatrix(k, l);
                                        if (l < ColSize)
                                            BuildEq[j] = BuildEq[j] + "&";
                                    }
                                    if (k < RowSize)
                                        BuildEq[j] = BuildEq[j] + "@";
                                }
                                BuildEq[j] = @"[\matrix(" + BuildEq[j] + ")] ";
                            }
                        }
                        else
                        {
                            BuildEq[j] = "Mat_ij";
                            BuildEq[j] = @"[\matrix(" + BuildEq[j] + ")] ";
                        }
                    }
                    else
                    {
                        BuildEq[j] = MatElements(Mathstr, Request);
                        BuildEq[j] = @"[\matrix(" + BuildEq[j] + ")] ";
                    }
                }
                MmultEqn = BuildEq[0] + @"\times " + BuildEq[1];
                tempequation = tempequation.Replace(TruncEqu, MmultEqn);
            }

            return tempequation;
        }


        private long[] MatrixSize(string StartAddress, string EndAddress)
        {
            long StartRow;
            long StartColumn;
            long EndRow;
            long EndColumn;
            long RowSize;
            long ColSize;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            StartRow = xlApp.Range[StartAddress].Row;
            StartColumn = xlApp.Range[StartAddress].Column;
            EndRow = xlApp.Range[EndAddress].Row;
            EndColumn = xlApp.Range[EndAddress].Column;

            RowSize = (EndRow - StartRow) + 1;
            ColSize = (EndColumn - StartColumn) + 1;

            return new[] { RowSize, ColSize };
        }


        private string MatElements(string Equation, string Request)
        {
            string tempequation;
            //string Mathstr; Mathstr = "";
            string[] splitRng;
            string BuildEq; BuildEq = "";
            int CountRange;
            string StartAddress;
            string EndAddress;
            long StartRow;
            long StartColumn;
            long[] MatSize;
            long RowNdx;
            long ColNdx;
            dynamic MatElm;
            string Reference;
            Excel.Worksheet wrkSheet;
            string RngAddr;
            Excel.Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            if (Equation.Contains("!"))
            {
                Reference = Equation.Split('!')[0];
                tempequation = Equation.Split('!')[1];
            }
            else
            {
                Reference = ActiveSheet.Name;
                tempequation = Equation;
            }
            wrkSheet = wb.Worksheets[Reference];


            CountRange = tempequation.Length - tempequation.Replace(":", "").Length;

            if (CountRange != 0)
            {
                if (tempequation.Contains(':'))
                {
                    splitRng = tempequation.Split(':');
                    StartAddress = splitRng[0];
                    EndAddress = splitRng[1];
                    StartRow = xlApp.Range[StartAddress].Row;
                    StartColumn = xlApp.Range[StartAddress].Column;

                    MatSize = MatrixSize(StartAddress, EndAddress);
                    for (RowNdx = StartRow; RowNdx <= StartRow + MatSize[0] - 1; RowNdx++)
                    {
                        for (ColNdx = StartColumn; ColNdx <= StartColumn + MatSize[1] - 1; ColNdx++)
                        {
                            if (Request == "FORMULA")
                            {
                                MatElm = wrkSheet.Cells[RowNdx, ColNdx].Formula;
                                if (MatElm.StartsWith("="))
                                    MatElm = MatElm.ToString().TrimStart('=');
                                RngAddr = xlApp.Cells[RowNdx, ColNdx].Address;
                                MatElm = ReplaceMatElems(wrkSheet, RngAddr, MatElm, Request);
                            }
                            else
                            {
                                MatElm = wrkSheet.Cells[RowNdx, ColNdx].Formula;
                                if (MatElm.StartsWith("="))
                                    MatElm = MatElm.ToString().TrimStart('=');

                                RngAddr = xlApp.Cells[RowNdx, ColNdx].Address;
                                MatElm = ReplaceMatElems(wrkSheet, RngAddr, MatElm, Request);
                            }

                            BuildEq += MatElm;

                            if (ColNdx < StartColumn + MatSize[1] - 1)
                                BuildEq += "&";
                        }
                        if (RowNdx < StartRow + MatSize[0] - 1)
                            BuildEq += "@";
                    }
                    BuildEq = BuildEq;
                }
                else
                    BuildEq = BuildEq;
            }

            return BuildEq;
        }

        private dynamic ReplaceMatElems(Excel.Worksheet wrkSheet, string RngAddr, string MatElm, string Request)
        {
            string Symbol;
            Excel.Range precRng;
            //Excel.Range targetRng;
            string CellAddress;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            wrkSheet.Select();

            if (wrkSheet.Range[RngAddr].HasFormula)
            {
                if (!MatElm.Contains(","))
                {
                    precRng = wrkSheet.Range[RngAddr].DirectPrecedents;
                    foreach (Excel.Range targetRng in precRng)
                    {
                        CellAddress = targetRng.Address.Replace("$", "");
                        if (Request == "FORMULA")
                        {
                            Symbol = GetEqSymbol(wrkSheet, CellAddress, Request);
                            if (Symbol != "")
                                MatElm = MatElm.Replace(CellAddress, Symbol);
                            else
                                MatElm = MatElm.Replace(CellAddress, Math.Round(wrkSheet.Range[CellAddress].Value, 2));
                        }
                        else
                            MatElm = MatElm.Replace(CellAddress, Convert.ToString(Math.Round(wrkSheet.Range[CellAddress].Value, 2)));
                    }
                    return MatElm;
                }
                else
                    return wrkSheet.Range[RngAddr].Value;
            }
            else if (Request == "FORMULA")
                return GetEqSymbol(wrkSheet, RngAddr, Request);
            else
                return wrkSheet.Range[RngAddr].Value;
        }
    }

    public class MathConverter
    {
        private string FontName 
        { get; 
           
            set; }
        public MathConverter()
        {

        }
        public MathConverter(string UserFont)
        {
            FontName = UserFont;
        }

        public void MathEquation(Word.Application wrdApp, Word.Document wrdDoc, string PrintEquation, Dictionary<string, string> DictAutoCorrect, string InlineText = "YES", bool usrfont=false)
        {
            Word.Range objRange;
            Word.OMath objEq;
            wrdApp.OMathAutoCorrect.UseOutsideOMath = true;


            objRange = wrdDoc.Range();
            wrdApp.ActiveDocument.Characters.Last.Select();
            wrdApp.Selection.Collapse();

            if (InlineText == "NO")
            {
                objRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
                objRange.MoveEnd();
                objRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
            }
            else
            {
                objRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);
                objRange.MoveEnd();
                objRange.InsertParagraphAfter();
                objRange.Collapse(Direction: Word.WdCollapseDirection.wdCollapseEnd);

            }

            objRange.Text = PrintEquation;

            if (PrintEquation.Contains(@"\"))
            {
                foreach (string AutoKey in DictAutoCorrect.Keys)
                {
                    if (objRange.Text.Contains(AutoKey))
                    {
                        objRange.Text = objRange.Text.Replace(AutoKey, DictAutoCorrect[AutoKey]);
                    }
                    if (!objRange.Text.Contains(@"\"))
                    {
                        break;
                    }
                }

            }
            objRange = wrdApp.Selection.OMaths.Add(objRange);
            objEq = objRange.OMaths[1];
            objEq.BuildUp();

            if (usrfont)
            {
                objEq.ConvertToNormalText();
                objEq.Range.Font.Name=FontName;
            }
        }

        public void TableEquation(Word.Application wrdApp, Word.Range objRange, string PrintEquation, Dictionary<string, string> DictAutoCorrect)
        {
            Word.OMath objEq;
            //Word.OMathAutoCorrectEntry aCorrect;
            wrdApp.OMathAutoCorrect.UseOutsideOMath = true;

            objRange.Text = PrintEquation;

            if (PrintEquation.Contains(@"\"))
            {
                foreach (string AutoKey in DictAutoCorrect.Keys)
                {
                    if (objRange.Text.Contains(AutoKey))
                    {
                        objRange.Text = objRange.Text.Replace(AutoKey, DictAutoCorrect[AutoKey]);
                    }
                    if (!objRange.Text.Contains(@"\"))
                    {
                        break;
                    }
                }

            }
            objRange = objRange.OMaths.Add(objRange);
            objEq = objRange.OMaths[1];
            objEq.BuildUp();

            //The below code is to keep the appearance of the table headers same as that of the texts.
            objEq.ConvertToNormalText();
            objEq.Range.Font.Name= FontName;
        }

    }
}
