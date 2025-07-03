using Nastranh5;
using StressUtilities;
using StressUtilities.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/** 
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

namespace FEM
{
    class LoadCombination
    {
        readonly int _ToTalLoadTypes = 4;

        private int NumberOfLoadType
        {
            get
            {
                return _ToTalLoadTypes;
            }
        }

        public LoadCombination()
        {

        }



        public void LaunchCombiForm()
        {
            IEnumerable<CombinationForm> FrmCollection = System.Windows.Forms.Application.OpenForms.OfType<CombinationForm>();
            if (FrmCollection.Any())
                FrmCollection.First().Focus();
            else
            {
                CombinationForm Combiform = new CombinationForm();
                Combiform.Show();
            }
        }


        public void CombineLoads(string FilePath, string[] FileList, string DataSource, string LCFileName, string LCTypesList, string ThermCaseList, string UnitThermalLoads, string ElemType, string ElemList, string OperationType, string ImportOption, string MapFileName)
        {
            //long RowNdx;
            //long ColNdx;
            int NCases;
            CombinationForm CombinationForm1 = new CombinationForm();
            string[] HeadersText = { };
            Dictionary<string, dynamic> ElmDict = null;
            Dictionary<string, dynamic> LoadCombinationDict;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            Excel.Range Rng;
            // Dim rptFileList() As String
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            
            
            LCTable LCTbl = new LCTable();

            try
            {
                CombinationForm1.WindowState = FormWindowState.Minimized;
                Rng = wb.Application.InputBox("Select the Start Cell of the Table.", "Obtain Range Object", Type: 8);
                CombinationForm1.WindowState = FormWindowState.Normal;
            }
            catch (Exception Ex)
            {
                Rng = null;
            }

            if (Rng == null)
            {
                MessageBox.Show("Cancelled by the user. File not imported");
                return;
            }
            xlApp.StatusBar = "Preparing Inputs...";

            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            xlApp.ScreenUpdating = false;
            // ReDim rptFileList(UBound(FileList))

            for (int i = 0; i < FileList.Length; i++)
                FileList[i] = Path.Combine(FilePath, FileList[i]);

            NCases = NumberOfLoadType;  // To be checked

            // --------------------------------------------------------
            LoadCombinationDict = LCTbl.LoadCases(NCases, ref LCFileName, ref HeadersText);  // To be added
            if (LoadCombinationDict == null)
                return;

            switch (DataSource)
            {
                case "rpt":
                    {
                        Readrpt RPTData = new Readrpt();
                        ElmDict = RPTData.ImportRPTfiles(Rng, FileList);
                        break;
                    }

                case "csv":
                    {
                        ReadCSV CSVData = new ReadCSV();
                        // Dim MapFileName As String = "E:\12_Demo\ParameterMap.smp"
                        ElmDict = CSVData.importCSVfiles(Rng, FileList, MapFileName);
                        break;
                    }
                case "h5":
                    {
                        H5ToDict H5Data = new H5ToDict();
                        // Dim MapFileName As String = "E:\12_Demo\ParameterMap.smp"
                        ElmDict = H5Data.importH5files(Rng, FileList, MapFileName);
                        break;
                    }
            }

            if (ElmDict == null && NCases >= 1)
                return;
            // -------------------------------------------------------
            long RowNdx = Rng.Row;
            long ColNdx = Rng.Column;

            xlApp.StatusBar = @"Combining Results...";
            CombineFEMData(Rng, ElmDict, LoadCombinationDict, FilePath, FileList, DataSource, LCFileName, LCTypesList, ThermCaseList, UnitThermalLoads, ElemType, ElemList, OperationType, ImportOption, HeadersText, MapFileName);

            ElmDict = null;
            LoadCombinationDict = null;

            xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            xlApp.ScreenUpdating = true;
            xlApp.StatusBar = null;
            MessageBox.Show(@"Loads are combined");
        }

        public void CombineFEMData(Excel.Range Rng, Dictionary<string, dynamic> elemDict, Dictionary<string, dynamic> LoadCombinationDict, string FilePath, string[] FileList, string DataSource, string LCFileName, string LCTypesList, string ThermCaseList, string UnitThermalLoads, string ElemType, string ElemList, string OperationType, string ImportOption, string[] HeadersText, string MapFileName)
        {
            //int i;
            // ------------Integers-----------------
            //int k;
            // ------------Double-----------------
            double CombinedLoad;
            double[] UnitThermalLoad = new double[4];
            List<string> UTLStr;
            string[] LoadCasesList;
            // ------------Arrays-----------------
            string[] FactorKeys;
            // -----------List ----------
            List<long> ElementList = new List<long>();
            List<long> LoadSources = new List<long>();
            // ------------strings-----------------
            string KeyThermal = "";
            //object ElemKey;
            string LCKey = "";

            // ------------Object-----------------
            //object CompKey;
            // ------------Dictionaries-----------------
            // Dim elemDict As Dictionary(Of String, Object)
            Dictionary<string, dynamic> LoadCaseDict;
            Dictionary<string, dynamic> combinedDict = new Dictionary<string, dynamic>();
            Dictionary<string, dynamic> ComponentDict;
            Dictionary<string, dynamic> combinedDictThermal = new Dictionary<string, dynamic>();
            Dictionary<string, dynamic> AverageDict = new Dictionary<string, dynamic>();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string LCSummary;
            Dictionary<string, string> MapDict = new Dictionary<string, string>();
            List<string> Tensor = new List<string>(new string[] { "X1", "Y1", "TXY1", "X2", "Y2", "TXY2", "X", "Y", "Z", "TXY", "TYZ", "TZX", "BM1", "BM2", "TS1", "TS2", "AF", "TTRQ", "WTRQ", "X1R", "X1I", "Y1R", "Y1I", "TXY1R", "TXY1I", "X2R", "X2I", "Y2R", "Y2I", "TXY2R", "TXY2I", "XR", "YR", "ZR", "TXYR", "TYZR", "TZXR", "XI", "YI", "ZI", "TXYI", "TYZI", "TZXI", "MY", "MZ", "QY", "QZ", "NX", "TX", "VB", "MB" });
            // -----End Declaration--------------------

            // Element_Count = UBound(ElementList) + 1
            ReadCSV CSVData = new ReadCSV();
            //General CommonData = new General();

            if (File.Exists(MapFileName))
                MapDict = CSVData.ParamMap(MapFileName);


            LoadSources = General.GetEntityList(LCTypesList);


            LoadCasesList = new string[LoadSources.Count - 1 + 1];
            FactorKeys = new string[LoadSources.Count - 1 + 1];

            if (ElemList.ToUpper() != "ALL")
                ElementList = General.GetEntityList(ElemList);

            for (int k = 0; k <= LoadSources.Count - 1; k++) // UBound(LoadSources)
            {
                LoadCasesList[k] = "LC" + System.Convert.ToString(LoadSources[k]);
                FactorKeys[k] = "LF" + System.Convert.ToString(LoadSources[k]);
            }

            if (ThermCaseList != "")
            {
                if (int.TryParse(ThermCaseList, out int _))
                    KeyThermal = "DT" + ThermCaseList.ToString();
                else
                {
                    MessageBox.Show(@"More than one thermal profile is not accounted by the utility. Only the first profile will be considered");
                    KeyThermal = "DT" + ThermCaseList.Substring(0, 1);
                }

                UTLStr = General.GetUnitThermalRanges(ref UnitThermalLoads);

                if (UTLStr.Count >= 3)
                {
                    UnitThermalLoad[0] = wb.ActiveSheet.range(UTLStr[0]).value;
                    UnitThermalLoad[1] = wb.ActiveSheet.range(UTLStr[1]).value;
                    UnitThermalLoad[2] = wb.ActiveSheet.range(UTLStr[2]).value;
                }
                else if (UTLStr.Count == 2)
                {
                    UnitThermalLoad[0] = wb.ActiveSheet.range(UTLStr[0]).value;
                    UnitThermalLoad[1] = wb.ActiveSheet.range(UTLStr[1]).value;
                    UnitThermalLoad[2] = 0;
                }
                else
                {
                    UnitThermalLoad[0] = wb.ActiveSheet.range(UnitThermalLoads).value;
                    UnitThermalLoad[1] = 0;
                    UnitThermalLoad[2] = 0;
                }
            }


            foreach (string ElemKey in elemDict.Keys)
            {
                LoadCaseDict = new Dictionary<string, dynamic>();
                // For Each LCKey In elemDict(ElemKey).Keys
                foreach (string LSKey in LoadCombinationDict.Keys)
                {
                    if (LoadCombinationDict.ContainsKey(LSKey))
                    {
                        LCSummary = "";
                        for (int i = 0; i < LoadCasesList.Length; i++)
                        {
                            // If LoadCombinationDict(LSKey)(LoadCasesList(i)) <> "" Then
                            if (!LoadCombinationDict[LSKey](LoadCasesList[i]) == null)
                            {
                                if (LCSummary == "")
                                    LCSummary = LoadCombinationDict[LSKey](FactorKeys[i]) + " x " + LoadCombinationDict[LSKey](LoadCasesList[i]);
                                else
                                    LCSummary = LCSummary + " + " + LoadCombinationDict[LSKey](FactorKeys[i]) + " x " + LoadCombinationDict[LSKey](LoadCasesList[i]);
                            }
                        }
                        ComponentDict = new Dictionary<string, dynamic>();

                        LCKey = LoadCombinationDict[LSKey](LoadCasesList[0]);
                        // ComponentDict.Add(LCKey, LCKey)
                        foreach (string CompKey in elemDict[ElemKey][LCKey].Keys)
                        {
                            switch (DataSource)
                            {
                                case "rpt":
                                    {
                                        if (double.TryParse(Convert.ToString(elemDict[ElemKey][LCKey][CompKey]), out double _))
                                        {
                                            CombinedLoad = CombineLoadCases(ref elemDict, ref LoadCombinationDict, ElemKey, LSKey, CompKey, ref LoadCasesList, ref FactorKeys, ref elemDict, DataSource);
                                            if (double.TryParse(CombinedLoad.ToString(), out double _))
                                                ComponentDict.Add(CompKey, CombinedLoad);
                                        }
                                        else if (CompKey == "DESCRIPTION")
                                            ComponentDict.Add(CompKey, LCSummary.Trim());
                                        else
                                            ComponentDict.Add(CompKey, elemDict[ElemKey][LCKey][CompKey]);
                                        break;
                                    }

                                case "csv":
                                    {
                                        if (Tensor.Contains(CompKey))
                                        {
                                            CombinedLoad = CombineLoadCases(ref elemDict, ref LoadCombinationDict, ElemKey, LSKey, CompKey, ref LoadCasesList, ref FactorKeys, ref elemDict, DataSource);
                                            if (double.TryParse(CombinedLoad.ToString(), out _))
                                                ComponentDict.Add(CompKey, CombinedLoad);
                                        }
                                        else if (MapDict.ContainsKey(CompKey))
                                        {
                                            if (Tensor.Contains(MapDict[CompKey]))
                                            {
                                                CombinedLoad = CombineLoadCases(ref elemDict, ref LoadCombinationDict, ElemKey, LSKey, CompKey, ref LoadCasesList, ref FactorKeys, ref elemDict, DataSource);
                                                if (double.TryParse(CombinedLoad.ToString(), out _))
                                                    ComponentDict.Add(CompKey, CombinedLoad);
                                            }
                                            else if (CompKey == "DESCRIPTION")
                                                ComponentDict.Add(CompKey, LCSummary.Trim());
                                            else if (!(CompKey == "MAX_PRINCIPAL" || CompKey == "MID_PRINCIPAL" || CompKey == "VONMISES" || CompKey == "MIN_PRINCIPAL"))
                                                ComponentDict.Add(CompKey, elemDict[ElemKey][LCKey][CompKey]);
                                        }
                                        else if (CompKey == "DESCRIPTION")
                                            ComponentDict.Add(CompKey, LCSummary.Trim());
                                        else if (!(CompKey == "MAX_PRINCIPAL" || CompKey == "MID_PRINCIPAL" || CompKey == "VONMISES" || CompKey == "MIN_PRINCIPAL"))
                                            ComponentDict.Add(CompKey, elemDict[ElemKey][LCKey][CompKey]);
                                        break;
                                    }
                            }
                        }

                        if (!ComponentDict.ContainsKey("DESCRIPTION"))
                            ComponentDict.Add("DESCRIPTION", LCSummary.Trim());

                        LoadCaseDict.Add(LSKey, ComponentDict);
                    }
                }
                combinedDict.Add(ElemKey, LoadCaseDict);
            }

            // elemDict = Nothing

            int AddThermal = 0;
            int AvgOption = 0;
            // ---------------------------------------------------------------
            switch (ElemType)
            {
                case "1D":
                case "NODE":
                    {
                        if (ThermCaseList != "")
                        {
                            combinedDictThermal = ThermalCombination(ref combinedDict, ref LoadCombinationDict, ref KeyThermal, ref UnitThermalLoad);
                            AddThermal = 1;
                        }

                        break;
                    }

                case "2D":
                    {
                        if (OperationType == "CombineAverage")
                        {
                            // CompKey2D = {"STRESSX1", "STRESSY1", "STRESSXY1", "MAXPRINCZ1", "MINPRINCZ1", "ANGLE1",
                            // "STRESSX2", "STRESSY2", "STRESSXY2", "MAXPRINCZ2", "MINPRINCZ2", "ANGLE2"}
                            if (ElemList.ToUpper() != "ALL")
                            {
                                AverageDict = AverageStress2D(ref combinedDict, ref ElementList, DataSource);
                                AvgOption = 1;
                            }
                        }

                        if (ThermCaseList != "")
                        {
                            combinedDictThermal = ThermalCombination(ref combinedDict, ref LoadCombinationDict, ref KeyThermal, ref UnitThermalLoad);
                            AddThermal = 1;
                        }

                        break;
                    }
            }
            // -------------------------------------------------------------------

            PrintCombinedResults(elemDict, combinedDict, Rng, LoadCombinationDict, LoadCasesList, FactorKeys, combinedDictThermal, AddThermal, ref ElementList, AverageDict, AvgOption, HeadersText, DataSource);


            // elemDict = Nothing

            combinedDict = null;
            ComponentDict = null;
            LoadCaseDict = null;
        }


        public void PrintCombinedResults(Dictionary<string, dynamic> rptElemDict, Dictionary<string, dynamic> combinedDict, Excel.Range Rng, Dictionary<string, dynamic> LoadCombinationDict, string[] LoadCasesList, string[] FactorKeys, Dictionary<string, dynamic> combinedDictThermal, int AddThermal, ref List<long> ElementList, Dictionary<string, dynamic> AverageDict, int AvgOption, string[] HeadersText, string DataSource)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string SheetName;
            string CellName;
            long RowNdx;
            long StartRow;
            long ColNdx;
            long StartCol;
            Excel.Worksheet TblSheet;
            string[] TblHeaders;
            // Dim LoadType As String
            int NCases = LoadCasesList.Length;
            List<string> LCList = new List<string>();
            //int m = 0;
            Dictionary<string, dynamic> tmpcombinedDict = new Dictionary<string, dynamic>();
            string ElmList = "";
            string LCID;


            TblSheet = wb.Worksheets[Rng.Parent.Name];

            TblSheet.Select();
            RowNdx = Rng[1, 1].Row + 2;

            ColNdx = Rng[1, 1].Column;
            StartCol = ColNdx;
            StartRow = RowNdx;

            SheetName = Rng.Parent.Name;
            CellName = Rng.Address.Replace("$", "");



            for (int j = 0; j <= NCases + 1 + AddThermal + AvgOption; j++)
            {
                if (j <= 0 + AddThermal + AvgOption)
                {
                    if (j == 0 && AvgOption == 1)
                    {
                        TblHeaders = TableHeader(ref AverageDict, true).Split(';');
                        for (int k = 0; k <= ElementList.Count - 1; k++)
                            ElmList = ElmList + ';' + ElementList[k];
                    }
                    else if (j == 0 + AvgOption && AddThermal == 1)
                        TblHeaders = TableHeader(ref combinedDictThermal, true).Split(';');
                    else
                        TblHeaders = TableHeader(ref combinedDict, true).Split(';');
                }
                else if (DataSource == "rpt")
                    TblHeaders = TableHeader(ref rptElemDict, true).Split(';');
                else
                    TblHeaders = TableHeader(ref rptElemDict, false).Split(';');

                for (int i = 0; i < TblHeaders.Length; i++)
                    TblSheet.Cells[RowNdx, StartCol + i].Value = TblHeaders[i];
                RowNdx++;

                if (j > 0 + AddThermal + AvgOption)
                {
                    tmpcombinedDict = null;
                    TblSheet.Cells[RowNdx - 2, StartCol].Value = "Load type: " + HeadersText[(j - AddThermal - AvgOption) * 2 - 1]; // LoadCasesList(j - 1 - AddThermal - AvgOption) 

                    foreach (string ElemKey in rptElemDict.Keys)
                    {
                        // For Each LCKey In rptElemDict(ElemKey).Keys
                        // m = 0
                        foreach (string LCKey in LoadCombinationDict.Keys)
                        {
                            if (LoadCombinationDict[LCKey](LoadCasesList[j - 1 - AddThermal - AvgOption]) != "")
                            {
                                // ReDim Preserve LCList(m)
                                LCID = LoadCombinationDict[LCKey](LoadCasesList[j - 1 - AddThermal - AvgOption]);
                                if (!LCList.Contains(LCID))
                                    LCList.Add(LCID);
                            }
                        }

                        for (int k = 0; k <= LCList.Count - 1; k++) // UBound(LCList)
                        {
                            ColNdx = StartCol;
                            if (rptElemDict[ElemKey].ContainsKey(LCList[k]))
                            {
                                TblSheet.Cells[RowNdx, ColNdx].Value = ElemKey;
                                ColNdx++;
                                // If DataSource = "cvs" Then
                                // TblSheet.Cells(RowNdx, ColNdx) = LCList(k)
                                // ColNdx +=  1
                                // End If

                                foreach (string CompKey in rptElemDict[ElemKey](LCList[k]).Keys)
                                {
                                    if (CompKey != "DESCRIPTION")
                                    {
                                        // Line = Line && ';' && rptElemDict(ElemKey)[LCKey][CompKey]
                                        TblSheet.Cells[RowNdx, ColNdx].Value = rptElemDict[ElemKey](LCList[k])[CompKey];
                                        ColNdx++;
                                    }
                                }
                                if (rptElemDict[ElemKey](LCList[k]).ContainsKey("DESCRIPTION"))
                                    TblSheet.Cells[RowNdx, ColNdx].Value = rptElemDict[ElemKey](LCList[k])("DESCRIPTION");
                                // TblSheet.Cells(RowNdx, ColNdx) = rptElemDict(ElemKey)(LCList(k))("DESCRIPTION")
                                RowNdx++;
                            }
                        } // LCKey
                          // Array.Clear(LCList, 0, LCList.Length)
                        LCList.Clear();
                    }
                }
                else
                {
                    if (j == 0 && AvgOption == 1)
                    {
                        tmpcombinedDict = AverageDict;
                        TblSheet.Cells[RowNdx - 2, StartCol].Value = "Average Loads/Stresses";
                        TblSheet.Cells[RowNdx - 3, StartCol].Value = "Element List: " + ElmList;
                    }
                    else if (j == 0 + AvgOption && AddThermal == 1)
                    {
                        tmpcombinedDict = combinedDictThermal;
                        TblSheet.Cells[RowNdx - 2, StartCol].Value = "Combined Loads+Thermal Loads";
                    }
                    else
                    {
                        TblSheet.Cells[RowNdx - 2, StartCol].Value = "Combined Loads without Thermal Loads";
                        tmpcombinedDict = combinedDict;
                    }


                    foreach (string ElemKey in tmpcombinedDict.Keys)
                    {
                        foreach (string LCKey in tmpcombinedDict[ElemKey].Keys)
                        {
                            ColNdx = StartCol;
                            TblSheet.Cells[RowNdx, ColNdx].Value = ElemKey;
                            ColNdx++;
                            foreach (string CompKey in tmpcombinedDict[ElemKey][LCKey].Keys)
                            {
                                if (CompKey != "DESCRIPTION")
                                {
                                    TblSheet.Cells[RowNdx, ColNdx].Value = tmpcombinedDict[ElemKey][LCKey][CompKey];
                                    ColNdx++;
                                }
                            }
                            if (tmpcombinedDict[ElemKey][LCKey].ContainsKey("DESCRIPTION"))
                                TblSheet.Cells[RowNdx, ColNdx].Value = tmpcombinedDict[ElemKey][LCKey]("DESCRIPTION");
                            RowNdx++;
                        }
                    }
                }

                long EndCol = StartCol - 1 + TblSheet.Range[TblSheet.Cells[StartRow, StartCol], TblSheet.Cells[StartRow, StartCol].End(Excel.XlDirection.xlToRight)].Count;
                long EndRow = StartRow - 1 + TblSheet.Range[TblSheet.Cells[StartRow, StartCol], TblSheet.Cells[StartRow, StartCol].End(Excel.XlDirection.xlDown)].Count;

                Excel.Range Selection = TblSheet.Range[TblSheet.Cells[StartRow, StartCol], TblSheet.Cells[EndRow, EndCol]]; // .Select

                // With Selection

                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                Selection.HorizontalAlignment = Excel.Constants.xlCenter;
                if (!(j == 0 && AvgOption == 1))
                    Selection.EntireColumn.AutoFit();

                StartCol = EndCol + 2;
                RowNdx = StartRow;
            }
        }

        //what is the purpose of the function below?
        //public void Component(ref string[] StressData, ref string[] FileHeaderKeys)
        //{
        //    int i;
        //    Dictionary<string, string>  ComponentDict = new Dictionary<string, string>();
        //    for (i = 1; i <StressData.Length; i++)
        //        ComponentDict.Add(FileHeaderKeys[i], StressData[i]);
        //}

        public dynamic CombineLoadCases(ref Dictionary<string, dynamic> ElmDict, ref Dictionary<string, dynamic> LoadCombinationDict, string ElemKey, string LCKey, string CompKey, ref string[] LoadCasesList, ref string[] FactorsKeys, ref Dictionary<string, dynamic> rptElementDict, string DataSource)
        {
            //int i;
            string LoadCase;
            double Factor = 1.0;
            double Component = 0.0;
            bool BoolCombination = false;
            double CombineCases = 0;
            for (int i = 0; i < LoadCasesList.Length; i++)
            {
                if (LoadCombinationDict[LCKey](LoadCasesList[i]) != "")
                {
                    LoadCase = LoadCombinationDict[LCKey](LoadCasesList[i]);
                    double.TryParse(LoadCombinationDict[LCKey](FactorsKeys[i]), out Factor);
                    if (ElmDict[ElemKey].ContainsKey(LoadCase))
                        double.TryParse(ElmDict[ElemKey][LoadCase][CompKey], out Component);
                    // Check for the strings
                    if (double.TryParse(Component.ToString(), out _))
                    {
                        switch (DataSource)
                        {
                            case "rpt":
                                {
                                    if (CompKey.Contains("Component") || CompKey.Contains("Rotation"))
                                    {
                                        CombineCases += Component * Factor;
                                        BoolCombination = true;
                                    }

                                    break;
                                }

                            case "csv":
                                {
                                    CombineCases += Component * Factor;
                                    BoolCombination = true;
                                    break;
                                }
                        }
                    }
                }
            }
            if (BoolCombination == false)
                return "None";

            return CombineCases;
        }


        public Dictionary<string, dynamic> AverageStress2D(ref Dictionary<string, dynamic> elemDict, ref List<long> ElementList, string DataSource)
        {
            Dictionary<string, dynamic> CompDict;
            Dictionary<string, dynamic> LCDict;
            int j;
            int nElements;
            //object LCKey;
            double CombinedLoad;
            Dictionary<string, dynamic> AverageStress;

            nElements = ElementList.Count;

            LCDict = new Dictionary<string, dynamic>();

            foreach (string LCKey in elemDict[ElementList[0].ToString()].Keys)
            {
                CompDict = new Dictionary<string, dynamic>();

                // CompDict.Add(LCKey, LCKey)
                foreach (string Compkey in elemDict[ElementList[0].ToString()][LCKey].keys)
                {
                    CombinedLoad = 0;
                    switch (DataSource)
                    {
                        case "rpt":
                            {
                                if (Compkey.Contains("Component"))
                                {
                                    for (j = 0; j <= ElementList.Count - 1; j++)
                                    {
                                        CombinedLoad += double.TryParse(elemDict[ElementList[j].ToString()][LCKey][Compkey], out double loadComp);
                                        CombinedLoad += loadComp;
                                    }
                                    CombinedLoad /= nElements;
                                    CompDict.Add(Compkey, CombinedLoad);
                                }

                                break;
                            }

                        case "csv":
                            {
                                switch (Compkey)
                                {
                                    case "X1":
                                    case "Y1":
                                    case "TXY1":
                                    case "X2":
                                    case "Y2":
                                    case "TXY2":
                                        {
                                            for (j = 0; j <= ElementList.Count - 1; j++)
                                            {
                                                CombinedLoad += double.TryParse(elemDict[ElementList[j].ToString()][LCKey][Compkey], out double loadComp);
                                                CombinedLoad += loadComp;
                                            }
                                            CombinedLoad /= nElements;
                                            CompDict.Add(Compkey, CombinedLoad);
                                            break;
                                        }
                                }

                                break;
                            }
                    }
                }
                CompDict.Add("DESCRIPTION", "Average Loads");
                // MAXPRINCZ1 = (STRESSX1 + STRESSY1) / 2 + (((STRESSX1 - STRESSY1) / 2) ^ 2 + STRESSXY1 ^ 2) ^ 0.5
                // MAXPRINCZ2 = (STRESSX2 + STRESSY2) / 2 + (((STRESSX2 - STRESSY2) / 2) ^ 2 + STRESSXY2 ^ 2) ^ 0.5
                // MINPRINCZ1 = (STRESSX1 + STRESSY1) / 2 - (((STRESSX1 - STRESSY1) / 2) ^ 2 + STRESSXY1 ^ 2) ^ 0.5
                // MINPRINCZ2 = (STRESSX2 + STRESSY2) / 2 - (((STRESSX2 - STRESSY2) / 2) ^ 2 + STRESSXY2 ^ 2) ^ 0.5
                // ANGLE1 = MaxPrincAngle(STRESSXY1, STRESSX1, STRESSY1)
                // ANGLE2 = MaxPrincAngle(STRESSXY2, STRESSX2, STRESSY2)

                // StressVector = {STRESSX1, STRESSY1, STRESSXY1, MAXPRINCZ1, MINPRINCZ1, ANGLE1,
                // STRESSX2, STRESSY2, STRESSXY2, MAXPRINCZ2, MINPRINCZ2, ANGLE2}

                LCDict.Add(LCKey, CompDict);
            }
            AverageStress = new Dictionary<string, dynamic>();
            AverageStress.Add(ElementList[0].ToString(), LCDict);
            CompDict = null;

            return AverageStress;
        }

        public double MaxPrincAngle(ref double STRESSXY, double STRESSX, ref double STRESSY)
        {
            double PrincAngle;
            if ((STRESSX - STRESSY) != 0)
                PrincAngle = Math.Atan(2 * STRESSXY / (STRESSX - STRESSY)) * 0.5 * 180.0 / Math.PI;
            else if (STRESSXY != 0)
                PrincAngle = 45.0;
            else
                PrincAngle = 0;
            return Math.Round(PrincAngle, 2);
        }

        public Dictionary<string, dynamic> ThermalCombination(ref Dictionary<string, dynamic> combinedDict, ref Dictionary<string, dynamic> LoadCombinationDict, ref string KeyThermal, ref double[] UnitThermalLoad)
        {
            Dictionary<string, dynamic> combinedDictThermal = new Dictionary<string, dynamic>();
            Dictionary<string, dynamic> ComponentDict;
            Dictionary<string, dynamic> LCDict;
            double UnitThermalX = UnitThermalLoad[0];
            double UnitThermalY = UnitThermalLoad[1];
            double UnitThermalZ = UnitThermalLoad[2];
            object CombinedVal;
            double deltaTemp;
            // Dim LCKey As String

            // combinedDictThermal = combinedDict

            foreach (string ElmKey in combinedDict.Keys)
            {
                LCDict = new Dictionary<string, dynamic>();
                foreach (string LCKey in combinedDict[ElmKey].Keys)
                {
                    // LCKey = LoadCombinationDict(LSKey)(1)
                    ComponentDict = new Dictionary<string, dynamic>();
                    foreach (string CompKey in combinedDict[ElmKey][LCKey].keys)
                    {
                        if (LoadCombinationDict[LCKey](KeyThermal) == "RT")
                            deltaTemp = 0;
                        else
                            deltaTemp = LoadCombinationDict[LCKey](KeyThermal);
                        switch (CompKey)
                        {
                            case "XComponent":
                            case "X1":
                            case "X2":
                            case "AF":
                                {
                                    CombinedVal = combinedDict[ElmKey][LCKey][CompKey] + deltaTemp * UnitThermalLoad[0];
                                    break;
                                }

                            case "YComponent":
                            case "Y1":
                            case "Y2":
                            case "TS1":
                                {
                                    CombinedVal = combinedDict[ElmKey][LCKey][CompKey] + deltaTemp * UnitThermalLoad[1];
                                    break;
                                }

                            case "ZComponent":
                            case "XYComponent":
                            case "TXY1":
                            case "TXY2":
                            case "TS2":
                                {
                                    CombinedVal = combinedDict[ElmKey][LCKey][CompKey] + deltaTemp * UnitThermalLoad[2];
                                    break;
                                }

                            default:
                                {
                                    CombinedVal = combinedDict[ElmKey][LCKey][CompKey];
                                    break;
                                }
                        }

                        // If InStr(CompKey, "XComponent") <> 0 Then
                        // CombinedVal = combinedDict(ElmKey)[LCKey][CompKey] + deltaTemp * UnitThermalLoad(0)
                        // ElseIf InStr(CompKey, "YComponent") <> 0 Then
                        // CombinedVal = combinedDict(ElmKey)[LCKey][CompKey] + deltaTemp * UnitThermalLoad(1)
                        // ElseIf InStr(CompKey, "ZComponent") <> 0 Or InStr(CompKey, "XYComponent") <> 0 Then
                        // CombinedVal = combinedDict(ElmKey)[LCKey][CompKey] + deltaTemp * UnitThermalLoad(2)
                        // Else
                        // CombinedVal = combinedDict(ElmKey)[LCKey][CompKey]
                        // End If

                        ComponentDict.Add(CompKey, CombinedVal);
                    }
                    ComponentDict.Add("DeltaTemp", LoadCombinationDict[LCKey](KeyThermal));
                    LCDict.Add(LCKey, ComponentDict);
                }
                combinedDictThermal.Add(ElmKey, LCDict);
            }

            return combinedDictThermal;
        }


        public string CriticalLoadCase(ref Dictionary<string, dynamic> combinedDict, string compKey)
        {
            //object LCKey;
            double MaxValue = 0, KeyValue = 0;
            string CriticalCase = "";
            foreach (string LCKey in combinedDict.Keys)
            {
                double.TryParse(combinedDict[LCKey][compKey], out KeyValue);
                if (KeyValue > MaxValue)
                {
                    CriticalCase = LCKey;
                    MaxValue = KeyValue;
                }
            }
            return CriticalCase;
        }

        public string TableHeader(ref Dictionary<string, dynamic> elemDict, bool descriptOpt)
        {
            string ElemKey = elemDict.Keys.First();
            string LcKey = ((Dictionary<string, dynamic>)elemDict[ElemKey]).Keys.First();
            string Header = "EntityID";

            foreach (string CompKey in elemDict[ElemKey][LcKey].Keys)
            {
                if (CompKey != "DESCRIPTION")
                    Header = Header + ';' + CompKey;
            }
            if (descriptOpt)
                Header = Header + ';' + "DESCRIPTION";

            return Header;
        }
    }
}
