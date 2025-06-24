using StressUtilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FEM
{
    class Readrpt
    {
        //static Regex regExSortLC = new Regex(@"^\s+Load Case:\s([A-Z 0-9 a-z]+):([\w\W\s\S\d\D]+)$");
        //static Regex regExSortLCGen = new Regex(@"^\s+Load Case:\s([A-Z 0-9 a-z]+),([\w\W\s\S\d\D]+)$");
        static Regex regExSortLC = new Regex(@"^\s+Load Case:\s([\w\W\s\S\d\D^:]+):([\w\W\s\S\d\D]+)$");  //To be  validated
        static Regex regExSortLCGen = new Regex(@"^\s+Load Case:\s([\w\W\s\S\d\D^:^,]+),([\w\W\s\S\d\D]+)$");//To be  validated
        static Regex regExLC = new Regex(@"^\s+(\d+)\s+([\w\W\s\S\d\D]{22})([\w\W\s\S\d\D]+)At\s+([A-Z a-z 0-9]+)$");
        static Regex regExLCNL = new Regex(@"^\s+(\d+)\s+([\w\W\s\S\d\D]{22})[\s+]?([\w\S\s]+\(NON\-LAYERED\))");    //NL- Non-Layered
        //static Regex regExLC = new Regex(@"^\s+(\d+)\s+([A-Z]{2}[0-9]+)([\w\W\s\S\d\D]+)At\s+([A-Z a-z 0-9]+)$");
        //static Regex regExLCNL = new Regex(@"^\s+(\d+)\s+([A-Z]{2}[0-9]+):[\s+]?([\w\S\s]+\(NON\-LAYERED\))")
        static Regex regEx1D = new Regex(@"^\s+Result\s(\S+)\s\S+\,\s(\S+)\s\S\sLayer\sAt\s(\S+)", RegexOptions.IgnorePatternWhitespace);
        static Regex regEx2D = new Regex(@"^\s+Result\s(\S+)\s\S+,\s+\S\sLayer\sAt\s\S(\d)", RegexOptions.IgnorePatternWhitespace);
        static Regex regExHeader = new Regex(@"^-Source\sID\W+Entity.ID([0-9 a-z A-Z\s-]*)");
        static Regex regExHeaderLC = new Regex(@"^\s+-Entity ID([0-9 a-z A-Z\s-]*)");
        static Regex regExValues;


        public Readrpt()
        {
        }


        public void ImportPatranRPTfile()
        {
            Excel.Range Rng;
            string[] rptFileList;

            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            Dictionary<string, dynamic> rptElemDict = new Dictionary<string, dynamic>();

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = wb.Path;  //GetFolderpath();
            openFileDialog1.Filter = @"Patran rpt files (*.rpt)|*.rpt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = true;
            openFileDialog1.RestoreDirectory = true;


            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    rptFileList = openFileDialog1.FileNames;
                }

                catch (Exception Ex)
                {
                    MessageBox.Show(@"Unable to read the rpt file(s) " + Ex.Message);
                    return;
                }
            }
            else
            {
                return;
            }

            try
            {
                Rng = wb.Application.InputBox("Select the Start Cell of the Table.", "Select Start Range", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);//to be checked 8
            }
            catch (Exception Ex)
            {
                Rng = null;
            }


            if (Rng == null)
            {
                MessageBox.Show("Cancelled by the user. File not imported");
            }
            else
            {
                xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;
                xlApp.ScreenUpdating = false;

                rptElemDict = ImportRPTfiles(Rng, rptFileList);  //ImportRPTfilesGeneral
                if (rptElemDict.Count != 0)
                {
                    xlApp.StatusBar = @"Preparing Loads Table";
                    PopulateResults(rptElemDict, Rng, true);
                    rptElemDict.Clear();

                    xlApp.StatusBar = false;
                    xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    xlApp.ScreenUpdating = true;

                    Marshal.ReleaseComObject(wb);
                    //MessageBox.Show(@"Patran .rpt File(s) Imported Successfully.");
                }
                else
                {
                    xlApp.StatusBar = false;
                    xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    xlApp.ScreenUpdating = true;

                    Marshal.ReleaseComObject(wb);
                    MessageBox.Show("The rpt files have not been read. Please check the rpt files for the \"Untitled\" load cases");
                }




            }


        }



        public Dictionary<string, dynamic> ImportRPTfiles(Excel.Range Rng, string[] rptFileList)
        {
            //int j;
            string TextLine = "";
            string ElemType = "";
            string LoadType = "Forces";
            string Layer = "";
            string[] rptHeaders;
            List<string> Headers = new List<string>();
            string LCDescription = "";
            string rptCompDictkey = "";

            bool ReadRPTinfoComplete;
            bool boolregEx1D;
            bool boolregEx2D;
            bool boolregExLC;
            bool boolregExHeader;
            bool boolElemTypeMatch = true;

            int LdCaseID;
            string LoadCase = "";
            string LayerLC = "";
            long ElementID = 0;

            Match match;
            string rptValue;
            string LOADCASEID = "";
            string SOURCEID = "";
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            string ElPosID = "";
            //string rptFile;
            bool ElPosIDChk = true;


            Match matches;
            GroupCollection submatches;


            Dictionary<string, string> LCParamDict = new Dictionary<string, string>();
            Dictionary<string, string> LCSortDict = new Dictionary<string, string>();
            Dictionary<string, dynamic> rptLCSortDict = new Dictionary<string, dynamic>();
            Dictionary<int, dynamic> rptLCDict = new Dictionary<int, dynamic>();
            Dictionary<string, string> rptCompDict = new Dictionary<string, string>();
            Dictionary<string, dynamic> rptLoadCaseDict = new Dictionary<string, dynamic>();
            Dictionary<string, dynamic> rptElemDict = new Dictionary<string, dynamic>();
            Dictionary<string, dynamic> rptElmPosDict = new Dictionary<string, dynamic>();
            Dictionary<string, string> dataDict = new Dictionary<string, string>();
            //Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string rptType = "SortByEntity";

            int FileCount = 0;
            int LineCount;
            Stopwatch sw = new Stopwatch();
            sw.Start();

            foreach (string rptFile in rptFileList)
            {
                ReadRPTinfoComplete = false;
                boolregEx1D = false;
                boolregEx2D = false;
                boolregExLC = false;
                boolregExHeader = false;
                FileCount++;

                using (StreamReader ainp = new StreamReader(rptFile)) //, Encoding.UTF8, true, 16384
                {
                    LineCount = 0;
                    rptLCDict = new Dictionary<int, dynamic>();
                    rptLoadCaseDict = new Dictionary<string, dynamic>();
                    xlApp.StatusBar = $"Processing File {FileCount}. Elapsed Time: { sw.Elapsed}.";//DateTime.Now.Subtract(startTime)

                    do
                    {
                        TextLine = ainp.ReadLine();

                        if (!boolregEx1D && (matches = regEx1D.Match(TextLine)).Success)
                        {
                            //matches = regEx1D.Matches(TextLine);
                            ElemType = matches.Groups[1].Value;
                            LoadType = matches.Groups[2].Value;
                            Layer = matches.Groups[3].Value;
                            boolregEx1D = true;
                        } //End of if
                        else if (!boolregEx2D && (matches = regEx2D.Match(TextLine)).Success)
                        {
                            //matches = regEx2D.Matches(TextLine);  //Check
                            ElemType = matches.Groups[1].Value;
                            LoadType = ElemType;
                            Layer = matches.Groups[2].Value;
                            boolregEx2D = true;
                        }//End of else if
                        else if (!boolregExLC && (matches = regExSortLC.Match(TextLine)).Success)
                        {
                            //matches = regExSortLC.Matches(TextLine);
                            LoadCase = matches.Groups[1].Value.Trim();
                            LCDescription = matches.Groups[2].Value;
                            boolregExLC = true;
                        }//End of else if

                        else if (!boolregExLC && (matches = regExSortLCGen.Match(TextLine)).Success)
                        {
                            //matches = regExSortLCGen.Matches(TextLine);
                            LoadCase = matches.Groups[1].Value.Trim();
                            LCDescription = matches.Groups[2].Value;
                            boolregExLC = true;
                        }//End of else if
                        else if ((matches = regExLC.Match(TextLine)).Success || (matches = regExLCNL.Match(TextLine)).Success)
                        {
                            //LoadCaseDict(matches, ref rptLCDict);
                            //boolregExLC = true;
                            string[] lcPart = new string[2];
                            LCDescription = "";
                            LdCaseID = int.Parse(matches.Groups[1].Value);
                            LoadCase = matches.Groups[2].Value.Trim();
                            if (LoadCase.Contains(":"))
                            {
                                lcPart = LoadCase.Split(new char[] { ':' }, 2, StringSplitOptions.RemoveEmptyEntries);
                                //LoadCase = LoadCase.Replace(":", "");
                                LoadCase = lcPart[0];
                                if (lcPart.Length > 1)
                                    LCDescription = lcPart[1];
                            }
                            LCDescription += matches.Groups[3].Value;
                            LayerLC = matches.Groups[4].Value.Replace(" ", "");

                            boolregExLC = true;

                            LCParamDict = new Dictionary<string, string>();

                            LCParamDict.Add("LOADCASE", LoadCase);
                            LCParamDict.Add("DESCRIPTION", LCDescription);
                            LCParamDict.Add("LCLAYER", LayerLC);

                            rptLCDict.Add(LdCaseID, LCParamDict);


                        }//End of else if
                        else if (!boolregExHeader && (regExHeader.IsMatch(TextLine) || regExHeaderLC.IsMatch(TextLine)))
                        {
                            rptHeaders = TextLine.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries); //new Char()
                            Headers = new List<string>();
                            Headers.AddRange(rptHeaders);
                            if (!Headers.Contains("El. Pos. ID"))
                            {
                                ElPosIDChk = false;
                            }

                            if (Headers[0] == "Entity ID")
                            {
                                rptType = "SortByLC";
                            }

                            Array.Clear(rptHeaders, 0, rptHeaders.Length);

                            regExValues = new Regex(rptResultsPattern(ref Headers) + ")", RegexOptions.Compiled);

                            boolregExHeader = true;
                        }//End of else if
                        else if (boolregExHeader == true)
                        {
                            if ((match = regExValues.Match(TextLine)).Success)
                            {
                                //match = regExValues.Match(TextLine);
                                submatches = match.Groups;
                                int j = 0;
                                foreach (Capture capture in submatches)
                                {
                                    try
                                    {
                                        if (rptType == "SortByEntity")
                                        {
                                            if (Headers[j] == "Source ID" || Headers[j] == "Loadcase ID")
                                            {
                                                LOADCASEID = rptLCDict[int.Parse(capture.Value)]["LOADCASE"];
                                                SOURCEID = capture.Value;
                                                Layer = rptLCDict[int.Parse(capture.Value)]["LCLAYER"];
                                            }
                                            else if (Headers[j] == "Entity ID")
                                            {
                                                ElementID = long.Parse(capture.Value);

                                                if (!rptElemDict.ContainsKey(ElementID.ToString()))
                                                {
                                                    rptCompDict = new Dictionary<string, string>();
                                                    rptLoadCaseDict = new Dictionary<string, dynamic>();
                                                    rptCompDict.Add("LoadCaseID", LOADCASEID);
                                                    rptCompDict.Add("DESCRIPTION", rptLCDict[int.Parse(SOURCEID)]["DESCRIPTION"]);
                                                    rptLoadCaseDict.Add(LOADCASEID, rptCompDict);
                                                    rptElemDict.Add(ElementID.ToString(), rptLoadCaseDict);
                                                }
                                            }
                                            else
                                            {
                                                if (!rptElemDict[ElementID.ToString()].ContainsKey(LOADCASEID))
                                                {
                                                    rptCompDict = new Dictionary<string, string>();
                                                    rptCompDict.Add("LoadCaseID", LOADCASEID);
                                                    rptCompDict.Add("DESCRIPTION", rptLCDict[int.Parse(SOURCEID)]["DESCRIPTION"]);
                                                    if (!rptElemDict[ElementID.ToString()].ContainsKey(LOADCASEID))
                                                        rptElemDict[ElementID.ToString()].Add(LOADCASEID, rptCompDict);
                                                }

                                                rptCompDictkey = GetRPTCompDictKey(Headers.ElementAt(j), ref ElemType, ref Layer, LoadType);

                                                if (rptElemDict[ElementID.ToString()][LOADCASEID].ContainsKey(rptCompDictkey))
                                                {
                                                    if (LoadType == "Rotational")
                                                    {
                                                        if (!rptElemDict[ElementID.ToString()][LOADCASEID].ContainsKey(rptCompDictkey + "P1"))
                                                        {
                                                            rptCompDictkey += "P1";
                                                        }
                                                        else if (!rptElemDict[ElementID.ToString()][LOADCASEID].ContainsKey(rptCompDictkey + "P2"))
                                                        {
                                                            rptCompDictkey += "P2";
                                                        }
                                                        rptElemDict[ElementID.ToString()][LOADCASEID].Add(rptCompDictkey, capture.Value);
                                                    }
                                                }
                                                else
                                                {
                                                    if (LoadType == "Rotational")
                                                    {
                                                        if (!rptElemDict[ElementID.ToString()][LOADCASEID].ContainsKey(rptCompDictkey + "P1"))
                                                        {
                                                            rptCompDictkey += "P1";
                                                        }
                                                        else if (!rptElemDict[ElementID.ToString()][LOADCASEID].ContainsKey(rptCompDictkey + "P2"))
                                                        {
                                                            rptCompDictkey += "P2";
                                                        }
                                                    }
                                                    rptElemDict[ElementID.ToString()][LOADCASEID].Add(rptCompDictkey, capture.Value);
                                                }
                                                rptValue = capture.Value;
                                            }
                                            j++;
                                        }
                                        else
                                        {
                                            if (Headers[j] == "Entity ID")
                                            {
                                                ElementID = int.Parse(capture.Value);
                                                if (!rptElemDict.ContainsKey(ElementID.ToString()))
                                                {
                                                    rptCompDict = new Dictionary<string, string>();
                                                    rptLoadCaseDict = new Dictionary<string, dynamic>();
                                                    rptCompDict.Add("LOADCASE", LoadCase);
                                                    rptCompDict.Add("DESCRIPTION", LCDescription);
                                                    rptLoadCaseDict.Add(LoadCase, rptCompDict);
                                                    rptElemDict.Add(ElementID.ToString(), rptLoadCaseDict);
                                                }
                                                else if (!rptElemDict[ElementID.ToString()].ContainsKey(LoadCase))
                                                {
                                                    rptCompDict = new Dictionary<string, string>();
                                                    rptLoadCaseDict = new Dictionary<string, dynamic>();
                                                    rptCompDict.Add("LOADCASE", LoadCase);
                                                    rptCompDict.Add("DESCRIPTION", LCDescription);
                                                    rptElemDict[ElementID.ToString()].Add(LoadCase, rptCompDict);
                                                }
                                            }
                                            else
                                            {
                                                if (!rptLoadCaseDict.ContainsKey(LoadCase))
                                                {
                                                    rptCompDict = new Dictionary<string, string>();
                                                    rptCompDict.Add("LOADCASE", LoadCase);
                                                    rptCompDict.Add("DESCRIPTION", LCDescription);
                                                    rptLoadCaseDict.Add(LoadCase, rptCompDict);
                                                }

                                                rptCompDictkey = GetRPTCompDictKey(Headers[j], ref ElemType, ref Layer, LoadType);

                                                if (rptElemDict[ElementID.ToString()][LoadCase].ContainsKey(rptCompDictkey))
                                                {
                                                    if (LoadType == "Rotational")
                                                    {
                                                        if (!rptElemDict[ElementID.ToString()](LoadCase).ContainsKey(rptCompDictkey + "P1"))
                                                            rptCompDictkey += "P1";
                                                        else if (!rptElemDict[ElementID.ToString()](LoadCase).ContainsKey(rptCompDictkey + "P2"))
                                                            rptCompDictkey += "P2";

                                                        rptElemDict[ElementID.ToString()](LoadCase).Add(rptCompDictkey, capture.Value);
                                                    }
                                                }
                                                else
                                                {
                                                    if (LoadType == "Rotational")
                                                    {
                                                        if (!rptElemDict[ElementID.ToString()](LoadCase).ContainsKey(rptCompDictkey + "P1"))
                                                            rptCompDictkey += "P1";
                                                        else if (!rptElemDict[ElementID.ToString()](LoadCase).ContainsKey(rptCompDictkey + "P2"))
                                                            rptCompDictkey += "P2";

                                                    }
                                                    rptElemDict[ElementID.ToString()](LoadCase).Add(rptCompDictkey, capture.Value);
                                                }
                                                rptValue = capture.Value;
                                            }
                                            j++;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        //Pass
                                    }
                                }  //End of foreach loop

                                // case 2 More than one entity IDs
                            }
                        }
                    } while (ainp.Peek() >= 0);    // End of do while
                } //end of using

                boolElemTypeMatch = true;
                boolregExHeader = false;
            }  //end of for each file loop
            sw.Stop();

            //Marshal.ReleaseComObject(wb);

            return rptElemDict;

        }

        /*private void LoadCaseDict(Match matches, ref Dictionary<int, dynamic> rptLCDict)
        {
            int LdCaseID;
            string LoadCase, LCDescription, LayerLC;

            LdCaseID = int.Parse(matches.Groups[1].Value);
            LoadCase = matches.Groups[2].Value;
            if (LoadCase.Contains(":"))
            {
                LoadCase = LoadCase.Replace(":", "");
            }
            LCDescription = matches.Groups[3].Value;
            LayerLC = matches.Groups[4].Value.Replace(" ", "");

            //boolregExLC = true;

            Dictionary<string, string>  LCParamDict = new Dictionary<string, string>();

            LCParamDict.Add("LOADCASE", LoadCase);
            LCParamDict.Add("DESCRIPTION", LCDescription);
            LCParamDict.Add("LCLAYER", LayerLC);

            rptLCDict.Add(LdCaseID, LCParamDict);

        }*/

        private string rptResultsPattern(ref List<string> Headers)
        {
            string PatternSourceID = @"\s+(\d+)";
            string PatternLoadcaseID = @"\s+(\d+)";
            string PatternSubcaseID = @"\s+(\d+)";
            string PatternEntityID = @"\s+(\d+)";
            string PatternElPosID = @"\s+(\d+)";
            string PatternXComponent = @"\s+(-?\d+\.\d+)";
            string PatternYComponent = @"\s+(-?\d+\.\d+)";
            string PatternZComponent = @"\s+(-?\d+\.\d+)";
            string PatternXYComponent = @"\s+(-?\d+\.\d+)";
            string PatternYZComponent = @"\s+(-?\d+\.\d+)";
            string PatternZXComponent = @"\s+(-?\d+\.\d+)";
            string PatternMaxPrincipal = @"\s+(-?\d+\.\d+)";
            string PatternMinPrincipal = @"\s+(-?\d+\.\d+)";
            string PatternMaxPrincipal2D = @"\s+(-?\d+\.\d+)";
            string PatternMinPrincipal2D = @"\s+(-?\d+\.\d+)";
            string PatternXLocation = @"\s+(-?\d+\.\d+)";
            string PatternYLocation = @"\s+(-?\d+\.\d+)";
            string PatternZLocation = @"\s+(-?\d+\.\d+)";
            string PatternvonMises = @"\s+(-?\d+\.\d+)";
            string PatternCID = @"\s+(\d+)";
            string PatternDefault = @"\s+(-?\d+\.\d+)";
            string PatternPropertyName = @"\s+(\S+\.\d+)";
            int i;
            string Stringval;
            string ResultsPattern = "^(?>";

            for (i = 0; i < Headers.Count; i++)
            {
                Stringval = Headers[i].ToString().Replace(" ", "");
                Stringval = Stringval.Replace(".", "");
                switch ("Pattern" + Stringval)
                {
                    case "PatternSourceID":
                        ResultsPattern += PatternSourceID;
                        break;
                    case "PatternLoadcaseID":
                        ResultsPattern += PatternLoadcaseID;
                        break;
                    case "PatternSubcaseID":
                        ResultsPattern += PatternSubcaseID;
                        break;
                    case "PatternEntityID":
                        ResultsPattern += PatternEntityID;
                        break;
                    case "PatternElPosID":
                        ResultsPattern += PatternElPosID;
                        break;
                    case "PatternXComponent":
                        ResultsPattern += PatternXComponent;
                        break;
                    case "PatternYComponent":
                        ResultsPattern += PatternYComponent;
                        break;
                    case "PatternZComponent":
                        ResultsPattern += PatternZComponent;
                        break;
                    case "PatternXYComponent":
                        ResultsPattern += PatternXYComponent;
                        break;
                    case "PatternYZComponent":
                        ResultsPattern += PatternYZComponent;
                        break;
                    case "PatternZXComponent":
                        ResultsPattern += PatternZXComponent;
                        break;
                    case "PatternMaxPrincipal":
                        ResultsPattern += PatternMaxPrincipal;
                        break;
                    case "PatternMinPrincipal":
                        ResultsPattern += PatternMinPrincipal;
                        break;
                    case "PatternMaxPrincipal2D":
                        ResultsPattern += PatternMaxPrincipal2D;
                        break;
                    case "PatternMinPrincipal2D":
                        ResultsPattern += PatternMinPrincipal2D;
                        break;
                    case "PatternXLocation":
                        ResultsPattern += PatternXLocation;
                        break;
                    case "PatternYLocation":
                        ResultsPattern += PatternYLocation;
                        break;
                    case "PatternZLocation":
                        ResultsPattern += PatternZLocation;
                        break;
                    case "PatternvonMises":
                        ResultsPattern += PatternvonMises;
                        break;
                    case "PatternCID":
                        ResultsPattern += PatternCID;
                        break;
                    case "PatternPropertyName":
                        ResultsPattern += PatternPropertyName;
                        break;
                    default:
                        ResultsPattern += PatternDefault;
                        break;
                }
            }
            return ResultsPattern;
        }

        private string AppendPattern(string rptResultsPattern, ref string PatternSourceID)
        {
            int PosIndex;
            string TempString = rptResultsPattern;
            int Increment = 1;
            if (rptResultsPattern.EndsWith("}"))
            {
                PosIndex = rptResultsPattern.LastIndexOf("{");
                int.TryParse(rptResultsPattern.Substring(PosIndex + 1, rptResultsPattern.Length - 2 - PosIndex), out Increment);
                TempString = rptResultsPattern.Substring(0, rptResultsPattern.Length - 2 - Increment.ToString().Length);
            }

            if (TempString.EndsWith(PatternSourceID))
            {
                rptResultsPattern = TempString + "{" + Increment + 1 + "}";
            }
            else
            {
                rptResultsPattern = rptResultsPattern + PatternSourceID;
            }

            return rptResultsPattern;
        }

        private string GetRPTCompDictKey(string Header, ref string ElemType, ref string Layer, string LoadType = "")
        {
            string RPTCompKey;
            string Stringval;

            if (!int.TryParse(Layer, out int result))
            {
                if (Layer.StartsWith("Z"))
                    Layer = Layer.Substring(Layer.Length - 2, 1);
                else
                    Layer = "";
            }


            Stringval = Header.Replace(" ", "");
            Stringval = Stringval.Replace(".", "");

            switch (Stringval)
            {
                case "SourceID":
                    RPTCompKey = "ID";  // "ID", "LOADCASEID", "LOADCASE", "POSID"
                    break;
                case "LoadcaseID":
                    RPTCompKey = "LOADCASEID";
                    break;
                case "SubcaseID":
                    RPTCompKey = "LOADCASEID";
                    break;
                case "EntityID":
                    RPTCompKey = "ELEMID";
                    break;
                case "ElPosID":
                    RPTCompKey = "POSID";
                    break;
                case "XComponent":
                case "XXComponent":
                    if (ElemType == "Bar" && LoadType == "Rotational")
                    {
                        Stringval = Stringval.Substring(0, 1) + "Rotation";

                    }
                    RPTCompKey = Stringval + Layer;
                    break;
                case "YComponent":
                case "YYComponent":
                    if (ElemType == "Bar" && LoadType == "Rotational")
                    {
                        Stringval = Stringval.Substring(0, 1) + "Rotation";
                    }
                    RPTCompKey = Stringval + Layer;
                    break;
                case "ZComponent":
                case "ZZComponent":
                    if (ElemType == "Bar" && LoadType == "Rotational")
                    {
                        Stringval = Stringval.Substring(0, 1) + "Rotation";
                    }
                    RPTCompKey = Stringval + Layer;
                    break;
                case "XYComponent":
                    RPTCompKey = Stringval + Layer;
                    break;
                case "MaxPrincipal":
                    RPTCompKey = Stringval + Layer;
                    break;
                case "MinPrincipal":
                    RPTCompKey = Stringval + Layer;
                    break;
                case "MaxPrincipal2D":
                    RPTCompKey = Stringval + Layer;
                    break;
                case "MinPrincipal2D":
                    RPTCompKey = Stringval + Layer;
                    break;
                default:
                    RPTCompKey = Stringval + Layer;
                    break;
            }
            return RPTCompKey;
        }

        private void PopulateResults(Dictionary<string, dynamic> elemDict, Excel.Range Rng, bool DescriptOpt)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            List<object> ResultData;

            Excel.Worksheet TblSheet = (Excel.Worksheet)wb.Worksheets[Rng.Parent.Name];

            TblSheet.Select();
            long RowNdx = Rng[1, 1].Row;
            long ColNdx = Rng[1, 1].Column;
            long StartCol = ColNdx;
            long StartRow = RowNdx;

            //string SheetName = Rng.Parent.Name;
            string CellName = Rng.Address.Replace("$", "");

            string[] TblHeaders = TableHeader(ref elemDict, DescriptOpt).Split(';');
            TblSheet.Range[TblSheet.Cells[RowNdx, ColNdx], TblSheet.Cells[RowNdx, ColNdx + TblHeaders.Length - 1]].Value = TblHeaders;

            RowNdx++;


            foreach (string ElemKey in elemDict.Keys)
            {
                foreach (string LCKey in elemDict[ElemKey].Keys)
                {
                    ResultData = new List<object>();
                    ResultData.Add(ElemKey);
                    foreach (string CompKey in elemDict[ElemKey][LCKey].Keys)
                    {
                        if (CompKey != "DESCRIPTION")
                        {
                            ResultData.Add(elemDict[ElemKey][LCKey][CompKey]);
                        }
                    }
                    if (elemDict[ElemKey][LCKey].ContainsKey("DESCRIPTION"))
                    {
                        ResultData.Add(elemDict[ElemKey][LCKey]["DESCRIPTION"]);
                    }

                    TblSheet.Range[TblSheet.Cells[RowNdx, ColNdx], TblSheet.Cells[RowNdx, ColNdx + ResultData.Count - 1]].Value = Array.ConvertAll(ResultData.ToArray(), s => double.TryParse(s.ToString(), out double xresult) ? xresult : s);
                    RowNdx++;
                }
            }

            long EndCol = StartCol - 1 + TblSheet.Range[TblSheet.Cells[StartRow, StartCol], TblSheet.Cells[StartRow, StartCol].End(Excel.XlDirection.xlToRight)].Count;
            long EndRow = StartRow - 1 + TblSheet.Range[TblSheet.Cells[StartRow, StartCol], TblSheet.Cells[StartRow, StartCol].End(Excel.XlDirection.xlDown)].Count;

            Excel.Range Selection = TblSheet.Range[TblSheet.Cells[StartRow, StartCol], TblSheet.Cells[EndRow, EndCol]]; //.Select

            //With Selection
            Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            Selection.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Selection.EntireColumn.AutoFit();

            Marshal.ReleaseComObject(wb);
        }

        public string TableHeader(ref Dictionary<string, dynamic> elemDict, bool DescriptOpt)
        {
            string ElemKey;
            string LCKey;
            //string CompKey;
            string Header;

            ElemKey = elemDict.Keys.First();
            LCKey = ((Dictionary<string, dynamic>)elemDict[ElemKey]).Keys.First();

            Header = "EntityID";
            foreach (string CompKey in elemDict[ElemKey][LCKey].Keys)
            {
                if (CompKey != "DESCRIPTION")
                    Header = Header + ";" + CompKey;
            }
            if (DescriptOpt)
            {
                Header = Header + ";" + "DESCRIPTION";
            }
            return Header;
        }

    }
}
