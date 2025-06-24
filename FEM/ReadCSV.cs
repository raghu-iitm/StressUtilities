using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace FEM
{
    class ReadCSV
    {
        public ReadCSV()
        {
        }

        public Dictionary<string, dynamic> importCSVfiles(Excel.Range Rng, string[] cvsFileList, string MapFileName)
        {
            string[] ResultData;
            //int i;
            StreamReader ainp;
            Dictionary<string, string> MapDict = new Dictionary<string, string>();
            Dictionary<string, string> LCParamDict = new Dictionary<string, string>();
            Dictionary<string, string> LCSortDict = new Dictionary<string, string>();
            Dictionary<int, string> csvCompIndex = new Dictionary<int, string>();
            Dictionary<int, dynamic> csvLCDict = new Dictionary<int, dynamic>();
            Dictionary<string, string> csvCompDict = new Dictionary<string, string>();
            Dictionary<string, dynamic> csvLoadCaseDict = new Dictionary<string, dynamic>();
            Dictionary<string, dynamic> csvElemDict = new Dictionary<string, dynamic>();
            int FileCount = 0;
            int LineCount;
            bool CSVinfoComplete;
            string ElemID = "";
            string SCID = "";
            bool ChkHeader;
            //int placeholder = 0;
            string listSeparator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;

            if (File.Exists(MapFileName))
            {
                MapDict = ParamMap(MapFileName);
                if (MapDict.ContainsKey("LSEP"))
                    listSeparator = MapDict["LSEP"];
            }

            foreach (string cvsFile in cvsFileList)
            {
                CSVinfoComplete = false;
                FileCount++;
                ainp = new System.IO.StreamReader(cvsFile);
                LineCount = 0;
                csvLCDict = new Dictionary<int, dynamic>();
                csvLoadCaseDict = new Dictionary<string, dynamic>();

                ChkHeader = true;

                while (ainp.Peek() != -1)
                {
                    ResultData = ainp.ReadLine().Split(listSeparator.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                    if (ChkHeader == true)
                    {
                        csvCompIndex = new Dictionary<int, string>();
                        for (int i = 0; i <= ResultData.Length - 1; i++)
                        {
                            ResultData[i] = ResultData[i].Trim();
                            ResultData[i] = ResultData[i].Replace(" ", "");

                            if (MapDict.ContainsKey(ResultData[i]))
                            {
                                if (MapDict[ResultData[i]] == "EID" || MapDict[ResultData[i]] == "ID" || MapDict[ResultData[i]] == "SUBCASE")
                                    ResultData[i] = MapDict[ResultData[i]];
                            }
                            csvCompIndex.Add(i, ResultData[i]);
                        }
                        if (ResultData[0] == "EID" || ResultData[0] == "ID")
                            ChkHeader = false;
                    }

                    if (ChkHeader == false && double.TryParse(ResultData[0], out double _))
                    {
                        try
                        {
                            csvCompDict = new Dictionary<string, string>();
                            for (int i = 0; i <= ResultData.Length - 1; i++)
                            {
                                if (!(csvCompIndex[i] == "EID" || csvCompIndex[i] == "ID" || csvCompIndex[i] == "SUBCASE"))
                                    csvCompDict.Add(csvCompIndex[i], ResultData[i]);
                                else if (csvCompIndex[i] == "EID" || csvCompIndex[i] == "ID")
                                    ElemID = ResultData[i];
                                else if (csvCompIndex[i] == "SUBCASE")
                                {
                                    SCID = ResultData[i];
                                    csvCompDict.Add(csvCompIndex[i], SCID);
                                }
                            }

                            if (!csvElemDict.ContainsKey(ElemID))
                            {
                                csvLoadCaseDict = new Dictionary<string, object>();
                                // csvCompDict.Add("SUBCASE", SCID)
                                csvLoadCaseDict.Add(SCID, csvCompDict);
                                csvElemDict.Add(ElemID, csvLoadCaseDict);
                            }
                            else if (!csvElemDict[ElemID].ContainsKey(SCID))
                            {
                                csvLoadCaseDict = new Dictionary<string, object>();
                                // csvLoadCaseDict.Add(csvCompDict("SUBCASE"), csvCompDict)
                                csvElemDict[ElemID].Add(SCID, csvCompDict);
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
                ainp.Close();
            }
            return csvElemDict;
        }


        public Dictionary<string, string> ParamMap(string MapFileName)
        {
            string MapData;
            string[] MapArray;
            string[] MapKeys;
            string Key;
            string Value;
            Dictionary<string, string> Params = new Dictionary<string, string>();
            //StreamReader ainp;

            StreamReader ainp = new System.IO.StreamReader(MapFileName);
            while (ainp.Peek() != -1)
            {
                MapData = ainp.ReadLine();
                if (!string.IsNullOrWhiteSpace(MapData))
                {
                    if (MapData.Substring(0, 1) != "#")
                    {
                        MapArray = MapData.Split('|');
                        if (MapArray[1].Contains(","))
                        {
                            MapKeys = MapArray[1].Split(',');
                            for (int i = 0; i < MapKeys.Length; i++)
                            {
                                Key = MapKeys[i].Trim();
                                Key = Key.Replace(" ", "");
                                Value = MapArray[0].Trim();
                                if (Key != "")
                                    Params.Add(Key, Value);
                            }
                        }
                        else
                        {
                            Key = MapArray[1].Trim();
                            Key = Key.Replace(" ", "");
                            Value = MapArray[0].Trim();
                            if (Key != "")
                            {
                                if (Value == "LSEP")
                                    Params.Add(Value, Key);
                                else
                                    Params.Add(Key, Value);
                            }
                        }
                    }
                }
            }
            ainp.Close();

            return Params;
        }

    }
}
