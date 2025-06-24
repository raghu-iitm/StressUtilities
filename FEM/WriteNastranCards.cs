using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace StressUtilities.FEM
{
    public class WriteNastranCards
    {
        private string StartCell = "B3";
        public WriteNastranCards()
        {
        }

        public void WriteCards(List<string> CardList, string FileName, string TargetOS, string CardFormat)
        {
            
            string eol; // //Properties.Settings.Default.TargetOS;
            switch (TargetOS)
            {
                case "Windows":
                    eol = "\r\n";
                    break;
                case "Unix/Linux":
                    eol = "\n";
                    break;
                default:
                    eol = "\n";
                    break;
            }

            StreamWriter sw = new StreamWriter(FileName);
            if (CardList.Contains("SUBCASE"))
            {
                WriteSubcaseToTextFile("SUBCASE", sw, eol);
            }
            foreach(string Card in CardList)
            {
                if (Card!="SUBCASE")
                {
                    WriteToTextFile(Card, sw, eol, CardFormat);
                }

            }
            sw.Close();
            MessageBox.Show(string.Format("Nastran Cards have been written to the File \n {0}",FileName));
        }

        private void WriteToTextFile(string Card, StreamWriter sw, string eol, string CardFormat)
        {
            int Incr;
            long RowNdx, ColNdx;
            string CellValue, cardstring="";
            bool FirstLine;
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            Range Selection = wb.Worksheets[Card].Range[StartCell];
            RowNdx = Selection.Row;
            ColNdx = Selection.Column;
            int ColCount = wb.Worksheets[Card].Range[Selection, Selection.End[XlDirection.xlToRight]].Count();
            int RemainingFieldCount = ColCount;

            string defaultLanguage = System.Globalization.CultureInfo.CurrentCulture.ToString();
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            
            StringBuilder outString = new StringBuilder();
            //Write Results here
            switch (CardFormat)
            {
                case "SMALL":
                    outString.Append("$-------1-------2-------3-------4-------5-------6-------7-------8-------9-------" + eol);
                    break;
                case "LARGE":
                    outString.Append("$-------1---------------2---------------3---------------4---------------5-------" + eol);
                    break;
                case "FREE":
                    outString.Append("$-------1-------2-------3-------4-------5-------6-------7-------8-------9-------" + eol);
                    break;
            }

            while (!string.IsNullOrEmpty(wb.Worksheets[Card].Cells[RowNdx+1,ColNdx].Value))
            {
                Incr = 10;
                FirstLine = true;
                for (int i = 0; i < ColCount; i++)
                {
                    if (string.IsNullOrEmpty(wb.Worksheets[Card].Cells[RowNdx + 1, ColNdx + i].Text))
                        CellValue = "";
                    else
                        CellValue = wb.Worksheets[Card].Cells[RowNdx + 1, ColNdx + i].Text.ToString();

                    CellValue=realnumber(CellValue);

                    switch (CardFormat)
                    {
                        case "SMALL":
                            outString.Append(string.Format("{0,-8}", CellValue));
                            if ((i + 1) % 9 == 0)
                            {
                                if (ColCount - i > 1)
                                    outString.Append(eol + "        ");
                                else
                                    outString.Append(eol);
                            }
                            else if (i + 1 == ColCount)
                            {
                                outString.Append(eol);
                            }
                            break;
                        case "LARGE":
                            if(i==0)
                            {
                                //CellValue += "*";
                                outString.Append(string.Format("{0,-8}", CellValue + "*"));
                                cardstring = CellValue.Substring(0,4);
                            }
                            else if ((i + 1) % 6 == 0 && FirstLine)
                            {
                                //CellValue += "*";
                                outString.Append(string.Format("{0,-8}", "")); //"*"+ cardstring + Incr
                                Incr++;
                                FirstLine = false;
                            }
                            else if ((i-5) % 5 == 0 && !FirstLine)
                            {
                                //CellValue += "*";
                                outString.Append(string.Format("{0,-8}", "")); //"*" + cardstring + Incr
                                Incr++;
                                //FirstLine = false;
                            }
                            else 
                            {
                                outString.Append(string.Format("{0,-16}", CellValue));
                            }
                                
                            if ((i + 1) % 6 == 0 && FirstLine)
                            {
                                if (ColCount - i > 1)
                                {
                                    outString.Append(string.Format("{0}{1,-8}", eol, "*" )); //+ cardstring + Incr
                                    Incr++;
                                }
                                else
                                    outString.Append(eol);
                            }
                            else if ((i - 5) % 5 == 0 && !FirstLine)
                            {
                                if (ColCount - i > 1)
                                {
                                    outString.Append(string.Format("{0}{1,-8}", eol, "*" ));//+ cardstring + Incr
                                    Incr++;
                                }
                                else
                                    outString.Append(eol);
                            }
                            else if (i + 1 == ColCount)
                            {
                                outString.Append(string.Format("{0,-8}{1}", "", eol)); //"*" + cardstring + Incr
                            }

                            break;
                        case "FREE":
                            outString.Append(string.Format("{0}", CellValue));
                            if ((i + 1) % 9 == 0)
                            {
                                if (ColCount - i > 1)
                                    outString.Append(eol + ",");
                                else
                                    outString.Append(eol);
                            }
                            else if (i + 1 == ColCount)
                            {
                                outString.Append(eol);
                            }
                            else
                            {
                                outString.Append(",");
                            }
                            break;

                    }
                }
                RowNdx++;
            }
            sw.Write(outString.ToString(),Encoding.ASCII);


            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo(defaultLanguage);
            outString.Clear();
            
        }


        private void WriteSubcaseToTextFile(string Card, StreamWriter sw, string eol)
        {
            long RowNdx, ColNdx, StartRow;
            string CellValue, Param, TrimmedString;
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            Range Selection = wb.Worksheets[Card].Range[StartCell];
            RowNdx = Selection.Row;
            StartRow = RowNdx;
            ColNdx = Selection.Column;
            int ColCount = wb.Worksheets[Card].Range[Selection, Selection.End[XlDirection.xlToRight]].Count();

            StringBuilder outString = new StringBuilder();
            //Write Results here

            while (!string.IsNullOrEmpty(wb.Worksheets[Card].Cells[RowNdx + 1, ColNdx].Text))
            {
                for (int i = 0; i < ColCount; i++)
                {
                    Param = wb.Worksheets[Card].Cells[StartRow, ColNdx + i].Text.ToString();
                    CellValue = wb.Worksheets[Card].Cells[RowNdx + 1, ColNdx + i].Text.ToString();
                    if (!string.IsNullOrEmpty(CellValue))
                    {
                        if (Param == "SUBCASE")
                        {
                            TrimmedString = string.Format("{0} {1}", Param, CellValue);
                            if (TrimmedString.Length > 72)
                                TrimmedString = TrimmedString.Substring(0, 72);
                            outString.Append(string.Format("{0}{1}", TrimmedString, eol));
                        }
                        else
                        {
                            TrimmedString = string.Format("    {0} = {1}", Param, CellValue);
                            if (TrimmedString.Length > 72)
                                TrimmedString = TrimmedString.Substring(0, 72);
                            outString.Append(string.Format("{0}{1}", TrimmedString, eol));
                        }
                    }
                }
                RowNdx++;
            }
            sw.Write(outString.ToString(), Encoding.ASCII);
            outString.Clear();
        }

        // the code needs to be updated.
        private string realnumber(string CellValue)
        {
            bool ValueCheck=double.TryParse(CellValue, out double result);
            if(ValueCheck)
            {
                if (CellValue.Contains("E"))
                {
                    string[] SplitValue = CellValue.Split('E');
                    if (!SplitValue[0].Contains("."))
                    {
                        SplitValue[0] += ".";
                    }
                    //return CellValue;
                    if (SplitValue[1].Contains("+0"))
                    {
                        SplitValue[1] = SplitValue[1].Replace("+0","+");
                    }
                    else if (SplitValue[1].Contains("-0"))
                    {
                        SplitValue[1] = SplitValue[1].Replace("-0", "-");
                    }
                    CellValue = SplitValue[0] + SplitValue[1];
                }

            }
            return CellValue;
        }

    }
}
