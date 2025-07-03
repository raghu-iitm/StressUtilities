using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/** 
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

namespace StressUtilities
{
    public enum SheetName
    {
        SHEET_NAME_REPORT = 1,
        SHEET_NAME_INPUT = 2,
        SHEET_NAME_REFERENCE = 3
    }
    class General
    {

        //private const int CONST_TOTAL_LOAD_TYPES = 4;

        public static string GetSheetName(int eNumValue)
        {
            string[] SheetNames = new[] { "ReportOptions", "Input", "Reference" };
            return SheetNames[eNumValue - 1];
        }

        public static string GetFolderpath()
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            return wb.Path;
        }


        public bool IsFileOpen(FileInfo file)
        {
            //FileStream myStream = null;
            bool FileStatus = false;
            try
            {
                FileStream myStream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                myStream.Close();
            }
            catch (Exception ex)
            {
                if (!file.Exists)
                    FileStatus = false;
                else
                    FileStatus = true;
            }
            return FileStatus;
        }


        public static List<long> GetEntityList(string EntityList)
        {
            //string[]  ExpandedList;
            long increment = 1;
            List<long> ListFinal = new List<long>();

            if (string.IsNullOrEmpty(EntityList))
                return new List<long>();

            string[] listInter = EntityList.Split(new char[] { ';', ',', '\n', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            //long entity = 0;

            foreach (string listItem in listInter)
            {
                if (listItem.Contains(":"))
                {
                    
                    string[] ExpandedList = listItem.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                    long.TryParse(ExpandedList[0], out long start);
                    long.TryParse(ExpandedList[1], out long last);

                    if (ExpandedList.Length >= 3)
                        long.TryParse(ExpandedList[2], out increment);

                    for (long i = start; i <= last; i += increment)
                    {
                        ListFinal.Add(i);
                    }
                }
                else if (long.TryParse(listItem, out long _))
                    ListFinal.Add(long.Parse(listItem));  //listItem previously
            }
            return ListFinal;
            //return ListFinal.Distinct().OrderBy(x => x).ToList();
        }

        public static List<string> GetUnitThermalRanges(ref string UnitThermalLoads)
        {
            //string[] listInter;
            List<string> ListFinal = new List<string>();
           
            Excel.Application xlapp = Globals.ThisAddIn.Application;

            if (string.IsNullOrEmpty(UnitThermalLoads))
                return new List<string>();

            string[] listInter = UnitThermalLoads.Split(new char[] { ';', ',', '\n', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            //long entity = 0;

            foreach (string listItem in listInter)
            {
                if (listItem.Contains(":"))
                {
                    //Excel.Range CellsRng;
                    Excel.Range CellsRng = xlapp.Range[listItem];
                    foreach (Range cell in CellsRng)
                        ListFinal.Add(cell.Address[false, false, XlReferenceStyle.xlA1, false, false].ToString());
                }
                else
                    ListFinal.Add(listItem);
            }
            return ListFinal;
        }

        public static string BrowseFolder()
        {
            string MapFilePath="";
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = wb.Path,
                Filter = string.Empty,
                Multiselect = false,
                RestoreDirectory = true,
                Title = "Select Folder...",
                CheckFileExists = false,
                CheckPathExists = false,
                FileName = "dummy"
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    MapFilePath = openFileDialog1.FileName.Replace(openFileDialog1.SafeFileName, "");

                    
                }
                catch (Exception Ex)
                {
                    MessageBox.Show("Cannot read file from disk. Original error: " + Ex.Message);
                    return null;
                }

            }
            return MapFilePath;
        }


    }
}
