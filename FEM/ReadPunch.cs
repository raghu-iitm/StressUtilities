using StressUtilities;
using StressUtilities.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

/** 
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

namespace FEM
{
    class ReadPunch
    {
        public ReadPunch()
        {

        }

        private bool StatusFlag = false;
        //private bool OverallStatus = false;

        public void LaunchPunchForm()
        {
            IEnumerable<FormReadPunch> FrmCollection = Application.OpenForms.OfType<FormReadPunch>();

            if (FrmCollection.Any())
                FrmCollection.First().Focus();
            else
            {
                FormReadPunch F06form = new FormReadPunch();
                F06form.Show();
            }
        }

        public void ReadPunchResults(List<string> FileList, string ElemList, string Request, string EntityType)
        {
            Excel.Range Rng;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlapp = Globals.ThisAddIn.Application;
            //DialogResult Response;
            //List<string> Headings;
            //int i;
            //string ElemType;
            //General Common = new General();
            //string SearchSting = "/crdb/groups/group[@name='NASTRAN']/group[@name='RESULT']/group[@name='ELEMENTAL']/group[@name='STRESS']/dataset[@name='AXIF2_CPLX']";

            //List<long> ElementList;  // Changed from long to string

            List<long> ElementList = General.GetEntityList(ElemList);

            if (ElementList.Count == 0)
            {
                DialogResult Response = MessageBox.Show(@"Entity list is empty. All entities will be imported. Are you sure?", @"Warning!", MessageBoxButtons.YesNo);
                if (Response == DialogResult.No)
                    return;
            }

            try
            {
                Rng = wb.Application.InputBox("Select the Start Cell for populating the results.", "Obtain Range Object", Type: 8);
            }
            catch (Exception Ex)
            {
                //Rng = null;
                MessageBox.Show(@"Cancelled by the user. File not imported");
                return;
            }




            xlapp.Calculation = Excel.XlCalculation.xlCalculationManual;
            xlapp.ScreenUpdating = false;

            ImportPunchfiles(Rng, FileList, Request, ElementList, ref EntityType);

            ElementList = null;
            xlapp.StatusBar = false;
            xlapp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            xlapp.ScreenUpdating = true;

            Marshal.ReleaseComObject(wb);
            if (StatusFlag == true)
                MessageBox.Show(@"Nastran .pch File(s) Imported Successfully.");
            else
                MessageBox.Show(@"Could not import the results from Nastran .pch File(s). Entity ID is incorrect or the element type or requested type is not supported.");

            /*Headings = DataTypes(SearchSting);
            for (i = 0; i < Headings.Count; i++)
            {
                Console.WriteLine(Headings[i]);
                Console.ReadLine();
            }*/


            StatusFlag = false;
        }


        private void ImportPunchfiles(Excel.Range Rng, List<string> FileList, string Request, List<long> EntityList, ref string EntityType)
        {
            int FileCount = 0, Count; //LineCount
            //long FileSize;
            string TextLine, Title, SubTitle, Lable;
            object[] ResultData, DataSet2;
            bool BeginProcessing = false, NextLineStatus = false;

            string OutputType, SubCaseID = "";
            string RequestType = "", EigenValue, ElementType, OutputEntityType = "NODAL"; //PointID
            long ColNdx = 1;
            // -----Excel Parameters
            //long StartCol;
            Excel.Worksheet OutSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Application xlApp = Globals.ThisAddIn.Application;

            OutSheet = wb.ActiveSheet;
            long RowNdx = Rng.Row;
            long StartCol = Rng.Column;


            foreach (string punchfile in FileList)
            {
                using (StreamReader ainp = new StreamReader(punchfile))
                {
                    while (ainp.Peek() >= 0)
                    {
                        TextLine = ainp.ReadLine();

                        if (TextLine.StartsWith("$TITLE"))
                        {
                            BeginProcessing = false;
                            Title = TextLine.Substring(0, 72).Split('=')[1].Trim();
                            while (!TextLine.StartsWith(" "))
                            {
                                TextLine = ainp.ReadLine();
                                if (TextLine.StartsWith("$SUBTITLE"))
                                    SubTitle = TextLine.Substring(0, 72).Split('=')[1].Trim();
                                else if (TextLine.StartsWith("$LABEL"))
                                    Lable = TextLine.Substring(0, 72).Split('=')[1].Trim();
                                else if (TextLine.StartsWith("$" + Request))
                                {
                                    if (TextLine.Contains("="))
                                        RequestType = TextLine.Substring(0, 72).Split('=')[1].Trim();
                                    else
                                        RequestType = TextLine.Substring(0, 72).Trim();

                                    BeginProcessing = true;
                                }
                                else if (TextLine.StartsWith("$SUBCASE ID"))
                                    SubCaseID = TextLine.Substring(0, 72).Split('=')[1].Trim();
                                else if (TextLine.StartsWith("$RANDOM ID"))
                                    SubCaseID = TextLine.Substring(0, 72).Split('=')[1].Trim();
                                else if (TextLine.StartsWith("$POINT ID"))
                                    SubCaseID = TextLine.Substring(0, 72).Split('=')[1].Trim();
                                else if (TextLine.StartsWith("$ELEMENT TYPE"))
                                {
                                    ElementType = TextLine.Substring(0, 72).Split('=')[1].Trim();
                                    OutputEntityType = "ELEMENTAL";
                                }
                                else if (TextLine.StartsWith("$REAL OUTPUT"))
                                {
                                    if (TextLine.Contains("="))
                                        OutputType = TextLine.Substring(0, 72).Split('=')[1].Trim();
                                    else
                                        OutputType = TextLine.Substring(0, 72).Trim();
                                }
                                else if (TextLine.StartsWith("$EIGENVALUE"))
                                    EigenValue = TextLine.Substring(0, 72).Split('=')[1].Trim();
                            }
                        }

                        if (EntityType != OutputEntityType)
                            BeginProcessing = false;


                        if (BeginProcessing == true && !TextLine.StartsWith("$"))
                        {
                            Count = 0;
                            if (TextLine.StartsWith(" "))
                            {
                                NextLineStatus = false;

                                TextLine = TextLine.Substring(0, 72);
                                ResultData = TextLine.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);


                                if (EntityList.Count == 0 || EntityList.Contains(long.Parse(ResultData[0].ToString())))
                                {
                                    ColNdx = StartCol;
                                    RowNdx++;
                                    NextLineStatus = true;
                                    OutSheet.Cells[RowNdx, ColNdx].Value = SubCaseID;
                                    ColNdx++;
                                    OutSheet.Range[OutSheet.Cells[RowNdx, ColNdx], OutSheet.Cells[RowNdx, ColNdx + ResultData.Length - 1]].Value = Array.ConvertAll(ResultData, s => double.TryParse(s.ToString(), out double xresult) ? xresult : s);
                                    ColNdx += ResultData.Length;
                                    /*for (int i = 0; i < ResultData.Length; i++)
                                    {
                                        OutSheet.Cells[RowNdx, ColNdx].Value = ResultData[i];
                                        ColNdx++;
                                    }*/
                                    StatusFlag = true;
                                }

                                Count = ResultData.Length;
                            }
                            else if (TextLine.StartsWith("-CONT-") && NextLineStatus)
                            {
                                TextLine = TextLine.Substring(7, 65);  //72-7=65 Characters
                                DataSet2 = TextLine.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                                OutSheet.Range[OutSheet.Cells[RowNdx, ColNdx], OutSheet.Cells[RowNdx, ColNdx + DataSet2.Length - 1]].Value = Array.ConvertAll(DataSet2, s => double.TryParse(s.ToString(), out double xresult) ? xresult : s);
                                ColNdx += DataSet2.Length;
                                /*for (int i = 1; i < DataSet2.Length; i++)
                                {
                                    OutSheet.Cells[RowNdx, ColNdx].Value = DataSet2[i];
                                    ColNdx++;
                                }*/
                            }
                        }

                    }
                }
            }
        }

        //TODO - update this code to include the json instead of xml
       /* private List<string> DataTypes(string SearchSting)
        {
            string xmlFile = StressUtilities.Properties.Resources.DataType;
            XmlDocument xmldoc = new XmlDocument();


            xmldoc.LoadXml(xmlFile);
            XmlNode root = xmldoc.DocumentElement;

            List<string> Headings = new List<string>();
            int AttrCount;

            XmlNodeList nodeList = xmldoc.SelectNodes(SearchSting);
            XmlNode entitynode = xmldoc.SelectSingleNode(SearchSting);
            int count = entitynode.ChildNodes.Count;


            count = 0;
            foreach (XmlNode node in entitynode.ChildNodes)
            {
                AttrCount = node.Attributes.Count;
                if (node.Attributes[0].Value != "DOMAIN_ID")
                {
                    Headings.Add(node.Attributes[0].Value);
                    count++;
                }
            }

            return Headings;
        }*/



        public static string[] PunchRequestList()
        {
            string[] ResultList = new[] { "DISPLACEMENTS", "OLOADS", "SPCF", "ELEMENT FORCES", "ELEMENT STRESSES", "EIGENVALUE SUMMARY", "EIGENVECTOR", "GRID POINT SINGULARITY TABLE", "EIGENVALUE ANALYSIS SUMMARY", "VELOCITY", "ACCELERATION", "NON-LINEAR-FORCES", "GRID POINT WEIGHT OUTPUT", "EIGENVECTOR (SOLUTION SET)", "DISPLACEMENTS (SOLUTION SET)", "VELOCITY (SOLUTION SET)", "ACCELERATION (SOLUTION SET)", "ELEMENT STRAIN ENERGIES", "GRID POINT FORCE BALANCE", "STRESS AT GRID POINTS", "STRAIN/CURVATURE AT GRID POINTS", "ELEMENT INTERNAL FORCES And MOMENTS", "ELEMENT ORIENTED FORCES", "ELEMENT PRESSURES", "COMPOSITE FAILURE INDICIES", "GRID POINT STRESS/PLANE STRESS", "GRID POINT STRESS VOLUME DIRECT", "GRID POINT STRESS VOLUME PRINCIPAL", "ELEMENT STRESS DISCONTINUITIES", "ELEMENT STRESS DISCONTINUITIES DIRECT", "ELEMENT STRESS DISCONTINUITIES PRINCIPAL", "GRID POINT STRESS DISCONTINUITIES", "GRID POINT SRESS DISCONTINUITIES DIRECT", "GRID POINT STRESS DISCON PRINCIPAL", "GRID POINT STRESS/PLAIN STRAIN", "ELEMENT KINETIC ENERGY", "ELEMENT ENERGY LOSS PER CYCLE", "MAX/MIN SUMMARY INFORMATION", "MPCF", "MODAL GRID POINT KINETIC ENERGY", "TEMPERATURE", "HEAT FLOW AT LOAD POINTS", "HEAT FLOW AT CONTSTRAINT POINTS", "ELEMENT GRADIENTS And FLUXES", "ENTHALPY", "H DOT", "CROSS-PSDF", "CROSS-CORRELATION FUNCTION" };

            return ResultList;
        }



    }
}
