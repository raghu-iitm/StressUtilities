﻿using FEM;
using Nastranh5;
using Report;
using StressUtilities.FEM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

/*
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/


namespace StressUtilities
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        internal static Office.IRibbonUI StressUtilityRibbon;
        private bool LicenseStatus = true; //change to false

        public Ribbon()
        {
            StressUtilityRibbon = this.ribbon;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("StressUtilities.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Ribbon Actions
        public void RibbonActions(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "BtnInsrptFile":
                    Readrpt ReadFile = new Readrpt();
                    ReadFile.ImportPatranRPTfile();
                    break;
                case "Btnf06":
                    Readf06 fo6frm = new Readf06();
                    fo6frm.LaunchF06Form();
                    break;
                case "BtnReadPunch":
                    ReadPunch PunchRead = new ReadPunch();
                    PunchRead.LaunchPunchForm();
                    break;
                case "BtnReadHDF5":
                    H5DBread HDF5Form = new H5DBread();
                    HDF5Form.LaunchHDF5Form();
                    break;
                case "BtnExptLCTbl":
                    LCTable LCCombination = new LCTable();
                    LCCombination.LCTableTemplate();
                    break;
                case "BtnCombLoadCase":
                    LoadCombination LComb = new LoadCombination();
                    LComb.LaunchCombiForm();
                    break;
                case "BtnTbleVert":
                    WriteReport Report = new WriteReport();
                    Report.AddCustomTableSingle();
                    break;
                case "BtnInsTbl":
                    WriteReport ReportTbl = new WriteReport();
                    ReportTbl.AddCustomTable();
                    break;
                case "BtnTblRename":
                    WriteReport TblRename = new WriteReport();
                    TblRename.AutoTableNames();
                    break;
                case "BtnInsRef":
                    Reference TblRef = new Reference();
                    TblRef.InsertRefTable();
                    break;
                case "BtnPrepareReport":
                    WriteReport ReportForm = new WriteReport();
                    ReportForm.LaunchReportContentsForm();
                    break;
                case "BtnRefresh":
                    {
                        Microsoft.Office.Interop.Excel.Application xlApp = Globals.ThisAddIn.Application;
                        xlApp.Application.ScreenUpdating = true;
                        xlApp.Application.StatusBar = "";
                    }
                    break;
                //case "BtnReport":
                //    WriteReport LaunchReportForm = new WriteReport();
                //    LaunchReportForm.LaunchReportForm();
                //break;
                //case "BtnNastranCards":
                //    NastranCards NCards = new NastranCards();
                //    NCards.WriteNastranCards();
                //    break;
                //case "BtnWriteCards":
                //    NastranCards NCards2File = new NastranCards();
                //    NCards2File.WriteNastranCards();
                //    break;
                case "BtnAbout":
                    IEnumerable<AboutStrUtilities> FrmCollection = System.Windows.Forms.Application.OpenForms.OfType<AboutStrUtilities>();
                    if (FrmCollection.Any())
                    {
                        FrmCollection.First().Focus();
                    }
                    else
                    {
                        AboutStrUtilities frmAbout = new AboutStrUtilities();
                        frmAbout.ShowDialog();
                    }
                    break;
                case "BtnHelp":
                    string AddInLocation = Assembly.GetExecutingAssembly().CodeBase.ToString();
                    View_Help(AddInLocation);
                    break;
                /*case "BtnLicense":
                    LicenseValidation license = new LicenseValidation();
                    license.ImportLicense();

                    this.ribbon.Invalidate();
                    break;*/

            }
        }

        private void BtnSettings_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SettingsPane.Visible = ((Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender).Checked;
        }

        private void View_Help(string asmLocation)
        {
            string tempName = "UserGuideStressUtilities.chm";
            string assyname = "StressUtilities.DLL";

            string location = asmLocation.Substring(asmLocation.Length-(asmLocation.Length - 8), asmLocation.Length - 8);  //In order to remove file: ///

            location = location.Substring(0,location.Length - assyname.Length);
            location = location.Replace(@"/", @"\");

            System.Diagnostics.Process.Start(location + tempName);

        }

        public bool BtnSettingsPressed(Office.IRibbonControl control) 
        { 
            return Globals.ThisAddIn.CustomTaskPanes[0].Visible; 
        }

        public void BtnSettingsToggle(Office.IRibbonControl control, bool value) 
        { 
            Globals.ThisAddIn.CustomTaskPanes[0].Visible = value; 
        }

        public bool BtnNastranCardsPressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.CustomTaskPanes[1].Visible;
        }

        public void BtnNastranCardsToggle(Office.IRibbonControl control, bool value)
        {
            Globals.ThisAddIn.CustomTaskPanes[1].Visible = value;
        }

        public bool BtnReportPressed(Office.IRibbonControl control)
        {

            return Globals.ThisAddIn.CustomTaskPanes[2].Visible;
        }

        public void BtnReportToggle(Office.IRibbonControl control, bool value)
        {
            Microsoft.Office.Interop.Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            try
            {
                DateTime LastSavedTime = (DateTime)wb.BuiltinDocumentProperties.item["Last Save Time"].value;
                Globals.ThisAddIn.CustomTaskPanes[2].Visible = value;
            }
            catch (Exception)
            {
                MessageBox.Show(@"The workbook is not saved. Please save the workbook before proceeding.");
            }
        }

        public bool BtnWriteCardsPressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.CustomTaskPanes[3].Visible;
        }

        public void BtnWriteCardsToggle(Office.IRibbonControl control, bool value)
        {
            Globals.ThisAddIn.CustomTaskPanes[3].Visible = value;
        }

        public void refresh()
        {
            ribbon.InvalidateControl("BtnSettings");
            ribbon.InvalidateControl("BtnNastranCards");
            ribbon.InvalidateControl("BtnReport");
            ribbon.InvalidateControl("BtnWriteCards");
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
