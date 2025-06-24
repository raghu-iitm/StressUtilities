using System.Windows.Forms;
//using RibbonHelp.core;

namespace StressUtilities
{
    public partial class ThisAddIn
    {
        private Ribbon ribbon;
        private Forms.SettingsControl SettingCntrl;
        private Forms.NastranCardsControl CardCntrl;
        private Forms.ReportControl ReportCntrl;
        private Forms.WriteCardControl WriteCardCntrl;

        private Microsoft.Office.Tools.CustomTaskPane UtilityCustomTaskPane;
        //private RibbonHelpContext ribbonHelp;

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return UtilityCustomTaskPane;
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new Ribbon();
            return ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SettingCntrl = new Forms.SettingsControl();
            CardCntrl = new Forms.NastranCardsControl();
            ReportCntrl = new Forms.ReportControl();
            WriteCardCntrl = new Forms.WriteCardControl();

            UtilityCustomTaskPane = this.CustomTaskPanes.Add(SettingCntrl, "Settings");
            UtilityCustomTaskPane = this.CustomTaskPanes.Add(CardCntrl, "Nastran Cards");
            UtilityCustomTaskPane = this.CustomTaskPanes.Add(ReportCntrl, "Report");
            UtilityCustomTaskPane = this.CustomTaskPanes.Add(WriteCardCntrl, "Write Nastran Cards");
            
            //UtilityCustomTaskPane.Width = 310;
            UtilityCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            
            //TaskPane.Width = UtilityCustomTaskPane.Width;

            SettingCntrl.AutoScaleMode = AutoScaleMode.Inherit;
            CardCntrl.AutoScaleMode = AutoScaleMode.Inherit;
            ReportCntrl.AutoScaleMode = AutoScaleMode.Inherit;
            WriteCardCntrl.AutoScaleMode = AutoScaleMode.Inherit;
            
            UtilityCustomTaskPane.VisibleChanged += new System.EventHandler(UtilityCustomTaskPane_VisibleChanged);

        }

        private void UtilityCustomTaskPane_VisibleChanged(object sender, System.EventArgs e)
        {
            ribbon.refresh();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            SettingCntrl.Dispose();
            CardCntrl.Dispose();
            ReportCntrl.Dispose();
            WriteCardCntrl.Dispose();
            UtilityCustomTaskPane.Dispose();
            
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
