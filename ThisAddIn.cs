using Microsoft.Office.Core;
using System.Collections.Generic;
using System.Configuration;
using System.Windows.Forms;

/*
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

namespace StressUtilities
{
    public partial class ThisAddIn
    {
        private Ribbon ribbon;
        private Forms.SettingsControl SettingCntrl;
        private Forms.NastranCardsControl CardCntrl;
        private Forms.ReportControl ReportCntrl;
        private Forms.WriteCardControl WriteCardCntrl;

        //private Microsoft.Office.Tools.CustomTaskPane UtilityCustomTaskPane;
        //private RibbonHelpContext ribbonHelp;

        /*private Microsoft.Office.Tools.CustomTaskPane SettingsPane;
        private Microsoft.Office.Tools.CustomTaskPane CardsPane;
        private Microsoft.Office.Tools.CustomTaskPane ReportPane;
        private Microsoft.Office.Tools.CustomTaskPane WriteCardsPane;*/

        public Microsoft.Office.Tools.CustomTaskPane SettingsPane { get; private set; }
        public Microsoft.Office.Tools.CustomTaskPane CardsPane { get; private set; }
        public Microsoft.Office.Tools.CustomTaskPane ReportPane { get; private set; }
        public Microsoft.Office.Tools.CustomTaskPane WriteCardsPane { get; private set; }

        private Dictionary<string, CustomTaskPane> taskPanes = new Dictionary<string, CustomTaskPane>();

        public CustomTaskPane GetTaskPane(string key)
        {
            return taskPanes.TryGetValue(key, out var pane) ? pane : null;
        }

        /*public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return SettingsPane;
            }
        }*/

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

            SettingsPane = this.CustomTaskPanes.Add(SettingCntrl, "Settings");
            CardsPane = this.CustomTaskPanes.Add(CardCntrl, "Nastran Cards");
            ReportPane = this.CustomTaskPanes.Add(ReportCntrl, "Report");
            WriteCardsPane = this.CustomTaskPanes.Add(WriteCardCntrl, "Write Nastran Cards");

            /* UtilityCustomTaskPane = this.CustomTaskPanes.Add(SettingCntrl, "Settings");
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

             UtilityCustomTaskPane.VisibleChanged += new System.EventHandler(UtilityCustomTaskPane_VisibleChanged);*/

            SettingsPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            CardsPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            ReportPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            WriteCardsPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;

            SettingsPane.VisibleChanged += UtilityCustomTaskPane_VisibleChanged;
            CardsPane.VisibleChanged += UtilityCustomTaskPane_VisibleChanged;
            ReportPane.VisibleChanged += UtilityCustomTaskPane_VisibleChanged;
            WriteCardsPane.VisibleChanged += UtilityCustomTaskPane_VisibleChanged;

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
            //UtilityCustomTaskPane.Dispose();
            
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
