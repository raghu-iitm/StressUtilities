using System;
using System.Windows.Forms;

namespace StressUtilities.Forms
{
    public partial class SettingsControl : UserControl
    {
        public SettingsControl()
        {
            InitializeComponent();
            this.Load += new EventHandler(SettingsControl_Load);
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.MaxHDFRows = long.Parse(HDF5RowBox.Text);
            Properties.Settings.Default.WorkingDirectory= WorkingDirectoryBox.Text;
            Properties.Settings.Default.OptionUnits = unitOptionBox.Checked;
        }

        private void BrowseBox_Click(object sender, EventArgs e)
        {
            this.WorkingDirectoryBox.Text = General.BrowseFolder();
        }

        private void SettingsControl_Load(object sender, EventArgs e)
        {
            HDF5RowBox.Text= Properties.Settings.Default.MaxHDFRows.ToString();
            WorkingDirectoryBox.Text=Properties.Settings.Default.WorkingDirectory;
            unitOptionBox.Checked = Properties.Settings.Default.OptionUnits;
        }

    }
}
