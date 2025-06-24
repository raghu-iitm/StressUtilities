using System;
using StressUtilities.FEM;
using System.Windows.Forms;

namespace StressUtilities.Forms
{
     
    public partial class NastranCardsControl : UserControl
    {
        private int bulkEntryBoxSelIndx = 0;
        private int loadTypeBoxSelIndx = 0;
        //private int Cmbbox3SelIndx = 0;
        public NastranCardsControl()
        {
            InitializeComponent();
            this.Load += new EventHandler(NastranCards_Load);
        }

     

        private void BtnApply_Click(object sender, EventArgs e)
        {
            NastranCards NasCard = new NastranCards();
            if (nastranCardsBox.Text!=null)
                NasCard.WriteNastranCards(nastranCardsBox.Text);
        }


        private void NastranCards_Load(object sender, EventArgs e)
        {
            string[] BulkEntryType = { "Load Cards", "SUBCASE" };
            string[] LoadTypes = { "Point Loads", "Distributed Loads", "Inertia Loads", "Thermal Loads", "Enforced Motion", 
                                    "Element Deformation", "Frequency-Dependent Loads", "Time-Dependent Loads", "Load Combination" };
            bulkEntryBox.Items.AddRange(BulkEntryType);
            loadTypeBox.Items.AddRange(LoadTypes);
            bulkEntryBox.SelectedIndex = 0;
            loadTypeBox.SelectedIndex = 0;
            nastranCardsBox.SelectedIndex = 0;

        }

        private void blkEntryBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (bulkEntryBox.SelectedIndex == 1)
            {
                loadTypeBox.Enabled = false;
                label2.Enabled = false;
                nastranCardsBox.Items.Clear();
                nastranCardsBox.Items.Add("SUBCASE");
                nastranCardsBox.SelectedIndex = 0;
            }
            else
            {
                loadTypeBox.Enabled = true;
                label2.Enabled = true;
            }
            bulkEntryBoxSelIndx = bulkEntryBox.SelectedIndex;

        }

        private void loadTypeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] PointLoads = { "FORCE", "FORCE1", "FORCE2", "MOMENT", "MOMENT1", "MOMENT2", "SLOAD" };
            string[] DistributedLoads = { "PLOAD", "PLOAD1", "PLOAD2", "PLOADB3", "PLOAD4", "PRESAX", "PLOADX1" }; 
            string[] InertiaLoads = { "GRAV", "RFORCE", "ACCEL", "ACCEL1" };
            string[] ThermalLoads = { "TEMP", "TEMPD", "TEMPAX", "TEMPBC" };
            string[] EnforcedMotion = { "SPC", "SPC1", "SPCD", "SPCAX" };
            string[] ElementDeformation = { "DEFORM" };
            string[] FrequencyDependentLoads = { "DAREA", "RLOAD1", "RLOAD2" };
            string[] TimeDependentLoads = { "DAREA", "DLOAD", "TLOAD1", "TLOAD2" };
            string[] LoadCombination = { "LOAD", "LOADT" };

            loadTypeBoxSelIndx = loadTypeBox.SelectedIndex;
            nastranCardsBox.Items.Clear();
            switch (loadTypeBox.Text)
            {
                case "Point Loads":                    
                    nastranCardsBox.Items.AddRange( PointLoads);
                    break;
                case "Distributed Loads":
                    nastranCardsBox.Items.AddRange(DistributedLoads);
                    break;
                case "Inertia Loads":
                    nastranCardsBox.Items.AddRange(InertiaLoads);
                    break;
                case "Thermal Loads":
                    nastranCardsBox.Items.AddRange(ThermalLoads);
                    break;
                case "Enforced Motion":
                    nastranCardsBox.Items.AddRange(EnforcedMotion);
                    break;
                case "Element Deformation":
                    nastranCardsBox.Items.AddRange(ElementDeformation);
                    break;
                case "Frequency-Dependent Loads":
                    nastranCardsBox.Items.AddRange(FrequencyDependentLoads);
                    break;
                case "Time-Dependent Loads":
                    nastranCardsBox.Items.AddRange(TimeDependentLoads);
                    break;
                case "Load Combination":
                    nastranCardsBox.Items.AddRange(LoadCombination);
                    break;
            }
            nastranCardsBox.SelectedIndex = 0;
        }

        //private void nastranCardsBox_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //        Cmbbox3SelIndx = comboBox3.SelectedIndex;
        //}
    }
}
