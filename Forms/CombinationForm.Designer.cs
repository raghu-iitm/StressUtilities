namespace StressUtilities.Forms
{
    partial class CombinationForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CombinationForm));
            this.DataSourceh5 = new System.Windows.Forms.RadioButton();
            this.TextBoxMapFile = new System.Windows.Forms.TextBox();
            this.DataSourceCSV = new System.Windows.Forms.RadioButton();
            this.DataSourceRPT = new System.Windows.Forms.RadioButton();
            this.Label7 = new System.Windows.Forms.Label();
            this.FileListBox = new System.Windows.Forms.ListBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnCombLoads = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.ElemList = new System.Windows.Forms.TextBox();
            this.LCFileBox = new System.Windows.Forms.TextBox();
            this.unitThermalBox = new System.Windows.Forms.TextBox();
            this.ThermalSourceBox = new System.Windows.Forms.TextBox();
            this.LoadTypesBox = new System.Windows.Forms.TextBox();
            this.Label6 = new System.Windows.Forms.Label();
            this.FEPathBox = new System.Windows.Forms.TextBox();
            this.GroupBox5 = new System.Windows.Forms.GroupBox();
            this.FormulaOption = new System.Windows.Forms.RadioButton();
            this.ValueOption = new System.Windows.Forms.RadioButton();
            this.Label10 = new System.Windows.Forms.Label();
            this.GroupBox4 = new System.Windows.Forms.GroupBox();
            this.Label5 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.btbBrowseLC = new System.Windows.Forms.PictureBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.BrowseBox = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.CellSelectBox = new System.Windows.Forms.PictureBox();
            this.ThermalOptionBox = new System.Windows.Forms.CheckBox();
            this.OpOptionCombineAvg = new System.Windows.Forms.RadioButton();
            this.OpOptionCombine = new System.Windows.Forms.RadioButton();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.SelTypeOptionNode = new System.Windows.Forms.RadioButton();
            this.SelTypeOption3D = new System.Windows.Forms.RadioButton();
            this.SelTypeOption2D = new System.Windows.Forms.RadioButton();
            this.SelTypeOption1D = new System.Windows.Forms.RadioButton();
            this.GroupBox3 = new System.Windows.Forms.GroupBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnMapFile = new System.Windows.Forms.PictureBox();
            this.grpElmType = new System.Windows.Forms.GroupBox();
            this.GroupBox5.SuspendLayout();
            this.GroupBox4.SuspendLayout();
            this.GroupBox2.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btbBrowseLC)).BeginInit();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.BrowseBox)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.CellSelectBox)).BeginInit();
            this.GroupBox1.SuspendLayout();
            this.GroupBox3.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnMapFile)).BeginInit();
            this.grpElmType.SuspendLayout();
            this.SuspendLayout();
            // 
            // DataSourceh5
            // 
            this.DataSourceh5.AutoSize = true;
            this.DataSourceh5.Enabled = false;
            this.DataSourceh5.Location = new System.Drawing.Point(285, 25);
            this.DataSourceh5.Margin = new System.Windows.Forms.Padding(2);
            this.DataSourceh5.Name = "DataSourceh5";
            this.DataSourceh5.Size = new System.Drawing.Size(53, 17);
            this.DataSourceh5.TabIndex = 3;
            this.DataSourceh5.TabStop = true;
            this.DataSourceh5.Text = "HDF5";
            this.DataSourceh5.UseVisualStyleBackColor = true;
            // 
            // TextBoxMapFile
            // 
            this.TextBoxMapFile.Enabled = false;
            this.TextBoxMapFile.Location = new System.Drawing.Point(6, 7);
            this.TextBoxMapFile.Margin = new System.Windows.Forms.Padding(2);
            this.TextBoxMapFile.Name = "TextBoxMapFile";
            this.TextBoxMapFile.Size = new System.Drawing.Size(306, 20);
            this.TextBoxMapFile.TabIndex = 2;
            this.ToolTip1.SetToolTip(this.TextBoxMapFile, "Mapping File is required if the CSV file contains the headings different from tha" +
        "t produced by Nastran H5 Reader");
            // 
            // DataSourceCSV
            // 
            this.DataSourceCSV.AutoSize = true;
            this.DataSourceCSV.Location = new System.Drawing.Point(150, 25);
            this.DataSourceCSV.Margin = new System.Windows.Forms.Padding(2);
            this.DataSourceCSV.Name = "DataSourceCSV";
            this.DataSourceCSV.Size = new System.Drawing.Size(69, 17);
            this.DataSourceCSV.TabIndex = 1;
            this.DataSourceCSV.TabStop = true;
            this.DataSourceCSV.Text = ".csv Files";
            this.DataSourceCSV.UseVisualStyleBackColor = true;
            // 
            // DataSourceRPT
            // 
            this.DataSourceRPT.AutoSize = true;
            this.DataSourceRPT.Location = new System.Drawing.Point(16, 25);
            this.DataSourceRPT.Margin = new System.Windows.Forms.Padding(2);
            this.DataSourceRPT.Name = "DataSourceRPT";
            this.DataSourceRPT.Size = new System.Drawing.Size(98, 17);
            this.DataSourceRPT.TabIndex = 0;
            this.DataSourceRPT.TabStop = true;
            this.DataSourceRPT.Text = "Patran .rpt Files";
            this.DataSourceRPT.UseVisualStyleBackColor = true;
            // 
            // Label7
            // 
            this.Label7.AutoSize = true;
            this.Label7.Location = new System.Drawing.Point(8, 62);
            this.Label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label7.Name = "Label7";
            this.Label7.Size = new System.Drawing.Size(67, 13);
            this.Label7.TabIndex = 0;
            this.Label7.Text = "Mapping File";
            // 
            // FileListBox
            // 
            this.FileListBox.FormattingEnabled = true;
            this.FileListBox.HorizontalScrollbar = true;
            this.FileListBox.Location = new System.Drawing.Point(3, 15);
            this.FileListBox.Margin = new System.Windows.Forms.Padding(2);
            this.FileListBox.Name = "FileListBox";
            this.FileListBox.Size = new System.Drawing.Size(179, 459);
            this.FileListBox.TabIndex = 0;
            this.ToolTip1.SetToolTip(this.FileListBox, "Double Click on the File Name to remove it");
            this.FileListBox.DoubleClick += new System.EventHandler(this.FileListBox_DoubleClick);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(421, 506);
            this.btnClose.Margin = new System.Windows.Forms.Padding(2);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(112, 24);
            this.btnClose.TabIndex = 17;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnCombLoads
            // 
            this.btnCombLoads.Location = new System.Drawing.Point(242, 506);
            this.btnCombLoads.Margin = new System.Windows.Forms.Padding(2);
            this.btnCombLoads.Name = "btnCombLoads";
            this.btnCombLoads.Size = new System.Drawing.Size(100, 24);
            this.btnCombLoads.TabIndex = 16;
            this.btnCombLoads.Text = "Combine Loads";
            this.btnCombLoads.UseVisualStyleBackColor = true;
            this.btnCombLoads.Click += new System.EventHandler(this.btnCombLoads_Click);
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(69, 506);
            this.btnReset.Margin = new System.Windows.Forms.Padding(2);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(94, 24);
            this.btnReset.TabIndex = 15;
            this.btnReset.Text = "Reset Form";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // ElemList
            // 
            this.ElemList.Location = new System.Drawing.Point(85, 58);
            this.ElemList.Margin = new System.Windows.Forms.Padding(2);
            this.ElemList.Multiline = true;
            this.ElemList.Name = "ElemList";
            this.ElemList.Size = new System.Drawing.Size(307, 48);
            this.ElemList.TabIndex = 1;
            this.ElemList.Text = "All";
            // 
            // LCFileBox
            // 
            this.LCFileBox.Location = new System.Drawing.Point(11, 2);
            this.LCFileBox.Margin = new System.Windows.Forms.Padding(2);
            this.LCFileBox.Name = "LCFileBox";
            this.LCFileBox.Size = new System.Drawing.Size(308, 20);
            this.LCFileBox.TabIndex = 1;
            // 
            // unitThermalBox
            // 
            this.unitThermalBox.Location = new System.Drawing.Point(3, 2);
            this.unitThermalBox.Margin = new System.Windows.Forms.Padding(2);
            this.unitThermalBox.Name = "unitThermalBox";
            this.unitThermalBox.Size = new System.Drawing.Size(275, 20);
            this.unitThermalBox.TabIndex = 1;
            // 
            // ThermalSourceBox
            // 
            this.ThermalSourceBox.Location = new System.Drawing.Point(298, 162);
            this.ThermalSourceBox.Margin = new System.Windows.Forms.Padding(2);
            this.ThermalSourceBox.Name = "ThermalSourceBox";
            this.ThermalSourceBox.Size = new System.Drawing.Size(94, 20);
            this.ThermalSourceBox.TabIndex = 1;
            // 
            // LoadTypesBox
            // 
            this.LoadTypesBox.Location = new System.Drawing.Point(85, 162);
            this.LoadTypesBox.Margin = new System.Windows.Forms.Padding(2);
            this.LoadTypesBox.Name = "LoadTypesBox";
            this.LoadTypesBox.Size = new System.Drawing.Size(94, 20);
            this.LoadTypesBox.TabIndex = 1;
            this.LoadTypesBox.Text = "1";
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(8, 209);
            this.Label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(99, 13);
            this.Label6.TabIndex = 0;
            this.Label6.Text = "Unit Thermal Loads";
            // 
            // FEPathBox
            // 
            this.FEPathBox.Location = new System.Drawing.Point(5, 3);
            this.FEPathBox.Margin = new System.Windows.Forms.Padding(2);
            this.FEPathBox.Name = "FEPathBox";
            this.FEPathBox.Size = new System.Drawing.Size(306, 20);
            this.FEPathBox.TabIndex = 1;
            // 
            // GroupBox5
            // 
            this.GroupBox5.Controls.Add(this.FormulaOption);
            this.GroupBox5.Controls.Add(this.ValueOption);
            this.GroupBox5.Location = new System.Drawing.Point(8, 442);
            this.GroupBox5.Margin = new System.Windows.Forms.Padding(2);
            this.GroupBox5.Name = "GroupBox5";
            this.GroupBox5.Padding = new System.Windows.Forms.Padding(2);
            this.GroupBox5.Size = new System.Drawing.Size(397, 50);
            this.GroupBox5.TabIndex = 18;
            this.GroupBox5.TabStop = false;
            this.GroupBox5.Text = "Output Type";
            // 
            // FormulaOption
            // 
            this.FormulaOption.AutoSize = true;
            this.FormulaOption.Enabled = false;
            this.FormulaOption.Location = new System.Drawing.Point(210, 22);
            this.FormulaOption.Margin = new System.Windows.Forms.Padding(2);
            this.FormulaOption.Name = "FormulaOption";
            this.FormulaOption.Size = new System.Drawing.Size(62, 17);
            this.FormulaOption.TabIndex = 1;
            this.FormulaOption.TabStop = true;
            this.FormulaOption.Text = "Formula";
            this.FormulaOption.UseVisualStyleBackColor = true;
            // 
            // ValueOption
            // 
            this.ValueOption.AutoSize = true;
            this.ValueOption.Location = new System.Drawing.Point(16, 22);
            this.ValueOption.Margin = new System.Windows.Forms.Padding(2);
            this.ValueOption.Name = "ValueOption";
            this.ValueOption.Size = new System.Drawing.Size(57, 17);
            this.ValueOption.TabIndex = 0;
            this.ValueOption.TabStop = true;
            this.ValueOption.Text = "Values";
            this.ValueOption.UseVisualStyleBackColor = true;
            // 
            // Label10
            // 
            this.Label10.AutoSize = true;
            this.Label10.Location = new System.Drawing.Point(121, 542);
            this.Label10.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label10.Name = "Label10";
            this.Label10.Size = new System.Drawing.Size(363, 13);
            this.Label10.TabIndex = 19;
            this.Label10.Text = "Copyright © 2020-2023 Raghavendra Prasad Laxman. All Rights Reserved.";
            // 
            // GroupBox4
            // 
            this.GroupBox4.Controls.Add(this.FileListBox);
            this.GroupBox4.Location = new System.Drawing.Point(412, 12);
            this.GroupBox4.Margin = new System.Windows.Forms.Padding(2);
            this.GroupBox4.Name = "GroupBox4";
            this.GroupBox4.Padding = new System.Windows.Forms.Padding(2);
            this.GroupBox4.Size = new System.Drawing.Size(186, 480);
            this.GroupBox4.TabIndex = 14;
            this.GroupBox4.TabStop = false;
            this.GroupBox4.Text = "File List";
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(215, 166);
            this.Label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(82, 13);
            this.Label5.TabIndex = 0;
            this.Label5.Text = "Thermal Source";
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(8, 129);
            this.Label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(39, 13);
            this.Label3.TabIndex = 0;
            this.Label3.Text = "LC File";
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(8, 166);
            this.Label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(73, 13);
            this.Label4.TabIndex = 0;
            this.Label4.Text = "Loads Source";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(8, 76);
            this.Label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(64, 13);
            this.Label2.TabIndex = 0;
            this.Label2.Text = "Element List";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(8, 27);
            this.Label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(54, 13);
            this.Label1.TabIndex = 0;
            this.Label1.Text = "FEM Path";
            // 
            // GroupBox2
            // 
            this.GroupBox2.Controls.Add(this.panel4);
            this.GroupBox2.Controls.Add(this.panel3);
            this.GroupBox2.Controls.Add(this.panel1);
            this.GroupBox2.Controls.Add(this.ElemList);
            this.GroupBox2.Controls.Add(this.ThermalSourceBox);
            this.GroupBox2.Controls.Add(this.LoadTypesBox);
            this.GroupBox2.Controls.Add(this.Label6);
            this.GroupBox2.Controls.Add(this.Label5);
            this.GroupBox2.Controls.Add(this.Label3);
            this.GroupBox2.Controls.Add(this.Label4);
            this.GroupBox2.Controls.Add(this.Label2);
            this.GroupBox2.Controls.Add(this.Label1);
            this.GroupBox2.Location = new System.Drawing.Point(8, 208);
            this.GroupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.GroupBox2.Size = new System.Drawing.Size(397, 232);
            this.GroupBox2.TabIndex = 12;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "Inputs";
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.btbBrowseLC);
            this.panel4.Controls.Add(this.LCFileBox);
            this.panel4.Location = new System.Drawing.Point(74, 126);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(321, 26);
            this.panel4.TabIndex = 4;
            // 
            // btbBrowseLC
            // 
            this.btbBrowseLC.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btbBrowseLC.Image = global::StressUtilities.Properties.Resources.OpenFile;
            this.btbBrowseLC.Location = new System.Drawing.Point(300, 3);
            this.btbBrowseLC.Margin = new System.Windows.Forms.Padding(2);
            this.btbBrowseLC.Name = "btbBrowseLC";
            this.btbBrowseLC.Size = new System.Drawing.Size(18, 16);
            this.btbBrowseLC.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.btbBrowseLC.TabIndex = 5;
            this.btbBrowseLC.TabStop = false;
            this.btbBrowseLC.Click += new System.EventHandler(this.btbBrowseLC_Click);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.BrowseBox);
            this.panel3.Controls.Add(this.FEPathBox);
            this.panel3.Location = new System.Drawing.Point(80, 22);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(315, 27);
            this.panel3.TabIndex = 3;
            // 
            // BrowseBox
            // 
            this.BrowseBox.Image = global::StressUtilities.Properties.Resources.FolderBottomPanel;
            this.BrowseBox.Location = new System.Drawing.Point(292, 5);
            this.BrowseBox.Name = "BrowseBox";
            this.BrowseBox.Size = new System.Drawing.Size(18, 16);
            this.BrowseBox.TabIndex = 2;
            this.BrowseBox.TabStop = false;
            this.BrowseBox.Click += new System.EventHandler(this.BrowseBox_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.CellSelectBox);
            this.panel1.Controls.Add(this.unitThermalBox);
            this.panel1.Location = new System.Drawing.Point(112, 203);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(281, 25);
            this.panel1.TabIndex = 0;
            // 
            // CellSelectBox
            // 
            this.CellSelectBox.Image = global::StressUtilities.Properties.Resources.SelectCell;
            this.CellSelectBox.InitialImage = global::StressUtilities.Properties.Resources.SelectCell;
            this.CellSelectBox.Location = new System.Drawing.Point(257, 4);
            this.CellSelectBox.Margin = new System.Windows.Forms.Padding(2);
            this.CellSelectBox.Name = "CellSelectBox";
            this.CellSelectBox.Size = new System.Drawing.Size(18, 17);
            this.CellSelectBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.CellSelectBox.TabIndex = 2;
            this.CellSelectBox.TabStop = false;
            this.CellSelectBox.Click += new System.EventHandler(this.CellSelectBox_Click);
            // 
            // ThermalOptionBox
            // 
            this.ThermalOptionBox.AutoSize = true;
            this.ThermalOptionBox.Location = new System.Drawing.Point(285, 24);
            this.ThermalOptionBox.Margin = new System.Windows.Forms.Padding(2);
            this.ThermalOptionBox.Name = "ThermalOptionBox";
            this.ThermalOptionBox.Size = new System.Drawing.Size(91, 17);
            this.ThermalOptionBox.TabIndex = 1;
            this.ThermalOptionBox.Text = "Thermal Load";
            this.ThermalOptionBox.UseVisualStyleBackColor = true;
            // 
            // OpOptionCombineAvg
            // 
            this.OpOptionCombineAvg.AutoSize = true;
            this.OpOptionCombineAvg.Location = new System.Drawing.Point(150, 24);
            this.OpOptionCombineAvg.Margin = new System.Windows.Forms.Padding(2);
            this.OpOptionCombineAvg.Name = "OpOptionCombineAvg";
            this.OpOptionCombineAvg.Size = new System.Drawing.Size(112, 17);
            this.OpOptionCombineAvg.TabIndex = 0;
            this.OpOptionCombineAvg.TabStop = true;
            this.OpOptionCombineAvg.Text = "Combine+Average";
            this.OpOptionCombineAvg.UseVisualStyleBackColor = true;
            // 
            // OpOptionCombine
            // 
            this.OpOptionCombine.AutoSize = true;
            this.OpOptionCombine.Checked = true;
            this.OpOptionCombine.Location = new System.Drawing.Point(16, 24);
            this.OpOptionCombine.Margin = new System.Windows.Forms.Padding(2);
            this.OpOptionCombine.Name = "OpOptionCombine";
            this.OpOptionCombine.Size = new System.Drawing.Size(66, 17);
            this.OpOptionCombine.TabIndex = 0;
            this.OpOptionCombine.TabStop = true;
            this.OpOptionCombine.Text = "Combine";
            this.OpOptionCombine.UseVisualStyleBackColor = true;
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.ThermalOptionBox);
            this.GroupBox1.Controls.Add(this.OpOptionCombineAvg);
            this.GroupBox1.Controls.Add(this.OpOptionCombine);
            this.GroupBox1.Location = new System.Drawing.Point(8, 60);
            this.GroupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.GroupBox1.Size = new System.Drawing.Size(397, 50);
            this.GroupBox1.TabIndex = 11;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "Operation";
            // 
            // SelTypeOptionNode
            // 
            this.SelTypeOptionNode.AutoSize = true;
            this.SelTypeOptionNode.Location = new System.Drawing.Point(313, 17);
            this.SelTypeOptionNode.Margin = new System.Windows.Forms.Padding(2);
            this.SelTypeOptionNode.Name = "SelTypeOptionNode";
            this.SelTypeOptionNode.Size = new System.Drawing.Size(51, 17);
            this.SelTypeOptionNode.TabIndex = 0;
            this.SelTypeOptionNode.Text = "Node";
            this.SelTypeOptionNode.UseVisualStyleBackColor = true;
            // 
            // SelTypeOption3D
            // 
            this.SelTypeOption3D.AutoSize = true;
            this.SelTypeOption3D.Location = new System.Drawing.Point(214, 17);
            this.SelTypeOption3D.Margin = new System.Windows.Forms.Padding(2);
            this.SelTypeOption3D.Name = "SelTypeOption3D";
            this.SelTypeOption3D.Size = new System.Drawing.Size(39, 17);
            this.SelTypeOption3D.TabIndex = 0;
            this.SelTypeOption3D.Text = "3D";
            this.SelTypeOption3D.UseVisualStyleBackColor = true;
            // 
            // SelTypeOption2D
            // 
            this.SelTypeOption2D.AutoSize = true;
            this.SelTypeOption2D.Location = new System.Drawing.Point(115, 17);
            this.SelTypeOption2D.Margin = new System.Windows.Forms.Padding(2);
            this.SelTypeOption2D.Name = "SelTypeOption2D";
            this.SelTypeOption2D.Size = new System.Drawing.Size(39, 17);
            this.SelTypeOption2D.TabIndex = 0;
            this.SelTypeOption2D.Text = "2D";
            this.SelTypeOption2D.UseVisualStyleBackColor = true;
            // 
            // SelTypeOption1D
            // 
            this.SelTypeOption1D.AutoSize = true;
            this.SelTypeOption1D.Location = new System.Drawing.Point(16, 17);
            this.SelTypeOption1D.Margin = new System.Windows.Forms.Padding(2);
            this.SelTypeOption1D.Name = "SelTypeOption1D";
            this.SelTypeOption1D.Size = new System.Drawing.Size(39, 17);
            this.SelTypeOption1D.TabIndex = 0;
            this.SelTypeOption1D.Text = "1D";
            this.SelTypeOption1D.UseVisualStyleBackColor = true;
            // 
            // GroupBox3
            // 
            this.GroupBox3.Controls.Add(this.panel2);
            this.GroupBox3.Controls.Add(this.DataSourceh5);
            this.GroupBox3.Controls.Add(this.DataSourceCSV);
            this.GroupBox3.Controls.Add(this.DataSourceRPT);
            this.GroupBox3.Controls.Add(this.Label7);
            this.GroupBox3.Location = new System.Drawing.Point(8, 115);
            this.GroupBox3.Margin = new System.Windows.Forms.Padding(2);
            this.GroupBox3.Name = "GroupBox3";
            this.GroupBox3.Padding = new System.Windows.Forms.Padding(2);
            this.GroupBox3.Size = new System.Drawing.Size(397, 89);
            this.GroupBox3.TabIndex = 13;
            this.GroupBox3.TabStop = false;
            this.GroupBox3.Text = "Data Source";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnMapFile);
            this.panel2.Controls.Add(this.TextBoxMapFile);
            this.panel2.Location = new System.Drawing.Point(79, 51);
            this.panel2.Margin = new System.Windows.Forms.Padding(2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(314, 34);
            this.panel2.TabIndex = 4;
            // 
            // btnMapFile
            // 
            this.btnMapFile.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnMapFile.Image = global::StressUtilities.Properties.Resources.OpenFile;
            this.btnMapFile.Location = new System.Drawing.Point(293, 9);
            this.btnMapFile.Margin = new System.Windows.Forms.Padding(2);
            this.btnMapFile.Name = "btnMapFile";
            this.btnMapFile.Size = new System.Drawing.Size(18, 16);
            this.btnMapFile.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.btnMapFile.TabIndex = 3;
            this.btnMapFile.TabStop = false;
            this.btnMapFile.Click += new System.EventHandler(this.btnMapFile_Click);
            // 
            // grpElmType
            // 
            this.grpElmType.Controls.Add(this.SelTypeOptionNode);
            this.grpElmType.Controls.Add(this.SelTypeOption3D);
            this.grpElmType.Controls.Add(this.SelTypeOption2D);
            this.grpElmType.Controls.Add(this.SelTypeOption1D);
            this.grpElmType.Location = new System.Drawing.Point(8, 12);
            this.grpElmType.Margin = new System.Windows.Forms.Padding(2);
            this.grpElmType.Name = "grpElmType";
            this.grpElmType.Padding = new System.Windows.Forms.Padding(2);
            this.grpElmType.Size = new System.Drawing.Size(397, 43);
            this.grpElmType.TabIndex = 10;
            this.grpElmType.TabStop = false;
            this.grpElmType.Text = "Entity Type";
            // 
            // CombinationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(606, 567);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnCombLoads);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.GroupBox5);
            this.Controls.Add(this.Label10);
            this.Controls.Add(this.GroupBox4);
            this.Controls.Add(this.GroupBox2);
            this.Controls.Add(this.GroupBox1);
            this.Controls.Add(this.GroupBox3);
            this.Controls.Add(this.grpElmType);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "CombinationForm";
            this.Text = "Load Combination";
            this.GroupBox5.ResumeLayout(false);
            this.GroupBox5.PerformLayout();
            this.GroupBox4.ResumeLayout(false);
            this.GroupBox2.ResumeLayout(false);
            this.GroupBox2.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btbBrowseLC)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.BrowseBox)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.CellSelectBox)).EndInit();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.GroupBox3.ResumeLayout(false);
            this.GroupBox3.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnMapFile)).EndInit();
            this.grpElmType.ResumeLayout(false);
            this.grpElmType.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.RadioButton DataSourceh5;
        internal System.Windows.Forms.TextBox TextBoxMapFile;
        internal System.Windows.Forms.ToolTip ToolTip1;
        internal System.Windows.Forms.RadioButton DataSourceCSV;
        internal System.Windows.Forms.RadioButton DataSourceRPT;
        internal System.Windows.Forms.Label Label7;
        internal System.Windows.Forms.ListBox FileListBox;
        internal System.Windows.Forms.Button btnClose;
        internal System.Windows.Forms.Button btnCombLoads;
        internal System.Windows.Forms.Button btnReset;
        internal System.Windows.Forms.TextBox ElemList;
        internal System.Windows.Forms.TextBox LCFileBox;
        internal System.Windows.Forms.TextBox unitThermalBox;
        internal System.Windows.Forms.TextBox ThermalSourceBox;
        internal System.Windows.Forms.TextBox LoadTypesBox;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.TextBox FEPathBox;
        internal System.Windows.Forms.GroupBox GroupBox5;
        internal System.Windows.Forms.RadioButton FormulaOption;
        internal System.Windows.Forms.RadioButton ValueOption;
        internal System.Windows.Forms.Label Label10;
        internal System.Windows.Forms.GroupBox GroupBox4;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.GroupBox GroupBox2;
        internal System.Windows.Forms.CheckBox ThermalOptionBox;
        internal System.Windows.Forms.RadioButton OpOptionCombineAvg;
        internal System.Windows.Forms.RadioButton OpOptionCombine;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.RadioButton SelTypeOptionNode;
        internal System.Windows.Forms.RadioButton SelTypeOption3D;
        internal System.Windows.Forms.RadioButton SelTypeOption2D;
        internal System.Windows.Forms.RadioButton SelTypeOption1D;
        internal System.Windows.Forms.GroupBox GroupBox3;
        internal System.Windows.Forms.GroupBox grpElmType;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox CellSelectBox;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.PictureBox btnMapFile;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.PictureBox BrowseBox;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.PictureBox btbBrowseLC;
    }
}