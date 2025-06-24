namespace Nastranh5
{
    partial class Nash5
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Nash5));
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.h5Menu = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.resultsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.singleFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.FileListBox = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnLCFile = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SubCaseBox = new System.Windows.Forms.TextBox();
            this.textBoxEntity = new System.Windows.Forms.TextBox();
            this.textBoxLC = new System.Windows.Forms.TextBox();
            this.textBoxDataSet = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.SelectCellBox = new System.Windows.Forms.PictureBox();
            this.StartCell = new System.Windows.Forms.TextBox();
            this.PrincRequestBox = new System.Windows.Forms.CheckBox();
            this.vMRequestBox = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cBLocation = new System.Windows.Forms.ComboBox();
            this.cBOptions = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.h5Toolstrip = new System.Windows.Forms.ToolStrip();
            this.btnOpenDB = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnDBExtract = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.BtnNodeExpand = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.BtnAllNodeExpand = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.btnCloseDB = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.btnCloseAllDB = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
            this.tpltToolStrip = new System.Windows.Forms.ToolStrip();
            this.btnLCTemplate = new System.Windows.Forms.ToolStripButton();
            this.h5Menu.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SelectCellBox)).BeginInit();
            this.flowLayoutPanel1.SuspendLayout();
            this.h5Toolstrip.SuspendLayout();
            this.tpltToolStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            resources.ApplyResources(this.treeView1, "treeView1");
            this.treeView1.Name = "treeView1";
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // h5Menu
            // 
            this.h5Menu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.h5Menu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.resultsToolStripMenuItem});
            resources.ApplyResources(this.h5Menu, "h5Menu");
            this.h5Menu.Name = "h5Menu";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.closeToolStripMenuItem,
            this.closeAllToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            resources.ApplyResources(this.fileToolStripMenuItem, "fileToolStripMenuItem");
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            resources.ApplyResources(this.openToolStripMenuItem, "openToolStripMenuItem");
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            resources.ApplyResources(this.closeToolStripMenuItem, "closeToolStripMenuItem");
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
            // 
            // closeAllToolStripMenuItem
            // 
            this.closeAllToolStripMenuItem.Name = "closeAllToolStripMenuItem";
            resources.ApplyResources(this.closeAllToolStripMenuItem, "closeAllToolStripMenuItem");
            this.closeAllToolStripMenuItem.Click += new System.EventHandler(this.closeAllToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            resources.ApplyResources(this.exitToolStripMenuItem, "exitToolStripMenuItem");
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // resultsToolStripMenuItem
            // 
            this.resultsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.singleFileToolStripMenuItem});
            this.resultsToolStripMenuItem.Name = "resultsToolStripMenuItem";
            resources.ApplyResources(this.resultsToolStripMenuItem, "resultsToolStripMenuItem");
            // 
            // singleFileToolStripMenuItem
            // 
            this.singleFileToolStripMenuItem.Name = "singleFileToolStripMenuItem";
            resources.ApplyResources(this.singleFileToolStripMenuItem, "singleFileToolStripMenuItem");
            this.singleFileToolStripMenuItem.Click += new System.EventHandler(this.singleFileToolStripMenuItem_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            resources.ApplyResources(this.statusStrip1, "statusStrip1");
            this.statusStrip1.Name = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            resources.ApplyResources(this.toolStripStatusLabel1, "toolStripStatusLabel1");
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            resources.ApplyResources(this.toolStripProgressBar1, "toolStripProgressBar1");
            // 
            // FileListBox
            // 
            this.FileListBox.FormattingEnabled = true;
            resources.ApplyResources(this.FileListBox, "FileListBox");
            this.FileListBox.Name = "FileListBox";
            this.FileListBox.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.FileListBox_MouseDoubleClick);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.treeView1);
            this.groupBox1.Controls.Add(this.splitter1);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnLCFile);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.SubCaseBox);
            this.groupBox3.Controls.Add(this.textBoxEntity);
            this.groupBox3.Controls.Add(this.textBoxLC);
            this.groupBox3.Controls.Add(this.textBoxDataSet);
            this.groupBox3.Controls.Add(this.FileListBox);
            resources.ApplyResources(this.groupBox3, "groupBox3");
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.TabStop = false;
            // 
            // btnLCFile
            // 
            resources.ApplyResources(this.btnLCFile, "btnLCFile");
            this.btnLCFile.Name = "btnLCFile";
            this.btnLCFile.UseVisualStyleBackColor = true;
            this.btnLCFile.Click += new System.EventHandler(this.btnLCFile_Click);
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // SubCaseBox
            // 
            resources.ApplyResources(this.SubCaseBox, "SubCaseBox");
            this.SubCaseBox.Name = "SubCaseBox";
            // 
            // textBoxEntity
            // 
            resources.ApplyResources(this.textBoxEntity, "textBoxEntity");
            this.textBoxEntity.Name = "textBoxEntity";
            // 
            // textBoxLC
            // 
            resources.ApplyResources(this.textBoxLC, "textBoxLC");
            this.textBoxLC.Name = "textBoxLC";
            // 
            // textBoxDataSet
            // 
            resources.ApplyResources(this.textBoxDataSet, "textBoxDataSet");
            this.textBoxDataSet.Name = "textBoxDataSet";
            this.textBoxDataSet.ReadOnly = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.panel1);
            this.groupBox2.Controls.Add(this.PrincRequestBox);
            this.groupBox2.Controls.Add(this.vMRequestBox);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.cBLocation);
            this.groupBox2.Controls.Add(this.cBOptions);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label5);
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.SelectCellBox);
            this.panel1.Controls.Add(this.StartCell);
            resources.ApplyResources(this.panel1, "panel1");
            this.panel1.Name = "panel1";
            // 
            // SelectCellBox
            // 
            this.SelectCellBox.ErrorImage = global::StressUtilities.Properties.Resources.SelectCell;
            this.SelectCellBox.Image = global::StressUtilities.Properties.Resources.SelectCell;
            this.SelectCellBox.InitialImage = global::StressUtilities.Properties.Resources.SelectCell;
            resources.ApplyResources(this.SelectCellBox, "SelectCellBox");
            this.SelectCellBox.Name = "SelectCellBox";
            this.SelectCellBox.TabStop = false;
            this.SelectCellBox.Click += new System.EventHandler(this.SelectCellBox_Click);
            // 
            // StartCell
            // 
            resources.ApplyResources(this.StartCell, "StartCell");
            this.StartCell.Name = "StartCell";
            // 
            // PrincRequestBox
            // 
            resources.ApplyResources(this.PrincRequestBox, "PrincRequestBox");
            this.PrincRequestBox.Name = "PrincRequestBox";
            this.PrincRequestBox.UseVisualStyleBackColor = true;
            // 
            // vMRequestBox
            // 
            resources.ApplyResources(this.vMRequestBox, "vMRequestBox");
            this.vMRequestBox.Name = "vMRequestBox";
            this.vMRequestBox.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            resources.ApplyResources(this.label7, "label7");
            this.label7.Name = "label7";
            // 
            // cBLocation
            // 
            this.cBLocation.FormattingEnabled = true;
            resources.ApplyResources(this.cBLocation, "cBLocation");
            this.cBLocation.Name = "cBLocation";
            // 
            // cBOptions
            // 
            resources.ApplyResources(this.cBOptions, "cBOptions");
            this.cBOptions.FormattingEnabled = true;
            this.cBOptions.Name = "cBOptions";
            // 
            // label6
            // 
            resources.ApplyResources(this.label6, "label6");
            this.label6.Name = "label6";
            // 
            // label8
            // 
            resources.ApplyResources(this.label8, "label8");
            this.label8.Name = "label8";
            // 
            // label5
            // 
            resources.ApplyResources(this.label5, "label5");
            this.label5.Name = "label5";
            // 
            // splitter1
            // 
            resources.ApplyResources(this.splitter1, "splitter1");
            this.splitter1.Name = "splitter1";
            this.splitter1.TabStop = false;
            // 
            // flowLayoutPanel1
            // 
            resources.ApplyResources(this.flowLayoutPanel1, "flowLayoutPanel1");
            this.flowLayoutPanel1.Controls.Add(this.h5Toolstrip);
            this.flowLayoutPanel1.Controls.Add(this.tpltToolStrip);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            // 
            // h5Toolstrip
            // 
            this.h5Toolstrip.AllowItemReorder = true;
            resources.ApplyResources(this.h5Toolstrip, "h5Toolstrip");
            this.h5Toolstrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.h5Toolstrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnOpenDB,
            this.toolStripSeparator1,
            this.btnDBExtract,
            this.toolStripSeparator2,
            this.BtnNodeExpand,
            this.toolStripSeparator3,
            this.BtnAllNodeExpand,
            this.toolStripSeparator4,
            this.btnCloseDB,
            this.toolStripSeparator5,
            this.btnCloseAllDB,
            this.toolStripSeparator6,
            this.btnExit,
            this.toolStripSeparator7});
            this.h5Toolstrip.Name = "h5Toolstrip";
            this.h5Toolstrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.h5Toolstrip.Stretch = true;
            this.h5Toolstrip.TabStop = true;
            // 
            // btnOpenDB
            // 
            this.btnOpenDB.Image = global::StressUtilities.Properties.Resources.AddDatabase;
            resources.ApplyResources(this.btnOpenDB, "btnOpenDB");
            this.btnOpenDB.MergeIndex = 1;
            this.btnOpenDB.Name = "btnOpenDB";
            this.btnOpenDB.Click += new System.EventHandler(this.btnOpenDB_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            resources.ApplyResources(this.toolStripSeparator1, "toolStripSeparator1");
            // 
            // btnDBExtract
            // 
            this.btnDBExtract.CheckOnClick = true;
            this.btnDBExtract.Image = global::StressUtilities.Properties.Resources.Execute;
            resources.ApplyResources(this.btnDBExtract, "btnDBExtract");
            this.btnDBExtract.MergeIndex = 2;
            this.btnDBExtract.Name = "btnDBExtract";
            this.btnDBExtract.Click += new System.EventHandler(this.btnDBExtract_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            resources.ApplyResources(this.toolStripSeparator2, "toolStripSeparator2");
            // 
            // BtnNodeExpand
            // 
            this.BtnNodeExpand.Image = global::StressUtilities.Properties.Resources.Expand;
            resources.ApplyResources(this.BtnNodeExpand, "BtnNodeExpand");
            this.BtnNodeExpand.MergeIndex = 3;
            this.BtnNodeExpand.Name = "BtnNodeExpand";
            this.BtnNodeExpand.Click += new System.EventHandler(this.BtnNodeExpand_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            resources.ApplyResources(this.toolStripSeparator3, "toolStripSeparator3");
            // 
            // BtnAllNodeExpand
            // 
            this.BtnAllNodeExpand.Image = global::StressUtilities.Properties.Resources.ExpandAll;
            this.BtnAllNodeExpand.MergeIndex = 3;
            this.BtnAllNodeExpand.Name = "BtnAllNodeExpand";
            resources.ApplyResources(this.BtnAllNodeExpand, "BtnAllNodeExpand");
            this.BtnAllNodeExpand.CheckedChanged += new System.EventHandler(this.BtnAllNodeExpand_CheckedChanged);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            resources.ApplyResources(this.toolStripSeparator4, "toolStripSeparator4");
            // 
            // btnCloseDB
            // 
            this.btnCloseDB.Image = global::StressUtilities.Properties.Resources.CloseDocument;
            resources.ApplyResources(this.btnCloseDB, "btnCloseDB");
            this.btnCloseDB.MergeIndex = 4;
            this.btnCloseDB.Name = "btnCloseDB";
            this.btnCloseDB.Click += new System.EventHandler(this.btnCloseDB_Click);
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            resources.ApplyResources(this.toolStripSeparator5, "toolStripSeparator5");
            // 
            // btnCloseAllDB
            // 
            this.btnCloseAllDB.Image = global::StressUtilities.Properties.Resources.CloseDocumentGroup;
            resources.ApplyResources(this.btnCloseAllDB, "btnCloseAllDB");
            this.btnCloseAllDB.MergeIndex = 5;
            this.btnCloseAllDB.Name = "btnCloseAllDB";
            this.btnCloseAllDB.Click += new System.EventHandler(this.btnCloseAllDB_Click);
            // 
            // toolStripSeparator6
            // 
            this.toolStripSeparator6.Name = "toolStripSeparator6";
            resources.ApplyResources(this.toolStripSeparator6, "toolStripSeparator6");
            // 
            // btnExit
            // 
            this.btnExit.Image = global::StressUtilities.Properties.Resources.Exit;
            resources.ApplyResources(this.btnExit, "btnExit");
            this.btnExit.MergeIndex = 6;
            this.btnExit.Name = "btnExit";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // toolStripSeparator7
            // 
            this.toolStripSeparator7.Name = "toolStripSeparator7";
            resources.ApplyResources(this.toolStripSeparator7, "toolStripSeparator7");
            // 
            // tpltToolStrip
            // 
            this.tpltToolStrip.AllowItemReorder = true;
            resources.ApplyResources(this.tpltToolStrip, "tpltToolStrip");
            this.flowLayoutPanel1.SetFlowBreak(this.tpltToolStrip, true);
            this.tpltToolStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.tpltToolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnLCTemplate});
            this.tpltToolStrip.Name = "tpltToolStrip";
            this.tpltToolStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.tpltToolStrip.Stretch = true;
            this.tpltToolStrip.TabStop = true;
            // 
            // btnLCTemplate
            // 
            resources.ApplyResources(this.btnLCTemplate, "btnLCTemplate");
            this.btnLCTemplate.Name = "btnLCTemplate";
            this.btnLCTemplate.Click += new System.EventHandler(this.btnLCTemplate_Click);
            // 
            // Nash5
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.h5Menu);
            this.Controls.Add(this.groupBox1);
            this.MainMenuStrip = this.h5Menu;
            this.Name = "Nash5";
            this.h5Menu.ResumeLayout(false);
            this.h5Menu.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SelectCellBox)).EndInit();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.h5Toolstrip.ResumeLayout(false);
            this.h5Toolstrip.PerformLayout();
            this.tpltToolStrip.ResumeLayout(false);
            this.tpltToolStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.MenuStrip h5Menu;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        //private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem closeAllToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.ToolStripMenuItem resultsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem singleFileToolStripMenuItem;
        private System.Windows.Forms.ListBox FileListBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBoxDataSet;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxEntity;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxLC;
        private System.Windows.Forms.ComboBox cBLocation;
        private System.Windows.Forms.ComboBox cBOptions;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.CheckBox PrincRequestBox;
        private System.Windows.Forms.CheckBox vMRequestBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.ToolStrip h5Toolstrip;
        private System.Windows.Forms.ToolStripButton btnOpenDB;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnDBExtract;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripButton btnCloseDB;
        private System.Windows.Forms.ToolStripButton btnCloseAllDB;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton btnExit;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStrip tpltToolStrip;
        private System.Windows.Forms.ToolStripButton btnLCTemplate;
        private System.Windows.Forms.Button btnLCFile;
        private System.Windows.Forms.ToolStripButton BtnAllNodeExpand;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
        private System.Windows.Forms.ToolStripButton BtnNodeExpand;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox SubCaseBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox SelectCellBox;
        private System.Windows.Forms.TextBox StartCell;
    }
}

