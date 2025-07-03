namespace StressUtilities.Forms
{
    /*
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/
*/

    partial class ReportContents
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportContents));
            this.CaptionText = new System.Windows.Forms.TextBox();
            this.Label12 = new System.Windows.Forms.Label();
            this.optTbl = new System.Windows.Forms.RadioButton();
            this.optCalcTbl = new System.Windows.Forms.RadioButton();
            this.optCalc = new System.Windows.Forms.RadioButton();
            this.DeleteBox = new System.Windows.Forms.Button();
            this.cmbTableList = new System.Windows.Forms.ComboBox();
            this.SummaryParam = new System.Windows.Forms.TextBox();
            this.btnPictures = new System.Windows.Forms.Button();
            this.PictureID = new System.Windows.Forms.TextBox();
            this.Label1 = new System.Windows.Forms.Label();
            this.CriticalItem = new System.Windows.Forms.TextBox();
            this.Parameters = new System.Windows.Forms.TextBox();
            this.CaptionBox = new System.Windows.Forms.TextBox();
            this.Label13 = new System.Windows.Forms.Label();
            this.SectionBox = new System.Windows.Forms.TextBox();
            this.OptionCharts = new System.Windows.Forms.RadioButton();
            this.Label8 = new System.Windows.Forms.Label();
            this.Label7 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.Label5 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.OptionPictures = new System.Windows.Forms.RadioButton();
            this.BtnAddtoList = new System.Windows.Forms.Button();
            this.FlowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.Label9 = new System.Windows.Forms.Label();
            this.ListBoxGeometry = new System.Windows.Forms.ListBox();
            this.ToolTipTbl = new System.Windows.Forms.ToolTip(this.components);
            this.Label10 = new System.Windows.Forms.Label();
            this.BtnClose = new System.Windows.Forms.Button();
            this.AddSelectedBox = new System.Windows.Forms.Button();
            this.Label3 = new System.Windows.Forms.Label();
            this.BtnPicturesList = new System.Windows.Forms.Button();
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.GroupBox2.SuspendLayout();
            this.GroupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // CaptionText
            // 
            this.CaptionText.Location = new System.Drawing.Point(96, 51);
            this.CaptionText.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.CaptionText.Name = "CaptionText";
            this.CaptionText.Size = new System.Drawing.Size(176, 20);
            this.CaptionText.TabIndex = 7;
            // 
            // Label12
            // 
            this.Label12.AutoSize = true;
            this.Label12.Location = new System.Drawing.Point(278, 66);
            this.Label12.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label12.Name = "Label12";
            this.Label12.Size = new System.Drawing.Size(135, 13);
            this.Label12.TabIndex = 14;
            this.Label12.Text = "(Critical Item for the Report)";
            // 
            // optTbl
            // 
            this.optTbl.AutoSize = true;
            this.optTbl.Location = new System.Drawing.Point(378, 205);
            this.optTbl.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.optTbl.Name = "optTbl";
            this.optTbl.Size = new System.Drawing.Size(52, 17);
            this.optTbl.TabIndex = 11;
            this.optTbl.TabStop = true;
            this.optTbl.Text = "Table";
            this.optTbl.UseVisualStyleBackColor = true;
            // 
            // optCalcTbl
            // 
            this.optCalcTbl.AutoSize = true;
            this.optCalcTbl.Location = new System.Drawing.Point(235, 205);
            this.optCalcTbl.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.optCalcTbl.Name = "optCalcTbl";
            this.optCalcTbl.Size = new System.Drawing.Size(110, 17);
            this.optCalcTbl.TabIndex = 10;
            this.optCalcTbl.TabStop = true;
            this.optCalcTbl.Text = "Calculation+Table";
            this.optCalcTbl.UseVisualStyleBackColor = true;
            // 
            // optCalc
            // 
            this.optCalc.AutoSize = true;
            this.optCalc.Location = new System.Drawing.Point(102, 205);
            this.optCalc.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.optCalc.Name = "optCalc";
            this.optCalc.Size = new System.Drawing.Size(120, 17);
            this.optCalc.TabIndex = 9;
            this.optCalc.TabStop = true;
            this.optCalc.Text = "Calculation (Default)";
            this.optCalc.UseVisualStyleBackColor = true;
            // 
            // DeleteBox
            // 
            this.DeleteBox.Location = new System.Drawing.Point(278, 24);
            this.DeleteBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DeleteBox.Name = "DeleteBox";
            this.DeleteBox.Size = new System.Drawing.Size(111, 20);
            this.DeleteBox.TabIndex = 3;
            this.DeleteBox.Text = "Delete Table Names";
            this.DeleteBox.UseVisualStyleBackColor = true;
            this.DeleteBox.Click += new System.EventHandler(this.DeleteBox_Click);
            // 
            // cmbTableList
            // 
            this.cmbTableList.FormattingEnabled = true;
            this.cmbTableList.Location = new System.Drawing.Point(95, 24);
            this.cmbTableList.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cmbTableList.Name = "cmbTableList";
            this.cmbTableList.Size = new System.Drawing.Size(176, 21);
            this.cmbTableList.TabIndex = 2;
            this.cmbTableList.SelectedIndexChanged += new System.EventHandler(this.cmbTableList_SelectedIndexChanged);
            // 
            // SummaryParam
            // 
            this.SummaryParam.Location = new System.Drawing.Point(95, 133);
            this.SummaryParam.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SummaryParam.Name = "SummaryParam";
            this.SummaryParam.Size = new System.Drawing.Size(451, 20);
            this.SummaryParam.TabIndex = 6;
            // 
            // btnPictures
            // 
            this.btnPictures.Location = new System.Drawing.Point(458, 17);
            this.btnPictures.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnPictures.Name = "btnPictures";
            this.btnPictures.Size = new System.Drawing.Size(88, 21);
            this.btnPictures.TabIndex = 14;
            this.btnPictures.Text = "Load Images";
            this.btnPictures.UseVisualStyleBackColor = true;
            this.btnPictures.Click += new System.EventHandler(this.btnPictures_Click);
            // 
            // PictureID
            // 
            this.PictureID.Location = new System.Drawing.Point(406, 51);
            this.PictureID.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.PictureID.Name = "PictureID";
            this.PictureID.Size = new System.Drawing.Size(140, 20);
            this.PictureID.TabIndex = 8;
            this.PictureID.Text = "Set";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(9, 55);
            this.Label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(75, 13);
            this.Label1.TabIndex = 1;
            this.Label1.Text = "Image Caption";
            // 
            // CriticalItem
            // 
            this.CriticalItem.Location = new System.Drawing.Point(95, 62);
            this.CriticalItem.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.CriticalItem.Name = "CriticalItem";
            this.CriticalItem.Size = new System.Drawing.Size(176, 20);
            this.CriticalItem.TabIndex = 4;
            // 
            // Parameters
            // 
            this.Parameters.Location = new System.Drawing.Point(95, 98);
            this.Parameters.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Parameters.Name = "Parameters";
            this.Parameters.Size = new System.Drawing.Size(451, 20);
            this.Parameters.TabIndex = 5;
            // 
            // CaptionBox
            // 
            this.CaptionBox.Location = new System.Drawing.Point(95, 169);
            this.CaptionBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.CaptionBox.Name = "CaptionBox";
            this.CaptionBox.Size = new System.Drawing.Size(176, 20);
            this.CaptionBox.TabIndex = 7;
            // 
            // Label13
            // 
            this.Label13.AutoSize = true;
            this.Label13.Location = new System.Drawing.Point(312, 55);
            this.Label13.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label13.Name = "Label13";
            this.Label13.Size = new System.Drawing.Size(82, 13);
            this.Label13.TabIndex = 0;
            this.Label13.Text = "Image Group ID";
            // 
            // SectionBox
            // 
            this.SectionBox.Location = new System.Drawing.Point(406, 169);
            this.SectionBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SectionBox.Name = "SectionBox";
            this.SectionBox.Size = new System.Drawing.Size(140, 20);
            this.SectionBox.TabIndex = 8;
            // 
            // OptionCharts
            // 
            this.OptionCharts.AutoSize = true;
            this.OptionCharts.Location = new System.Drawing.Point(293, 19);
            this.OptionCharts.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.OptionCharts.Name = "OptionCharts";
            this.OptionCharts.Size = new System.Drawing.Size(84, 17);
            this.OptionCharts.TabIndex = 13;
            this.OptionCharts.TabStop = true;
            this.OptionCharts.Text = "Excel Charts";
            this.OptionCharts.UseVisualStyleBackColor = true;
            // 
            // Label8
            // 
            this.Label8.AutoSize = true;
            this.Label8.Location = new System.Drawing.Point(8, 135);
            this.Label8.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label8.Name = "Label8";
            this.Label8.Size = new System.Drawing.Size(77, 13);
            this.Label8.TabIndex = 6;
            this.Label8.Text = "Table Columns";
            // 
            // Label7
            // 
            this.Label7.AutoSize = true;
            this.Label7.Location = new System.Drawing.Point(8, 206);
            this.Label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label7.Name = "Label7";
            this.Label7.Size = new System.Drawing.Size(47, 13);
            this.Label7.TabIndex = 5;
            this.Label7.Text = "Options*";
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(9, 99);
            this.Label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(64, 13);
            this.Label6.TabIndex = 4;
            this.Label6.Text = "Parameters*";
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(8, 63);
            this.Label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(65, 13);
            this.Label5.TabIndex = 3;
            this.Label5.Text = "Critical Item*";
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(8, 28);
            this.Label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(52, 13);
            this.Label4.TabIndex = 2;
            this.Label4.Text = "Table ID*";
            // 
            // OptionPictures
            // 
            this.OptionPictures.AutoSize = true;
            this.OptionPictures.Location = new System.Drawing.Point(159, 19);
            this.OptionPictures.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.OptionPictures.Name = "OptionPictures";
            this.OptionPictures.Size = new System.Drawing.Size(63, 17);
            this.OptionPictures.TabIndex = 12;
            this.OptionPictures.TabStop = true;
            this.OptionPictures.Text = "Pictures";
            this.OptionPictures.UseVisualStyleBackColor = true;
            // 
            // BtnAddtoList
            // 
            this.BtnAddtoList.Location = new System.Drawing.Point(572, 79);
            this.BtnAddtoList.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnAddtoList.Name = "BtnAddtoList";
            this.BtnAddtoList.Size = new System.Drawing.Size(140, 21);
            this.BtnAddtoList.TabIndex = 25;
            this.BtnAddtoList.Text = "Add &All Tables";
            this.BtnAddtoList.UseVisualStyleBackColor = true;
            this.BtnAddtoList.Click += new System.EventHandler(this.BtnAddtoList_Click);
            // 
            // FlowLayoutPanel1
            // 
            this.FlowLayoutPanel1.Location = new System.Drawing.Point(562, 15);
            this.FlowLayoutPanel1.Name = "FlowLayoutPanel1";
            this.FlowLayoutPanel1.Size = new System.Drawing.Size(160, 473);
            this.FlowLayoutPanel1.TabIndex = 23;
            // 
            // Label9
            // 
            this.Label9.AutoSize = true;
            this.Label9.Location = new System.Drawing.Point(311, 173);
            this.Label9.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label9.Name = "Label9";
            this.Label9.Size = new System.Drawing.Size(95, 13);
            this.Label9.TabIndex = 0;
            this.Label9.Text = "Applicable Section";
            // 
            // ListBoxGeometry
            // 
            this.ListBoxGeometry.FormattingEnabled = true;
            this.ListBoxGeometry.HorizontalScrollbar = true;
            this.ListBoxGeometry.Location = new System.Drawing.Point(18, 330);
            this.ListBoxGeometry.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.ListBoxGeometry.Name = "ListBoxGeometry";
            this.ListBoxGeometry.Size = new System.Drawing.Size(536, 186);
            this.ListBoxGeometry.TabIndex = 24;
            this.ToolTipTbl.SetToolTip(this.ListBoxGeometry, "Double Click to Remove");
            this.ListBoxGeometry.DoubleClick += new System.EventHandler(this.ListBoxGeometry_DoubleClick);
            // 
            // Label10
            // 
            this.Label10.AutoSize = true;
            this.Label10.Location = new System.Drawing.Point(128, 531);
            this.Label10.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label10.Name = "Label10";
            this.Label10.Size = new System.Drawing.Size(284, 13);
            this.Label10.TabIndex = 21;
            this.Label10.Text = "Copyright © 2020-2028 Stress Utilities. All Rights Reserved";
            // 
            // BtnClose
            // 
            this.BtnClose.Location = new System.Drawing.Point(572, 434);
            this.BtnClose.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(140, 25);
            this.BtnClose.TabIndex = 28;
            this.BtnClose.Text = "&Close";
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // AddSelectedBox
            // 
            this.AddSelectedBox.Location = new System.Drawing.Point(572, 167);
            this.AddSelectedBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.AddSelectedBox.Name = "AddSelectedBox";
            this.AddSelectedBox.Size = new System.Drawing.Size(140, 20);
            this.AddSelectedBox.TabIndex = 26;
            this.AddSelectedBox.Text = "Add &Selected Table";
            this.AddSelectedBox.UseVisualStyleBackColor = true;
            this.AddSelectedBox.Click += new System.EventHandler(this.AddSelectedBox_Click);
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(8, 171);
            this.Label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(73, 13);
            this.Label3.TabIndex = 1;
            this.Label3.Text = "Table Caption";
            // 
            // BtnPicturesList
            // 
            this.BtnPicturesList.Location = new System.Drawing.Point(572, 254);
            this.BtnPicturesList.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnPicturesList.Name = "BtnPicturesList";
            this.BtnPicturesList.Size = new System.Drawing.Size(140, 21);
            this.BtnPicturesList.TabIndex = 27;
            this.BtnPicturesList.Text = "Add &Images";
            this.BtnPicturesList.UseVisualStyleBackColor = true;
            this.BtnPicturesList.Click += new System.EventHandler(this.BtnPicturesList_Click);
            // 
            // GroupBox2
            // 
            this.GroupBox2.Controls.Add(this.Label12);
            this.GroupBox2.Controls.Add(this.optTbl);
            this.GroupBox2.Controls.Add(this.optCalcTbl);
            this.GroupBox2.Controls.Add(this.optCalc);
            this.GroupBox2.Controls.Add(this.DeleteBox);
            this.GroupBox2.Controls.Add(this.cmbTableList);
            this.GroupBox2.Controls.Add(this.SummaryParam);
            this.GroupBox2.Controls.Add(this.CriticalItem);
            this.GroupBox2.Controls.Add(this.Parameters);
            this.GroupBox2.Controls.Add(this.CaptionBox);
            this.GroupBox2.Controls.Add(this.SectionBox);
            this.GroupBox2.Controls.Add(this.Label8);
            this.GroupBox2.Controls.Add(this.Label7);
            this.GroupBox2.Controls.Add(this.Label6);
            this.GroupBox2.Controls.Add(this.Label5);
            this.GroupBox2.Controls.Add(this.Label4);
            this.GroupBox2.Controls.Add(this.Label9);
            this.GroupBox2.Controls.Add(this.Label3);
            this.GroupBox2.Location = new System.Drawing.Point(8, 9);
            this.GroupBox2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.GroupBox2.Size = new System.Drawing.Size(550, 244);
            this.GroupBox2.TabIndex = 20;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "Inputs";
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.OptionCharts);
            this.GroupBox1.Controls.Add(this.OptionPictures);
            this.GroupBox1.Controls.Add(this.btnPictures);
            this.GroupBox1.Controls.Add(this.PictureID);
            this.GroupBox1.Controls.Add(this.Label1);
            this.GroupBox1.Controls.Add(this.Label13);
            this.GroupBox1.Controls.Add(this.CaptionText);
            this.GroupBox1.Location = new System.Drawing.Point(8, 254);
            this.GroupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.GroupBox1.Size = new System.Drawing.Size(550, 268);
            this.GroupBox1.TabIndex = 22;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "Images";
            // 
            // ReportContents
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(729, 554);
            this.Controls.Add(this.BtnAddtoList);
            this.Controls.Add(this.ListBoxGeometry);
            this.Controls.Add(this.Label10);
            this.Controls.Add(this.BtnClose);
            this.Controls.Add(this.AddSelectedBox);
            this.Controls.Add(this.BtnPicturesList);
            this.Controls.Add(this.GroupBox2);
            this.Controls.Add(this.GroupBox1);
            this.Controls.Add(this.FlowLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "ReportContents";
            this.Text = "Prepare Data for Word Report";
            this.Load += new System.EventHandler(this.ReportContents_Load);
            this.GroupBox2.ResumeLayout(false);
            this.GroupBox2.PerformLayout();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.TextBox CaptionText;
        internal System.Windows.Forms.Label Label12;
        internal System.Windows.Forms.RadioButton optTbl;
        internal System.Windows.Forms.RadioButton optCalcTbl;
        internal System.Windows.Forms.RadioButton optCalc;
        internal System.Windows.Forms.Button DeleteBox;
        internal System.Windows.Forms.ComboBox cmbTableList;
        internal System.Windows.Forms.TextBox SummaryParam;
        internal System.Windows.Forms.Button btnPictures;
        internal System.Windows.Forms.TextBox PictureID;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.TextBox CriticalItem;
        internal System.Windows.Forms.TextBox Parameters;
        internal System.Windows.Forms.TextBox CaptionBox;
        internal System.Windows.Forms.Label Label13;
        internal System.Windows.Forms.TextBox SectionBox;
        internal System.Windows.Forms.RadioButton OptionCharts;
        internal System.Windows.Forms.Label Label8;
        internal System.Windows.Forms.Label Label7;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.RadioButton OptionPictures;
        internal System.Windows.Forms.Button BtnAddtoList;
        internal System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel1;
        internal System.Windows.Forms.Label Label9;
        internal System.Windows.Forms.ListBox ListBoxGeometry;
        internal System.Windows.Forms.ToolTip ToolTipTbl;
        internal System.Windows.Forms.Label Label10;
        internal System.Windows.Forms.Button BtnClose;
        internal System.Windows.Forms.Button AddSelectedBox;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Button BtnPicturesList;
        internal System.Windows.Forms.GroupBox GroupBox2;
        internal System.Windows.Forms.GroupBox GroupBox1;
    }
}