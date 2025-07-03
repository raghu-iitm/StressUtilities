namespace StressUtilities.Forms
{

 /*
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

    partial class FormReadPunch
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormReadPunch));
            this.BtnExtractPunch = new System.Windows.Forms.Button();
            this.SolNameBox = new System.Windows.Forms.TextBox();
            this.Label1 = new System.Windows.Forms.Label();
            this.SelectionGroup = new System.Windows.Forms.GroupBox();
            this.TextBox2 = new System.Windows.Forms.TextBox();
            this.Label5 = new System.Windows.Forms.Label();
            this.EntityList = new System.Windows.Forms.TextBox();
            this.TargetEntityBox = new System.Windows.Forms.ComboBox();
            this.RequestBox = new System.Windows.Forms.ComboBox();
            this.ComboBox2 = new System.Windows.Forms.ComboBox();
            this.SolutionTypeBox = new System.Windows.Forms.ComboBox();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.BtnClose = new System.Windows.Forms.Button();
            this.BtnBrowse = new System.Windows.Forms.Button();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.FileListBox = new System.Windows.Forms.ListBox();
            this.SelectionGroup.SuspendLayout();
            this.GroupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnExtractPunch
            // 
            this.BtnExtractPunch.Location = new System.Drawing.Point(359, 554);
            this.BtnExtractPunch.Margin = new System.Windows.Forms.Padding(4);
            this.BtnExtractPunch.Name = "BtnExtractPunch";
            this.BtnExtractPunch.Size = new System.Drawing.Size(139, 31);
            this.BtnExtractPunch.TabIndex = 14;
            this.BtnExtractPunch.Text = "Extract Results";
            this.BtnExtractPunch.UseVisualStyleBackColor = true;
            this.BtnExtractPunch.Click += new System.EventHandler(this.BtnExtractPunch_Click);
            // 
            // SolNameBox
            // 
            this.SolNameBox.Enabled = false;
            this.SolNameBox.Location = new System.Drawing.Point(265, 25);
            this.SolNameBox.Name = "SolNameBox";
            this.SolNameBox.Size = new System.Drawing.Size(288, 22);
            this.SolNameBox.TabIndex = 10;
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(-2, 272);
            this.Label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(69, 17);
            this.Label1.TabIndex = 2;
            this.Label1.Text = "Entity List";
            // 
            // SelectionGroup
            // 
            this.SelectionGroup.Controls.Add(this.SolNameBox);
            this.SelectionGroup.Controls.Add(this.Label1);
            this.SelectionGroup.Controls.Add(this.TextBox2);
            this.SelectionGroup.Controls.Add(this.Label5);
            this.SelectionGroup.Controls.Add(this.EntityList);
            this.SelectionGroup.Controls.Add(this.TargetEntityBox);
            this.SelectionGroup.Controls.Add(this.RequestBox);
            this.SelectionGroup.Controls.Add(this.ComboBox2);
            this.SelectionGroup.Controls.Add(this.SolutionTypeBox);
            this.SelectionGroup.Controls.Add(this.Label2);
            this.SelectionGroup.Controls.Add(this.Label6);
            this.SelectionGroup.Controls.Add(this.Label3);
            this.SelectionGroup.Controls.Add(this.Label4);
            this.SelectionGroup.Location = new System.Drawing.Point(8, 225);
            this.SelectionGroup.Margin = new System.Windows.Forms.Padding(4);
            this.SelectionGroup.Name = "SelectionGroup";
            this.SelectionGroup.Padding = new System.Windows.Forms.Padding(4);
            this.SelectionGroup.Size = new System.Drawing.Size(574, 321);
            this.SelectionGroup.TabIndex = 12;
            this.SelectionGroup.TabStop = false;
            this.SelectionGroup.Text = "Selections";
            // 
            // TextBox2
            // 
            this.TextBox2.Location = new System.Drawing.Point(119, 223);
            this.TextBox2.Margin = new System.Windows.Forms.Padding(4);
            this.TextBox2.Name = "TextBox2";
            this.TextBox2.Size = new System.Drawing.Size(435, 22);
            this.TextBox2.TabIndex = 7;
            // 
            // Label5
            // 
            this.Label5.AutoSize = true;
            this.Label5.Location = new System.Drawing.Point(-2, 223);
            this.Label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(87, 17);
            this.Label5.TabIndex = 2;
            this.Label5.Text = "Subcase IDs";
            // 
            // EntityList
            // 
            this.EntityList.Location = new System.Drawing.Point(119, 271);
            this.EntityList.Margin = new System.Windows.Forms.Padding(4);
            this.EntityList.Name = "EntityList";
            this.EntityList.Size = new System.Drawing.Size(435, 22);
            this.EntityList.TabIndex = 8;
            // 
            // TargetEntityBox
            // 
            this.TargetEntityBox.FormattingEnabled = true;
            this.TargetEntityBox.Location = new System.Drawing.Point(119, 123);
            this.TargetEntityBox.Margin = new System.Windows.Forms.Padding(4);
            this.TargetEntityBox.Name = "TargetEntityBox";
            this.TargetEntityBox.Size = new System.Drawing.Size(435, 24);
            this.TargetEntityBox.TabIndex = 5;
            this.TargetEntityBox.SelectedIndexChanged += new System.EventHandler(this.SolutionTypeBox_SelectedIndexChanged);
            this.TargetEntityBox.SelectionChangeCommitted += new System.EventHandler(this.TargetEntityBox_SelectionChangeCommitted);
            // 
            // RequestBox
            // 
            this.RequestBox.FormattingEnabled = true;
            this.RequestBox.Location = new System.Drawing.Point(119, 73);
            this.RequestBox.Margin = new System.Windows.Forms.Padding(4);
            this.RequestBox.Name = "RequestBox";
            this.RequestBox.Size = new System.Drawing.Size(435, 24);
            this.RequestBox.TabIndex = 4;
            this.RequestBox.SelectionChangeCommitted += new System.EventHandler(this.RequestBox_SelectionChangeCommitted);
            // 
            // ComboBox2
            // 
            this.ComboBox2.FormattingEnabled = true;
            this.ComboBox2.Location = new System.Drawing.Point(119, 173);
            this.ComboBox2.Margin = new System.Windows.Forms.Padding(4);
            this.ComboBox2.Name = "ComboBox2";
            this.ComboBox2.Size = new System.Drawing.Size(435, 24);
            this.ComboBox2.TabIndex = 6;
            // 
            // SolutionTypeBox
            // 
            this.SolutionTypeBox.FormattingEnabled = true;
            this.SolutionTypeBox.Location = new System.Drawing.Point(119, 23);
            this.SolutionTypeBox.Margin = new System.Windows.Forms.Padding(4);
            this.SolutionTypeBox.Name = "SolutionTypeBox";
            this.SolutionTypeBox.Size = new System.Drawing.Size(139, 24);
            this.SolutionTypeBox.TabIndex = 3;
            this.SolutionTypeBox.SelectedIndexChanged += new System.EventHandler(this.SolutionTypeBox_SelectedIndexChanged);
            this.SolutionTypeBox.SelectionChangeCommitted += new System.EventHandler(this.SolutionTypeBox_SelectionChangeCommitted);
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(-1, 27);
            this.Label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(95, 17);
            this.Label2.TabIndex = 2;
            this.Label2.Text = "Solution Type";
            // 
            // Label6
            // 
            this.Label6.AutoSize = true;
            this.Label6.Location = new System.Drawing.Point(-2, 125);
            this.Label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(94, 17);
            this.Label6.TabIndex = 2;
            this.Label6.Text = "Target Entity*";
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(-2, 174);
            this.Label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(95, 17);
            this.Label3.TabIndex = 2;
            this.Label3.Text = "Element Type";
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(-2, 76);
            this.Label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(110, 17);
            this.Label4.TabIndex = 2;
            this.Label4.Text = "Result Request*";
            // 
            // BtnClose
            // 
            this.BtnClose.Location = new System.Drawing.Point(64, 554);
            this.BtnClose.Margin = new System.Windows.Forms.Padding(4);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(139, 31);
            this.BtnClose.TabIndex = 13;
            this.BtnClose.Text = "&Close";
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // BtnBrowse
            // 
            this.BtnBrowse.Location = new System.Drawing.Point(173, 163);
            this.BtnBrowse.Margin = new System.Windows.Forms.Padding(4);
            this.BtnBrowse.Name = "BtnBrowse";
            this.BtnBrowse.Size = new System.Drawing.Size(216, 31);
            this.BtnBrowse.TabIndex = 2;
            this.BtnBrowse.Text = "Browse";
            this.BtnBrowse.UseVisualStyleBackColor = true;
            this.BtnBrowse.Click += new System.EventHandler(this.BtnBrowse_Click);
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.BtnBrowse);
            this.GroupBox1.Controls.Add(this.FileListBox);
            this.GroupBox1.Location = new System.Drawing.Point(8, 12);
            this.GroupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.GroupBox1.Size = new System.Drawing.Size(574, 205);
            this.GroupBox1.TabIndex = 11;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "File List";
            // 
            // FileListBox
            // 
            this.FileListBox.FormattingEnabled = true;
            this.FileListBox.ItemHeight = 16;
            this.FileListBox.Location = new System.Drawing.Point(9, 23);
            this.FileListBox.Margin = new System.Windows.Forms.Padding(4);
            this.FileListBox.Name = "FileListBox";
            this.FileListBox.Size = new System.Drawing.Size(544, 132);
            this.FileListBox.TabIndex = 1;
            this.FileListBox.DoubleClick += new System.EventHandler(this.FileListBox_DoubleClick);
            // 
            // FormReadPunch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(590, 596);
            this.Controls.Add(this.BtnExtractPunch);
            this.Controls.Add(this.SelectionGroup);
            this.Controls.Add(this.BtnClose);
            this.Controls.Add(this.GroupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormReadPunch";
            this.Text = "Read Punch Files";
            this.Load += new System.EventHandler(this.FormReadPunch_Load);
            this.SelectionGroup.ResumeLayout(false);
            this.SelectionGroup.PerformLayout();
            this.GroupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Button BtnExtractPunch;
        internal System.Windows.Forms.TextBox SolNameBox;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.GroupBox SelectionGroup;
        internal System.Windows.Forms.TextBox TextBox2;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.TextBox EntityList;
        internal System.Windows.Forms.ComboBox TargetEntityBox;
        internal System.Windows.Forms.ComboBox RequestBox;
        internal System.Windows.Forms.ComboBox ComboBox2;
        internal System.Windows.Forms.ComboBox SolutionTypeBox;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Button BtnClose;
        internal System.Windows.Forms.Button BtnBrowse;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.ListBox FileListBox;
    }
}