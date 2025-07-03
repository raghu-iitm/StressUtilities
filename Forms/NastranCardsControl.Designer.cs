namespace StressUtilities.Forms
{
    /*
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

    partial class NastranCardsControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.BtnApply = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.bulkEntryBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.loadTypeBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.nastranCardsBox = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // BtnApply
            // 
            this.BtnApply.Location = new System.Drawing.Point(53, 175);
            this.BtnApply.Margin = new System.Windows.Forms.Padding(2);
            this.BtnApply.Name = "BtnApply";
            this.BtnApply.Size = new System.Drawing.Size(82, 26);
            this.BtnApply.TabIndex = 5;
            this.BtnApply.Text = "&Apply";
            this.BtnApply.UseVisualStyleBackColor = true;
            this.BtnApply.Click += new System.EventHandler(this.BtnApply_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(81, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Bulk Data Entry";
            // 
            // bulkEntryBox
            // 
            this.bulkEntryBox.FormattingEnabled = true;
            this.bulkEntryBox.Location = new System.Drawing.Point(12, 35);
            this.bulkEntryBox.Name = "bulkEntryBox";
            this.bulkEntryBox.Size = new System.Drawing.Size(164, 21);
            this.bulkEntryBox.TabIndex = 7;
            this.bulkEntryBox.SelectedIndexChanged += new System.EventHandler(this.blkEntryBox_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Load Type";
            // 
            // loadTypeBox
            // 
            this.loadTypeBox.FormattingEnabled = true;
            this.loadTypeBox.Location = new System.Drawing.Point(12, 85);
            this.loadTypeBox.Name = "loadTypeBox";
            this.loadTypeBox.Size = new System.Drawing.Size(164, 21);
            this.loadTypeBox.TabIndex = 7;
            this.loadTypeBox.SelectedIndexChanged += new System.EventHandler(this.loadTypeBox_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 121);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Card";
            // 
            // nastranCardsBox
            // 
            this.nastranCardsBox.FormattingEnabled = true;
            this.nastranCardsBox.Location = new System.Drawing.Point(12, 135);
            this.nastranCardsBox.Name = "nastranCardsBox";
            this.nastranCardsBox.Size = new System.Drawing.Size(164, 21);
            this.nastranCardsBox.TabIndex = 7;
            // 
            // NastranCardsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.Controls.Add(this.nastranCardsBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.loadTypeBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.bulkEntryBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.BtnApply);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "NastranCardsControl";
            this.Size = new System.Drawing.Size(192, 401);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button BtnApply;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox bulkEntryBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox loadTypeBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox nastranCardsBox;
    }
}
