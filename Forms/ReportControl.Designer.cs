namespace StressUtilities.Forms
{
    /*
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/
    partial class ReportControl
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
            this.label1 = new System.Windows.Forms.Label();
            this.PathBox = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.BrowseBox = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.FileBox = new System.Windows.Forms.TextBox();
            this.BtnExport = new System.Windows.Forms.Button();
            this.FillForm = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.BrowseBox)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(12, 16);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Path*";
            // 
            // PathBox
            // 
            this.PathBox.Location = new System.Drawing.Point(14, 41);
            this.PathBox.Margin = new System.Windows.Forms.Padding(2);
            this.PathBox.Name = "PathBox";
            this.PathBox.Size = new System.Drawing.Size(193, 20);
            this.PathBox.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.Location = new System.Drawing.Point(8, 33);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(200, 32);
            this.panel1.TabIndex = 2;
            // 
            // BrowseBox
            // 
            this.BrowseBox.Image = global::StressUtilities.Properties.Resources.Folder;
            this.BrowseBox.Location = new System.Drawing.Point(187, 45);
            this.BrowseBox.Margin = new System.Windows.Forms.Padding(2);
            this.BrowseBox.Name = "BrowseBox";
            this.BrowseBox.Size = new System.Drawing.Size(18, 13);
            this.BrowseBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.BrowseBox.TabIndex = 3;
            this.BrowseBox.TabStop = false;
            this.BrowseBox.Click += new System.EventHandler(this.BrowseBox_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 82);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(169, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Report Name (Without Extension)*";
            // 
            // FileBox
            // 
            this.FileBox.Location = new System.Drawing.Point(14, 98);
            this.FileBox.Margin = new System.Windows.Forms.Padding(2);
            this.FileBox.Name = "FileBox";
            this.FileBox.Size = new System.Drawing.Size(192, 20);
            this.FileBox.TabIndex = 4;
            // 
            // BtnExport
            // 
            this.BtnExport.Location = new System.Drawing.Point(126, 138);
            this.BtnExport.Margin = new System.Windows.Forms.Padding(2);
            this.BtnExport.Name = "BtnExport";
            this.BtnExport.Size = new System.Drawing.Size(80, 26);
            this.BtnExport.TabIndex = 5;
            this.BtnExport.Text = "&Write Report";
            this.BtnExport.UseVisualStyleBackColor = true;
            this.BtnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // FillForm
            // 
            this.FillForm.Location = new System.Drawing.Point(14, 138);
            this.FillForm.Margin = new System.Windows.Forms.Padding(2);
            this.FillForm.Name = "FillForm";
            this.FillForm.Size = new System.Drawing.Size(80, 26);
            this.FillForm.TabIndex = 5;
            this.FillForm.Text = "&Auto Fill";
            this.FillForm.UseVisualStyleBackColor = true;
            this.FillForm.Click += new System.EventHandler(this.FillForm_Click);
            // 
            // ReportControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.Controls.Add(this.FillForm);
            this.Controls.Add(this.BtnExport);
            this.Controls.Add(this.FileBox);
            this.Controls.Add(this.BrowseBox);
            this.Controls.Add(this.PathBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "ReportControl";
            this.Size = new System.Drawing.Size(218, 401);
            ((System.ComponentModel.ISupportInitialize)(this.BrowseBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox PathBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox BrowseBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox FileBox;
        private System.Windows.Forms.Button BtnExport;
        private System.Windows.Forms.Button FillForm;
    }
}
