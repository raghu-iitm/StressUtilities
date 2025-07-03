namespace StressUtilities.Forms
{
    /*
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

    partial class SettingsControl
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
            this.WorkingDirectoryBox = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.BrowseBox = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.HDF5RowBox = new System.Windows.Forms.TextBox();
            this.BtnApply = new System.Windows.Forms.Button();
            this.unitOptionBox = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.BrowseBox)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(14, 9);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Working Directory";
            // 
            // WorkingDirectoryBox
            // 
            this.WorkingDirectoryBox.Location = new System.Drawing.Point(16, 34);
            this.WorkingDirectoryBox.Margin = new System.Windows.Forms.Padding(2);
            this.WorkingDirectoryBox.Name = "WorkingDirectoryBox";
            this.WorkingDirectoryBox.Size = new System.Drawing.Size(193, 20);
            this.WorkingDirectoryBox.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.Location = new System.Drawing.Point(10, 26);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(200, 32);
            this.panel1.TabIndex = 2;
            // 
            // BrowseBox
            // 
            this.BrowseBox.Image = global::StressUtilities.Properties.Resources.Folder;
            this.BrowseBox.Location = new System.Drawing.Point(189, 38);
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
            this.label2.Location = new System.Drawing.Point(14, 75);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(112, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Maximum HDF5 Rows";
            // 
            // HDF5RowBox
            // 
            this.HDF5RowBox.Location = new System.Drawing.Point(16, 91);
            this.HDF5RowBox.Margin = new System.Windows.Forms.Padding(2);
            this.HDF5RowBox.Name = "HDF5RowBox";
            this.HDF5RowBox.Size = new System.Drawing.Size(192, 20);
            this.HDF5RowBox.TabIndex = 4;
            this.HDF5RowBox.Text = "200000";
            // 
            // BtnApply
            // 
            this.BtnApply.Location = new System.Drawing.Point(62, 182);
            this.BtnApply.Margin = new System.Windows.Forms.Padding(2);
            this.BtnApply.Name = "BtnApply";
            this.BtnApply.Size = new System.Drawing.Size(101, 26);
            this.BtnApply.TabIndex = 5;
            this.BtnApply.Text = "&Apply";
            this.BtnApply.UseVisualStyleBackColor = true;
            this.BtnApply.Click += new System.EventHandler(this.BtnApply_Click);
            // 
            // unitOptionBox
            // 
            this.unitOptionBox.AutoSize = true;
            this.unitOptionBox.Location = new System.Drawing.Point(16, 142);
            this.unitOptionBox.Margin = new System.Windows.Forms.Padding(2);
            this.unitOptionBox.Name = "unitOptionBox";
            this.unitOptionBox.Size = new System.Drawing.Size(164, 17);
            this.unitOptionBox.TabIndex = 6;
            this.unitOptionBox.Text = "Show Units in the Calculation";
            this.unitOptionBox.UseVisualStyleBackColor = true;
            // 
            // SettingsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.Controls.Add(this.unitOptionBox);
            this.Controls.Add(this.BtnApply);
            this.Controls.Add(this.HDF5RowBox);
            this.Controls.Add(this.BrowseBox);
            this.Controls.Add(this.WorkingDirectoryBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "SettingsControl";
            this.Size = new System.Drawing.Size(218, 401);
            ((System.ComponentModel.ISupportInitialize)(this.BrowseBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox WorkingDirectoryBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox BrowseBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox HDF5RowBox;
        private System.Windows.Forms.Button BtnApply;
        private System.Windows.Forms.CheckBox unitOptionBox;
    }
}
