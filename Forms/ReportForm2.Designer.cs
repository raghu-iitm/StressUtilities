namespace StressUtilities.Forms
{
    partial class ReportForm2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportForm2));
            this.BtnClose = new System.Windows.Forms.Button();
            this.BtnExport = new System.Windows.Forms.Button();
            this.FileBox = new System.Windows.Forms.TextBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.BrowseIcon = new System.Windows.Forms.PictureBox();
            this.PathBox = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.BrowseIcon)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnClose
            // 
            this.BtnClose.Location = new System.Drawing.Point(111, 113);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(175, 37);
            this.BtnClose.TabIndex = 9;
            this.BtnClose.Text = "&Close";
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // BtnExport
            // 
            this.BtnExport.Location = new System.Drawing.Point(362, 113);
            this.BtnExport.Name = "BtnExport";
            this.BtnExport.Size = new System.Drawing.Size(175, 37);
            this.BtnExport.TabIndex = 10;
            this.BtnExport.Text = "Write Report";
            this.BtnExport.UseVisualStyleBackColor = true;
            this.BtnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // FileBox
            // 
            this.FileBox.Location = new System.Drawing.Point(111, 66);
            this.FileBox.Name = "FileBox";
            this.FileBox.Size = new System.Drawing.Size(340, 22);
            this.FileBox.TabIndex = 7;
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(457, 71);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(131, 17);
            this.Label3.TabIndex = 4;
            this.Label3.Text = "(Without Extension)";
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(13, 69);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(97, 17);
            this.Label2.TabIndex = 5;
            this.Label2.Text = "Report Name*";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(13, 27);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(42, 17);
            this.Label1.TabIndex = 6;
            this.Label1.Text = "Path*";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.BrowseIcon);
            this.panel1.Controls.Add(this.PathBox);
            this.panel1.Location = new System.Drawing.Point(102, 19);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(486, 30);
            this.panel1.TabIndex = 12;
            // 
            // BrowseIcon
            // 
            this.BrowseIcon.Image = global::StressUtilities.Properties.Resources.FolderBottomPanel;
            this.BrowseIcon.Location = new System.Drawing.Point(462, 5);
            this.BrowseIcon.Name = "BrowseIcon";
            this.BrowseIcon.Size = new System.Drawing.Size(20, 20);
            this.BrowseIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.BrowseIcon.TabIndex = 10;
            this.BrowseIcon.TabStop = false;
            this.BrowseIcon.Click += new System.EventHandler(this.BrowseIcon_Click);
            // 
            // PathBox
            // 
            this.PathBox.Location = new System.Drawing.Point(9, 4);
            this.PathBox.Name = "PathBox";
            this.PathBox.Size = new System.Drawing.Size(474, 22);
            this.PathBox.TabIndex = 9;
            // 
            // ReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(601, 174);
            this.Controls.Add(this.BtnClose);
            this.Controls.Add(this.BtnExport);
            this.Controls.Add(this.FileBox);
            this.Controls.Add(this.Label3);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ReportForm";
            this.Text = "Report Utility";
            this.Load += new System.EventHandler(this.ReportForm_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.BrowseIcon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        internal System.Windows.Forms.Button BtnClose;
        internal System.Windows.Forms.Button BtnExport;
        internal System.Windows.Forms.TextBox FileBox;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Label Label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox BrowseIcon;
        internal System.Windows.Forms.TextBox PathBox;
    }
}