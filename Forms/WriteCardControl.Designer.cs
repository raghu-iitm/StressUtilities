namespace StressUtilities.Forms
{
    partial class WriteCardControl
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
            this.label2 = new System.Windows.Forms.Label();
            this.BtnExport = new System.Windows.Forms.Button();
            this.AllCardsList = new System.Windows.Forms.ListBox();
            this.AddCards = new System.Windows.Forms.Button();
            this.FileBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.CardNameBox = new System.Windows.Forms.TextBox();
            this.btnAutofill = new System.Windows.Forms.Button();
            this.BrowseBox = new System.Windows.Forms.PictureBox();
            this.PathBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.OSBox = new System.Windows.Forms.ComboBox();
            this.CardFormatBox = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.BrowseBox)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(14, 124);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Nastran Card";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 174);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Selected Cards*";
            // 
            // BtnExport
            // 
            this.BtnExport.Location = new System.Drawing.Point(104, 398);
            this.BtnExport.Margin = new System.Windows.Forms.Padding(2);
            this.BtnExport.Name = "BtnExport";
            this.BtnExport.Size = new System.Drawing.Size(76, 26);
            this.BtnExport.TabIndex = 5;
            this.BtnExport.Text = "&Write Cards";
            this.BtnExport.UseVisualStyleBackColor = true;
            this.BtnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // AllCardsList
            // 
            this.AllCardsList.FormattingEnabled = true;
            this.AllCardsList.Location = new System.Drawing.Point(15, 190);
            this.AllCardsList.Name = "AllCardsList";
            this.AllCardsList.Size = new System.Drawing.Size(165, 95);
            this.AllCardsList.TabIndex = 6;
            this.AllCardsList.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.AllCardsList_MouseDoubleClick);
            // 
            // AddCards
            // 
            this.AddCards.Location = new System.Drawing.Point(138, 138);
            this.AddCards.Margin = new System.Windows.Forms.Padding(2);
            this.AddCards.Name = "AddCards";
            this.AddCards.Size = new System.Drawing.Size(42, 22);
            this.AddCards.TabIndex = 5;
            this.AddCards.Text = "&Add";
            this.AddCards.UseVisualStyleBackColor = true;
            this.AddCards.Click += new System.EventHandler(this.AddCards_Click);
            // 
            // FileBox
            // 
            this.FileBox.Location = new System.Drawing.Point(15, 87);
            this.FileBox.Margin = new System.Windows.Forms.Padding(2);
            this.FileBox.Name = "FileBox";
            this.FileBox.Size = new System.Drawing.Size(164, 20);
            this.FileBox.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 69);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "File Name*";
            // 
            // CardNameBox
            // 
            this.CardNameBox.Location = new System.Drawing.Point(15, 140);
            this.CardNameBox.Name = "CardNameBox";
            this.CardNameBox.Size = new System.Drawing.Size(118, 20);
            this.CardNameBox.TabIndex = 9;
            this.CardNameBox.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // btnAutofill
            // 
            this.btnAutofill.Location = new System.Drawing.Point(15, 398);
            this.btnAutofill.Margin = new System.Windows.Forms.Padding(2);
            this.btnAutofill.Name = "btnAutofill";
            this.btnAutofill.Size = new System.Drawing.Size(74, 26);
            this.btnAutofill.TabIndex = 5;
            this.btnAutofill.Text = "&Auto Fill";
            this.btnAutofill.UseVisualStyleBackColor = true;
            this.btnAutofill.Click += new System.EventHandler(this.btnAutofill_Click);
            // 
            // BrowseBox
            // 
            this.BrowseBox.Image = global::StressUtilities.Properties.Resources.Folder;
            this.BrowseBox.Location = new System.Drawing.Point(153, 11);
            this.BrowseBox.Margin = new System.Windows.Forms.Padding(2);
            this.BrowseBox.Name = "BrowseBox";
            this.BrowseBox.Size = new System.Drawing.Size(18, 13);
            this.BrowseBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.BrowseBox.TabIndex = 13;
            this.BrowseBox.TabStop = false;
            this.BrowseBox.Click += new System.EventHandler(this.BrowseBox_Click);
            // 
            // PathBox
            // 
            this.PathBox.Location = new System.Drawing.Point(7, 8);
            this.PathBox.Margin = new System.Windows.Forms.Padding(2);
            this.PathBox.Name = "PathBox";
            this.PathBox.Size = new System.Drawing.Size(165, 20);
            this.PathBox.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(14, 8);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(98, 19);
            this.label4.TabIndex = 10;
            this.label4.Text = "Path*";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.BrowseBox);
            this.panel1.Controls.Add(this.PathBox);
            this.panel1.Location = new System.Drawing.Point(8, 23);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(174, 37);
            this.panel1.TabIndex = 12;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 295);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(60, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Target OS*";
            // 
            // OSBox
            // 
            this.OSBox.FormattingEnabled = true;
            this.OSBox.Items.AddRange(new object[] {
            "Unix/Linux",
            "Windows"});
            this.OSBox.Location = new System.Drawing.Point(15, 312);
            this.OSBox.Name = "OSBox";
            this.OSBox.Size = new System.Drawing.Size(164, 21);
            this.OSBox.TabIndex = 13;
            // 
            // CardFormatBox
            // 
            this.CardFormatBox.FormattingEnabled = true;
            this.CardFormatBox.Items.AddRange(new object[] {
            "FREE",
            "SMALL",
            "LARGE"});
            this.CardFormatBox.Location = new System.Drawing.Point(15, 360);
            this.CardFormatBox.Name = "CardFormatBox";
            this.CardFormatBox.Size = new System.Drawing.Size(164, 21);
            this.CardFormatBox.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 345);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(64, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "Card Format";
            // 
            // WriteCardControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.Controls.Add(this.CardFormatBox);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.OSBox);
            this.Controls.Add(this.CardNameBox);
            this.Controls.Add(this.FileBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.AllCardsList);
            this.Controls.Add(this.btnAutofill);
            this.Controls.Add(this.AddCards);
            this.Controls.Add(this.BtnExport);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label4);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "WriteCardControl";
            this.Size = new System.Drawing.Size(197, 439);
            ((System.ComponentModel.ISupportInitialize)(this.BrowseBox)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button BtnExport;
        private System.Windows.Forms.ListBox AllCardsList;
        private System.Windows.Forms.Button AddCards;
        private System.Windows.Forms.TextBox FileBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox CardNameBox;
        private System.Windows.Forms.Button btnAutofill;
        private System.Windows.Forms.PictureBox BrowseBox;
        private System.Windows.Forms.TextBox PathBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox OSBox;
        private System.Windows.Forms.ComboBox CardFormatBox;
        private System.Windows.Forms.Label label6;
    }
}
