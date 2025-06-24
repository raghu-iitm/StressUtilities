namespace StressUtilities.Forms
{
    partial class ImportF06Form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImportF06Form));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnImportF06 = new System.Windows.Forms.Button();
            this.FilesListBox = new System.Windows.Forms.ListBox();
            this.RequestBox = new System.Windows.Forms.ComboBox();
            this.ElementList = new System.Windows.Forms.TextBox();
            this.Label1 = new System.Windows.Forms.Label();
            this.btnF06Add = new System.Windows.Forms.Button();
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.GroupBox2.SuspendLayout();
            this.GroupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(218, 435);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(148, 32);
            this.btnClose.TabIndex = 10;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnImportF06
            // 
            this.btnImportF06.Location = new System.Drawing.Point(5, 435);
            this.btnImportF06.Name = "btnImportF06";
            this.btnImportF06.Size = new System.Drawing.Size(148, 32);
            this.btnImportF06.TabIndex = 9;
            this.btnImportF06.Text = "Import .f06 Results";
            this.btnImportF06.UseVisualStyleBackColor = true;
            this.btnImportF06.Click += new System.EventHandler(this.btnImportF06_Click);
            // 
            // FilesListBox
            // 
            this.FilesListBox.FormattingEnabled = true;
            this.FilesListBox.ItemHeight = 16;
            this.FilesListBox.Location = new System.Drawing.Point(6, 20);
            this.FilesListBox.Name = "FilesListBox";
            this.FilesListBox.Size = new System.Drawing.Size(356, 148);
            this.FilesListBox.TabIndex = 2;
            this.FilesListBox.DoubleClick += new System.EventHandler(this.FilesListBox_DoubleClick);
            // 
            // RequestBox
            // 
            this.RequestBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.RequestBox.FormattingEnabled = true;
            this.RequestBox.Location = new System.Drawing.Point(142, 264);
            this.RequestBox.Name = "RequestBox";
            this.RequestBox.Size = new System.Drawing.Size(225, 24);
            this.RequestBox.TabIndex = 11;
            this.RequestBox.SelectionChangeCommitted += new System.EventHandler(this.RequestBox_SelectionChangeCommitted);
            // 
            // ElementList
            // 
            this.ElementList.Location = new System.Drawing.Point(6, 22);
            this.ElementList.Multiline = true;
            this.ElementList.Name = "ElementList";
            this.ElementList.Size = new System.Drawing.Size(356, 74);
            this.ElementList.TabIndex = 4;
            this.ElementList.Text = "Elm ";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(7, 271);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(110, 17);
            this.Label1.TabIndex = 12;
            this.Label1.Text = "Result Request*";
            // 
            // btnF06Add
            // 
            this.btnF06Add.Location = new System.Drawing.Point(128, 185);
            this.btnF06Add.Name = "btnF06Add";
            this.btnF06Add.Size = new System.Drawing.Size(112, 26);
            this.btnF06Add.TabIndex = 6;
            this.btnF06Add.Text = "Add *.f06 Files";
            this.btnF06Add.UseVisualStyleBackColor = true;
            this.btnF06Add.Click += new System.EventHandler(this.btnF06Add_Click);
            // 
            // GroupBox2
            // 
            this.GroupBox2.Controls.Add(this.ElementList);
            this.GroupBox2.Location = new System.Drawing.Point(5, 308);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Size = new System.Drawing.Size(370, 104);
            this.GroupBox2.TabIndex = 14;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "Entity List";
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.FilesListBox);
            this.GroupBox1.Controls.Add(this.btnF06Add);
            this.GroupBox1.Location = new System.Drawing.Point(5, 13);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(370, 229);
            this.GroupBox1.TabIndex = 13;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "Nastran *.f06 File List";
            // 
            // ImportF06Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(381, 480);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnImportF06);
            this.Controls.Add(this.RequestBox);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.GroupBox2);
            this.Controls.Add(this.GroupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ImportF06Form";
            this.Text = "Import *.f06 Files";
            this.Load += new System.EventHandler(this.ImportF06Form_Load);
            this.GroupBox2.ResumeLayout(false);
            this.GroupBox2.PerformLayout();
            this.GroupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Button btnClose;
        internal System.Windows.Forms.Button btnImportF06;
        internal System.Windows.Forms.ListBox FilesListBox;
        internal System.Windows.Forms.ComboBox RequestBox;
        internal System.Windows.Forms.TextBox ElementList;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Button btnF06Add;
        internal System.Windows.Forms.GroupBox GroupBox2;
        internal System.Windows.Forms.GroupBox GroupBox1;
    }
}