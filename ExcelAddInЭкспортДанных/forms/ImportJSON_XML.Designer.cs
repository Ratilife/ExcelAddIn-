namespace ExcelAddInЭкспортДанных.forms
{
    partial class Import_JSON_XML
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
            this.cmbSaveAs = new System.Windows.Forms.ComboBox();
            this.gbSaveAs = new System.Windows.Forms.GroupBox();
            this.btCreate = new System.Windows.Forms.Button();
            this.gbСreate = new System.Windows.Forms.GroupBox();
            this.rbNewSheet = new System.Windows.Forms.RadioButton();
            this.rbActivSeheet = new System.Windows.Forms.RadioButton();
            this.labelFilePath = new System.Windows.Forms.Label();
            this.tbFilePath = new System.Windows.Forms.TextBox();
            this.butFilePath = new System.Windows.Forms.Button();
            this.gbSaveAs.SuspendLayout();
            this.gbСreate.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbSaveAs
            // 
            this.cmbSaveAs.FormattingEnabled = true;
            this.cmbSaveAs.Items.AddRange(new object[] {
            "XML- данные (*.xml)",
            "JSON(*.json)"});
            this.cmbSaveAs.Location = new System.Drawing.Point(6, 19);
            this.cmbSaveAs.Name = "cmbSaveAs";
            this.cmbSaveAs.Size = new System.Drawing.Size(373, 21);
            this.cmbSaveAs.TabIndex = 7;
            // 
            // gbSaveAs
            // 
            this.gbSaveAs.Controls.Add(this.cmbSaveAs);
            this.gbSaveAs.Location = new System.Drawing.Point(12, 79);
            this.gbSaveAs.Name = "gbSaveAs";
            this.gbSaveAs.Size = new System.Drawing.Size(385, 49);
            this.gbSaveAs.TabIndex = 9;
            this.gbSaveAs.TabStop = false;
            this.gbSaveAs.Text = "Импортировать из:";
            // 
            // btCreate
            // 
            this.btCreate.Location = new System.Drawing.Point(12, 209);
            this.btCreate.Name = "btCreate";
            this.btCreate.Size = new System.Drawing.Size(231, 23);
            this.btCreate.TabIndex = 11;
            this.btCreate.Text = "Создать";
            this.btCreate.UseVisualStyleBackColor = true;
            this.btCreate.Click += new System.EventHandler(this.btCreate_Click);
            // 
            // gbСreate
            // 
            this.gbСreate.Controls.Add(this.rbNewSheet);
            this.gbСreate.Controls.Add(this.rbActivSeheet);
            this.gbСreate.Location = new System.Drawing.Point(12, 134);
            this.gbСreate.Name = "gbСreate";
            this.gbСreate.Size = new System.Drawing.Size(349, 49);
            this.gbСreate.TabIndex = 10;
            this.gbСreate.TabStop = false;
            this.gbСreate.Text = "Создать:";
            // 
            // rbNewSheet
            // 
            this.rbNewSheet.AutoSize = true;
            this.rbNewSheet.Location = new System.Drawing.Point(173, 26);
            this.rbNewSheet.Name = "rbNewSheet";
            this.rbNewSheet.Size = new System.Drawing.Size(106, 17);
            this.rbNewSheet.TabIndex = 1;
            this.rbNewSheet.Text = "На новом листе";
            this.rbNewSheet.UseVisualStyleBackColor = true;
            // 
            // rbActivSeheet
            // 
            this.rbActivSeheet.AutoSize = true;
            this.rbActivSeheet.Checked = true;
            this.rbActivSeheet.Location = new System.Drawing.Point(7, 26);
            this.rbActivSeheet.Name = "rbActivSeheet";
            this.rbActivSeheet.Size = new System.Drawing.Size(123, 17);
            this.rbActivSeheet.TabIndex = 0;
            this.rbActivSeheet.TabStop = true;
            this.rbActivSeheet.Text = "На активном листе";
            this.rbActivSeheet.UseVisualStyleBackColor = true;
            // 
            // labelFilePath
            // 
            this.labelFilePath.AutoSize = true;
            this.labelFilePath.Location = new System.Drawing.Point(12, 44);
            this.labelFilePath.Name = "labelFilePath";
            this.labelFilePath.Size = new System.Drawing.Size(77, 13);
            this.labelFilePath.TabIndex = 12;
            this.labelFilePath.Text = "Путь к файлу:";
            // 
            // tbFilePath
            // 
            this.tbFilePath.Location = new System.Drawing.Point(96, 44);
            this.tbFilePath.Name = "tbFilePath";
            this.tbFilePath.Size = new System.Drawing.Size(265, 20);
            this.tbFilePath.TabIndex = 13;
            // 
            // butFilePath
            // 
            this.butFilePath.Location = new System.Drawing.Point(370, 44);
            this.butFilePath.Name = "butFilePath";
            this.butFilePath.Size = new System.Drawing.Size(37, 20);
            this.butFilePath.TabIndex = 14;
            this.butFilePath.Text = "...";
            this.butFilePath.UseVisualStyleBackColor = true;
            this.butFilePath.Click += new System.EventHandler(this.butFilePath_Click);
            // 
            // Import_JSON_XML
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(420, 242);
            this.Controls.Add(this.butFilePath);
            this.Controls.Add(this.tbFilePath);
            this.Controls.Add(this.labelFilePath);
            this.Controls.Add(this.btCreate);
            this.Controls.Add(this.gbСreate);
            this.Controls.Add(this.gbSaveAs);
            this.Name = "Import_JSON_XML";
            this.Text = "Импорт";
            this.Load += new System.EventHandler(this.Import_JSON_XML_Load);
            this.gbSaveAs.ResumeLayout(false);
            this.gbСreate.ResumeLayout(false);
            this.gbСreate.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbSaveAs;
        private System.Windows.Forms.GroupBox gbSaveAs;
        private System.Windows.Forms.Button btCreate;
        private System.Windows.Forms.GroupBox gbСreate;
        private System.Windows.Forms.RadioButton rbNewSheet;
        private System.Windows.Forms.RadioButton rbActivSeheet;
        private System.Windows.Forms.Label labelFilePath;
        private System.Windows.Forms.TextBox tbFilePath;
        private System.Windows.Forms.Button butFilePath;
    }
}