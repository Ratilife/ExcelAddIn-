namespace ExcelAddInЭкспортДанных
{
    partial class ExportXlsxToDF
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
            this.txtRange = new System.Windows.Forms.TextBox();
            this.rdBook = new System.Windows.Forms.RadioButton();
            this.rbActiveSheet = new System.Windows.Forms.RadioButton();
            this.rbRange = new System.Windows.Forms.RadioButton();
            this.gBWhatExpounding = new System.Windows.Forms.GroupBox();
            this.btnSelectRange = new System.Windows.Forms.Button();
            this.chOpen = new System.Windows.Forms.CheckBox();
            this.cmbSaveAs = new System.Windows.Forms.ComboBox();
            this.gbSaveAs = new System.Windows.Forms.GroupBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.chHTMLOneDoc = new System.Windows.Forms.CheckBox();
            this.gBWhatExpounding.SuspendLayout();
            this.gbSaveAs.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtRange
            // 
            this.txtRange.Location = new System.Drawing.Point(188, 39);
            this.txtRange.Name = "txtRange";
            this.txtRange.Size = new System.Drawing.Size(229, 20);
            this.txtRange.TabIndex = 3;
            this.txtRange.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRange_KeyDown);
            // 
            // rdBook
            // 
            this.rdBook.AutoSize = true;
            this.rdBook.Location = new System.Drawing.Point(22, 105);
            this.rdBook.Name = "rdBook";
            this.rdBook.Size = new System.Drawing.Size(189, 17);
            this.rdBook.TabIndex = 2;
            this.rdBook.Text = "Все рабочие листы в этой книге";
            this.rdBook.UseVisualStyleBackColor = true;
            this.rdBook.CheckedChanged += new System.EventHandler(this.rdBook_CheckedChanged);
            // 
            // rbActiveSheet
            // 
            this.rbActiveSheet.AutoSize = true;
            this.rbActiveSheet.Checked = true;
            this.rbActiveSheet.Location = new System.Drawing.Point(22, 72);
            this.rbActiveSheet.Name = "rbActiveSheet";
            this.rbActiveSheet.Size = new System.Drawing.Size(118, 17);
            this.rbActiveSheet.TabIndex = 1;
            this.rbActiveSheet.TabStop = true;
            this.rbActiveSheet.Text = "Этот рабочий лист";
            this.rbActiveSheet.UseVisualStyleBackColor = true;
            // 
            // rbRange
            // 
            this.rbRange.AutoSize = true;
            this.rbRange.Location = new System.Drawing.Point(22, 39);
            this.rbRange.Name = "rbRange";
            this.rbRange.Size = new System.Drawing.Size(138, 17);
            this.rbRange.TabIndex = 0;
            this.rbRange.Text = "Выбранный диапазон:";
            this.rbRange.UseVisualStyleBackColor = true;
            // 
            // gBWhatExpounding
            // 
            this.gBWhatExpounding.Controls.Add(this.btnSelectRange);
            this.gBWhatExpounding.Controls.Add(this.txtRange);
            this.gBWhatExpounding.Controls.Add(this.rdBook);
            this.gBWhatExpounding.Controls.Add(this.rbActiveSheet);
            this.gBWhatExpounding.Controls.Add(this.rbRange);
            this.gBWhatExpounding.Location = new System.Drawing.Point(12, 12);
            this.gBWhatExpounding.Name = "gBWhatExpounding";
            this.gBWhatExpounding.Size = new System.Drawing.Size(467, 151);
            this.gBWhatExpounding.TabIndex = 2;
            this.gBWhatExpounding.TabStop = false;
            this.gBWhatExpounding.Text = "Выберете для экспорта:";
            // 
            // btnSelectRange
            // 
            this.btnSelectRange.Location = new System.Drawing.Point(424, 39);
            this.btnSelectRange.Name = "btnSelectRange";
            this.btnSelectRange.Size = new System.Drawing.Size(32, 23);
            this.btnSelectRange.TabIndex = 4;
            this.btnSelectRange.Text = "ОК";
            this.btnSelectRange.UseVisualStyleBackColor = true;
            // 
            // chOpen
            // 
            this.chOpen.AutoSize = true;
            this.chOpen.Location = new System.Drawing.Point(12, 256);
            this.chOpen.Name = "chOpen";
            this.chOpen.Size = new System.Drawing.Size(182, 17);
            this.chOpen.TabIndex = 6;
            this.chOpen.Text = "Открыть файл после экспорта";
            this.chOpen.UseVisualStyleBackColor = true;
            // 
            // cmbSaveAs
            // 
            this.cmbSaveAs.FormattingEnabled = true;
            this.cmbSaveAs.Items.AddRange(new object[] {
            "Книга Excel 97-2003(*.xls)",
            "Книга Excel с поддержкой макрасов (*.xlsm)",
            "Текст Юникод (*.txt)",
            "XML- данные (*.xml)",
            "Веб-страница (*.html)",
            "PDF(*.pdf)",
            "JSON(*.json)"});
            this.cmbSaveAs.Location = new System.Drawing.Point(6, 19);
            this.cmbSaveAs.Name = "cmbSaveAs";
            this.cmbSaveAs.Size = new System.Drawing.Size(444, 21);
            this.cmbSaveAs.TabIndex = 7;
            this.cmbSaveAs.SelectedIndexChanged += new System.EventHandler(this.cmbSaveAs_SelectedIndexChanged);
            // 
            // gbSaveAs
            // 
            this.gbSaveAs.Controls.Add(this.cmbSaveAs);
            this.gbSaveAs.Location = new System.Drawing.Point(12, 193);
            this.gbSaveAs.Name = "gbSaveAs";
            this.gbSaveAs.Size = new System.Drawing.Size(456, 57);
            this.gbSaveAs.TabIndex = 8;
            this.gbSaveAs.TabStop = false;
            this.gbSaveAs.Text = "Сохранить как:";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(387, 273);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 10;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(296, 273);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 9;
            this.btnOK.Text = "ОК";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // chHTMLOneDoc
            // 
            this.chHTMLOneDoc.AutoSize = true;
            this.chHTMLOneDoc.Location = new System.Drawing.Point(34, 170);
            this.chHTMLOneDoc.Name = "chHTMLOneDoc";
            this.chHTMLOneDoc.Size = new System.Drawing.Size(242, 17);
            this.chHTMLOneDoc.TabIndex = 11;
            this.chHTMLOneDoc.Text = "Экспорт книги в HTML одним документом";
            this.chHTMLOneDoc.UseVisualStyleBackColor = true;
            this.chHTMLOneDoc.CheckedChanged += new System.EventHandler(this.chHTMLOneDoc_CheckedChanged);
            // 
            // ExportXlsxToDF
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(486, 321);
            this.Controls.Add(this.chHTMLOneDoc);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.gbSaveAs);
            this.Controls.Add(this.chOpen);
            this.Controls.Add(this.gBWhatExpounding);
            this.Name = "ExportXlsxToDF";
            this.Text = "Экспорт листов";
            this.Load += new System.EventHandler(this.ExportXlsxToDF_Load);
            this.gBWhatExpounding.ResumeLayout(false);
            this.gBWhatExpounding.PerformLayout();
            this.gbSaveAs.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtRange;
        private System.Windows.Forms.RadioButton rdBook;
        private System.Windows.Forms.RadioButton rbActiveSheet;
        private System.Windows.Forms.RadioButton rbRange;
        private System.Windows.Forms.GroupBox gBWhatExpounding;
        private System.Windows.Forms.Button btnSelectRange;
        private System.Windows.Forms.CheckBox chOpen;
        private System.Windows.Forms.ComboBox cmbSaveAs;
        private System.Windows.Forms.GroupBox gbSaveAs;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.CheckBox chHTMLOneDoc;
    }
}