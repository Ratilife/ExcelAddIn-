namespace ExcelAddInЭкспортДанных
{
    partial class ExportXlsxToCsv
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
            this.gBWhatExpounding = new System.Windows.Forms.GroupBox();
            this.btnSelectRange = new System.Windows.Forms.Button();
            this.txtRange = new System.Windows.Forms.TextBox();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.grFileSettings = new System.Windows.Forms.GroupBox();
            this.labSeparator = new System.Windows.Forms.Label();
            this.labEncoding = new System.Windows.Forms.Label();
            this.chOpen = new System.Windows.Forms.CheckBox();
            this.cmbSeparator = new System.Windows.Forms.ComboBox();
            this.cmbEncoding = new System.Windows.Forms.ComboBox();
            this.gBWhatExpounding.SuspendLayout();
            this.grFileSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // gBWhatExpounding
            // 
            this.gBWhatExpounding.Controls.Add(this.btnSelectRange);
            this.gBWhatExpounding.Controls.Add(this.txtRange);
            this.gBWhatExpounding.Controls.Add(this.radioButton3);
            this.gBWhatExpounding.Controls.Add(this.radioButton2);
            this.gBWhatExpounding.Controls.Add(this.radioButton1);
            this.gBWhatExpounding.Location = new System.Drawing.Point(24, 45);
            this.gBWhatExpounding.Name = "gBWhatExpounding";
            this.gBWhatExpounding.Size = new System.Drawing.Size(467, 151);
            this.gBWhatExpounding.TabIndex = 1;
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
            this.btnSelectRange.Click += new System.EventHandler(this.btnSelectRange_Click);
            // 
            // txtRange
            // 
            this.txtRange.Location = new System.Drawing.Point(188, 39);
            this.txtRange.Name = "txtRange";
            this.txtRange.Size = new System.Drawing.Size(229, 20);
            this.txtRange.TabIndex = 3;
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(22, 105);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(189, 17);
            this.radioButton3.TabIndex = 2;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "Все рабочие листы в этой книге";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(22, 72);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(118, 17);
            this.radioButton2.TabIndex = 1;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "Этот рабочий лист";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(22, 39);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(138, 17);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Выбранный диапазон:";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(314, 359);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "ОК";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(405, 359);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // grFileSettings
            // 
            this.grFileSettings.Controls.Add(this.cmbEncoding);
            this.grFileSettings.Controls.Add(this.cmbSeparator);
            this.grFileSettings.Controls.Add(this.labEncoding);
            this.grFileSettings.Controls.Add(this.labSeparator);
            this.grFileSettings.Location = new System.Drawing.Point(24, 214);
            this.grFileSettings.Name = "grFileSettings";
            this.grFileSettings.Size = new System.Drawing.Size(467, 106);
            this.grFileSettings.TabIndex = 4;
            this.grFileSettings.TabStop = false;
            this.grFileSettings.Text = "Настройки файла CSV:";
            // 
            // labSeparator
            // 
            this.labSeparator.AutoSize = true;
            this.labSeparator.Location = new System.Drawing.Point(22, 37);
            this.labSeparator.Name = "labSeparator";
            this.labSeparator.Size = new System.Drawing.Size(76, 13);
            this.labSeparator.TabIndex = 0;
            this.labSeparator.Text = "Разделитель:";
            // 
            // labEncoding
            // 
            this.labEncoding.AutoSize = true;
            this.labEncoding.Location = new System.Drawing.Point(22, 65);
            this.labEncoding.Name = "labEncoding";
            this.labEncoding.Size = new System.Drawing.Size(65, 13);
            this.labEncoding.TabIndex = 1;
            this.labEncoding.Text = "Кодировка:";
            // 
            // chOpen
            // 
            this.chOpen.AutoSize = true;
            this.chOpen.Location = new System.Drawing.Point(24, 326);
            this.chOpen.Name = "chOpen";
            this.chOpen.Size = new System.Drawing.Size(182, 17);
            this.chOpen.TabIndex = 5;
            this.chOpen.Text = "Открыть файл после экспорта";
            this.chOpen.UseVisualStyleBackColor = true;
            // 
            // cmbSeparator
            // 
            this.cmbSeparator.FormattingEnabled = true;
            this.cmbSeparator.Items.AddRange(new object[] {
            "запятая",
            "точка с запятой",
            "табуляция",
            "вертикальная черта "});
            this.cmbSeparator.Location = new System.Drawing.Point(157, 37);
            this.cmbSeparator.Name = "cmbSeparator";
            this.cmbSeparator.Size = new System.Drawing.Size(260, 21);
            this.cmbSeparator.TabIndex = 2;
            this.cmbSeparator.SelectedIndexChanged += new System.EventHandler(this.cmbSeparator_SelectedIndexChanged);
            // 
            // cmbEncoding
            // 
            this.cmbEncoding.FormattingEnabled = true;
            this.cmbEncoding.Items.AddRange(new object[] {
            "Unicode(UTF-8)",
            "Кириллица(Windows)",
            "Кириллица(ISO)",
            "Кириллица(KOI8-R)",
            "Кириллица(KOI8-U)",
            "Кириллица(Mac)"});
            this.cmbEncoding.Location = new System.Drawing.Point(157, 65);
            this.cmbEncoding.Name = "cmbEncoding";
            this.cmbEncoding.Size = new System.Drawing.Size(260, 21);
            this.cmbEncoding.TabIndex = 3;
            this.cmbEncoding.SelectedIndexChanged += new System.EventHandler(this.cmbEncoding_SelectedIndexChanged);
            // 
            // ExportXlsxToCsv
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(509, 393);
            this.Controls.Add(this.chOpen);
            this.Controls.Add(this.grFileSettings);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.gBWhatExpounding);
            this.Name = "ExportXlsxToCsv";
            this.Text = "Экспорт в CSV";
            this.Load += new System.EventHandler(this.ExportXlsxToCsv_Load);
            this.gBWhatExpounding.ResumeLayout(false);
            this.gBWhatExpounding.PerformLayout();
            this.grFileSettings.ResumeLayout(false);
            this.grFileSettings.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox gBWhatExpounding;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.TextBox txtRange;
        private System.Windows.Forms.Button btnSelectRange;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.GroupBox grFileSettings;
        private System.Windows.Forms.Label labEncoding;
        private System.Windows.Forms.Label labSeparator;
        private System.Windows.Forms.ComboBox cmbEncoding;
        private System.Windows.Forms.ComboBox cmbSeparator;
        private System.Windows.Forms.CheckBox chOpen;
    }
}