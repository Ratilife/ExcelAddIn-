namespace ExcelAddInЭкспортДанных.forms
{
    partial class ExportCsvToXlsx
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
            this.labCsvFilePath = new System.Windows.Forms.Label();
            this.labXlsxFilePath = new System.Windows.Forms.Label();
            this.tbCsvFilePath = new System.Windows.Forms.TextBox();
            this.tbXlsxFilePath = new System.Windows.Forms.TextBox();
            this.butXlsxFilePath = new System.Windows.Forms.Button();
            this.butCsvFilePath = new System.Windows.Forms.Button();
            this.butOK = new System.Windows.Forms.Button();
            this.cbActiveWorkbook = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // labCsvFilePath
            // 
            this.labCsvFilePath.AutoSize = true;
            this.labCsvFilePath.Location = new System.Drawing.Point(22, 38);
            this.labCsvFilePath.Name = "labCsvFilePath";
            this.labCsvFilePath.Size = new System.Drawing.Size(174, 13);
            this.labCsvFilePath.TabIndex = 0;
            this.labCsvFilePath.Text = "путь к файлу с расширением csv";
            // 
            // labXlsxFilePath
            // 
            this.labXlsxFilePath.AutoSize = true;
            this.labXlsxFilePath.Location = new System.Drawing.Point(23, 80);
            this.labXlsxFilePath.Name = "labXlsxFilePath";
            this.labXlsxFilePath.Size = new System.Drawing.Size(174, 13);
            this.labXlsxFilePath.TabIndex = 1;
            this.labXlsxFilePath.Text = "путь к файлу с расширением xlsx";
            // 
            // tbCsvFilePath
            // 
            this.tbCsvFilePath.Location = new System.Drawing.Point(218, 38);
            this.tbCsvFilePath.Name = "tbCsvFilePath";
            this.tbCsvFilePath.Size = new System.Drawing.Size(319, 20);
            this.tbCsvFilePath.TabIndex = 2;
            // 
            // tbXlsxFilePath
            // 
            this.tbXlsxFilePath.Location = new System.Drawing.Point(218, 73);
            this.tbXlsxFilePath.Name = "tbXlsxFilePath";
            this.tbXlsxFilePath.Size = new System.Drawing.Size(319, 20);
            this.tbXlsxFilePath.TabIndex = 2;
            // 
            // butXlsxFilePath
            // 
            this.butXlsxFilePath.Location = new System.Drawing.Point(544, 73);
            this.butXlsxFilePath.Name = "butXlsxFilePath";
            this.butXlsxFilePath.Size = new System.Drawing.Size(37, 20);
            this.butXlsxFilePath.TabIndex = 3;
            this.butXlsxFilePath.Text = "...";
            this.butXlsxFilePath.UseVisualStyleBackColor = true;
            this.butXlsxFilePath.Click += new System.EventHandler(this.butXlsxFilePath_Click);
            // 
            // butCsvFilePath
            // 
            this.butCsvFilePath.Location = new System.Drawing.Point(544, 38);
            this.butCsvFilePath.Name = "butCsvFilePath";
            this.butCsvFilePath.Size = new System.Drawing.Size(37, 20);
            this.butCsvFilePath.TabIndex = 3;
            this.butCsvFilePath.Text = "...";
            this.butCsvFilePath.UseVisualStyleBackColor = true;
            this.butCsvFilePath.Click += new System.EventHandler(this.butCsvFilePath_Click);
            // 
            // butOK
            // 
            this.butOK.Location = new System.Drawing.Point(506, 119);
            this.butOK.Name = "butOK";
            this.butOK.Size = new System.Drawing.Size(75, 23);
            this.butOK.TabIndex = 4;
            this.butOK.Text = "OK";
            this.butOK.UseVisualStyleBackColor = true;
            this.butOK.Click += new System.EventHandler(this.butOK_Click);
            // 
            // cbActiveWorkbook
            // 
            this.cbActiveWorkbook.AutoSize = true;
            this.cbActiveWorkbook.Location = new System.Drawing.Point(26, 119);
            this.cbActiveWorkbook.Name = "cbActiveWorkbook";
            this.cbActiveWorkbook.Size = new System.Drawing.Size(211, 17);
            this.cbActiveWorkbook.TabIndex = 5;
            this.cbActiveWorkbook.Text = "перенести данные в активную книгу";
            this.cbActiveWorkbook.UseVisualStyleBackColor = true;
            this.cbActiveWorkbook.CheckedChanged += new System.EventHandler(this.cbActiveWorkbook_CheckedChanged);
            // 
            // ExportCsvToXlsx
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(593, 154);
            this.Controls.Add(this.cbActiveWorkbook);
            this.Controls.Add(this.butOK);
            this.Controls.Add(this.butCsvFilePath);
            this.Controls.Add(this.butXlsxFilePath);
            this.Controls.Add(this.tbXlsxFilePath);
            this.Controls.Add(this.tbCsvFilePath);
            this.Controls.Add(this.labXlsxFilePath);
            this.Controls.Add(this.labCsvFilePath);
            this.Name = "ExportCsvToXlsx";
            this.Text = "ExportCsvToXlsx";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labCsvFilePath;
        private System.Windows.Forms.Label labXlsxFilePath;
        private System.Windows.Forms.TextBox tbCsvFilePath;
        private System.Windows.Forms.TextBox tbXlsxFilePath;
        private System.Windows.Forms.Button butXlsxFilePath;
        private System.Windows.Forms.Button butCsvFilePath;
        private System.Windows.Forms.Button butOK;
        private System.Windows.Forms.CheckBox cbActiveWorkbook;
    }
}