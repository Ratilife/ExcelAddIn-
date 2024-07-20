namespace ExcelAddInЭкспортДанных
{
    partial class QRControl
    {
        /// <summary> 
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary> 
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.gbСhoice = new System.Windows.Forms.GroupBox();
            this.txtQRcodes = new System.Windows.Forms.TextBox();
            this.txtQRcode = new System.Windows.Forms.TextBox();
            this.rbMany = new System.Windows.Forms.RadioButton();
            this.rbOne = new System.Windows.Forms.RadioButton();
            this.gbOptions = new System.Windows.Forms.GroupBox();
            this.tbSize = new System.Windows.Forms.TrackBar();
            this.cbBackground = new System.Windows.Forms.ComboBox();
            this.cbColour = new System.Windows.Forms.ComboBox();
            this.lbBackground = new System.Windows.Forms.Label();
            this.lbColour = new System.Windows.Forms.Label();
            this.gbPicture = new System.Windows.Forms.GroupBox();
            this.pbPicture = new System.Windows.Forms.PictureBox();
            this.btCreate = new System.Windows.Forms.Button();
            this.cbPictureFile = new System.Windows.Forms.CheckBox();
            this.txtPathFolder = new System.Windows.Forms.TextBox();
            this.btPathFolder = new System.Windows.Forms.Button();
            this.btRange = new System.Windows.Forms.Button();
            this.gbСhoice.SuspendLayout();
            this.gbOptions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbSize)).BeginInit();
            this.gbPicture.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbPicture)).BeginInit();
            this.SuspendLayout();
            // 
            // gbСhoice
            // 
            this.gbСhoice.Controls.Add(this.btRange);
            this.gbСhoice.Controls.Add(this.txtQRcodes);
            this.gbСhoice.Controls.Add(this.txtQRcode);
            this.gbСhoice.Controls.Add(this.rbMany);
            this.gbСhoice.Controls.Add(this.rbOne);
            this.gbСhoice.Location = new System.Drawing.Point(20, 26);
            this.gbСhoice.Name = "gbСhoice";
            this.gbСhoice.Size = new System.Drawing.Size(348, 99);
            this.gbСhoice.TabIndex = 0;
            this.gbСhoice.TabStop = false;
            this.gbСhoice.Text = "Создать:";
            // 
            // txtQRcodes
            // 
            this.txtQRcodes.Location = new System.Drawing.Point(103, 56);
            this.txtQRcodes.Name = "txtQRcodes";
            this.txtQRcodes.Size = new System.Drawing.Size(203, 20);
            this.txtQRcodes.TabIndex = 3;
            // 
            // txtQRcode
            // 
            this.txtQRcode.Location = new System.Drawing.Point(103, 20);
            this.txtQRcode.Name = "txtQRcode";
            this.txtQRcode.Size = new System.Drawing.Size(239, 20);
            this.txtQRcode.TabIndex = 2;
            // 
            // rbMany
            // 
            this.rbMany.AutoSize = true;
            this.rbMany.Location = new System.Drawing.Point(7, 56);
            this.rbMany.Name = "rbMany";
            this.rbMany.Size = new System.Drawing.Size(76, 17);
            this.rbMany.TabIndex = 1;
            this.rbMany.TabStop = true;
            this.rbMany.Text = "QR - коды";
            this.rbMany.UseVisualStyleBackColor = true;
            // 
            // rbOne
            // 
            this.rbOne.AutoSize = true;
            this.rbOne.Location = new System.Drawing.Point(7, 20);
            this.rbOne.Name = "rbOne";
            this.rbOne.Size = new System.Drawing.Size(68, 17);
            this.rbOne.TabIndex = 0;
            this.rbOne.TabStop = true;
            this.rbOne.Text = "QR - код";
            this.rbOne.UseVisualStyleBackColor = true;
            // 
            // gbOptions
            // 
            this.gbOptions.Controls.Add(this.btPathFolder);
            this.gbOptions.Controls.Add(this.txtPathFolder);
            this.gbOptions.Controls.Add(this.cbPictureFile);
            this.gbOptions.Controls.Add(this.tbSize);
            this.gbOptions.Controls.Add(this.cbBackground);
            this.gbOptions.Controls.Add(this.cbColour);
            this.gbOptions.Controls.Add(this.lbBackground);
            this.gbOptions.Controls.Add(this.lbColour);
            this.gbOptions.Location = new System.Drawing.Point(20, 144);
            this.gbOptions.Name = "gbOptions";
            this.gbOptions.Size = new System.Drawing.Size(342, 157);
            this.gbOptions.TabIndex = 1;
            this.gbOptions.TabStop = false;
            this.gbOptions.Text = "Опции";
            // 
            // tbSize
            // 
            this.tbSize.Location = new System.Drawing.Point(10, 64);
            this.tbSize.Name = "tbSize";
            this.tbSize.Size = new System.Drawing.Size(326, 45);
            this.tbSize.TabIndex = 3;
            // 
            // cbBackground
            // 
            this.cbBackground.FormattingEnabled = true;
            this.cbBackground.Location = new System.Drawing.Point(276, 29);
            this.cbBackground.Name = "cbBackground";
            this.cbBackground.Size = new System.Drawing.Size(60, 21);
            this.cbBackground.TabIndex = 2;
            // 
            // cbColour
            // 
            this.cbColour.FormattingEnabled = true;
            this.cbColour.Location = new System.Drawing.Point(87, 30);
            this.cbColour.Name = "cbColour";
            this.cbColour.Size = new System.Drawing.Size(60, 21);
            this.cbColour.TabIndex = 2;
            // 
            // lbBackground
            // 
            this.lbBackground.AutoSize = true;
            this.lbBackground.Location = new System.Drawing.Point(203, 29);
            this.lbBackground.Name = "lbBackground";
            this.lbBackground.Size = new System.Drawing.Size(61, 13);
            this.lbBackground.TabIndex = 1;
            this.lbBackground.Text = "Цвет фона";
            // 
            // lbColour
            // 
            this.lbColour.AutoSize = true;
            this.lbColour.Location = new System.Drawing.Point(7, 30);
            this.lbColour.Name = "lbColour";
            this.lbColour.Size = new System.Drawing.Size(59, 13);
            this.lbColour.TabIndex = 0;
            this.lbColour.Text = "Цвет кода";
            // 
            // gbPicture
            // 
            this.gbPicture.Controls.Add(this.pbPicture);
            this.gbPicture.Location = new System.Drawing.Point(20, 307);
            this.gbPicture.Name = "gbPicture";
            this.gbPicture.Size = new System.Drawing.Size(336, 192);
            this.gbPicture.TabIndex = 2;
            this.gbPicture.TabStop = false;
            this.gbPicture.Text = "Картинка";
            // 
            // pbPicture
            // 
            this.pbPicture.Location = new System.Drawing.Point(63, 29);
            this.pbPicture.Name = "pbPicture";
            this.pbPicture.Size = new System.Drawing.Size(182, 146);
            this.pbPicture.TabIndex = 0;
            this.pbPicture.TabStop = false;
            // 
            // btCreate
            // 
            this.btCreate.Location = new System.Drawing.Point(83, 505);
            this.btCreate.Name = "btCreate";
            this.btCreate.Size = new System.Drawing.Size(182, 23);
            this.btCreate.TabIndex = 3;
            this.btCreate.Text = "Создать";
            this.btCreate.UseVisualStyleBackColor = true;
            // 
            // cbPictureFile
            // 
            this.cbPictureFile.AutoSize = true;
            this.cbPictureFile.Location = new System.Drawing.Point(10, 104);
            this.cbPictureFile.Name = "cbPictureFile";
            this.cbPictureFile.Size = new System.Drawing.Size(117, 17);
            this.cbPictureFile.TabIndex = 4;
            this.cbPictureFile.Text = "Сохранить в файл";
            this.cbPictureFile.UseVisualStyleBackColor = true;
            // 
            // txtPathFolder
            // 
            this.txtPathFolder.Location = new System.Drawing.Point(7, 128);
            this.txtPathFolder.Name = "txtPathFolder";
            this.txtPathFolder.Size = new System.Drawing.Size(285, 20);
            this.txtPathFolder.TabIndex = 5;
            // 
            // btPathFolder
            // 
            this.btPathFolder.Location = new System.Drawing.Point(298, 125);
            this.btPathFolder.Name = "btPathFolder";
            this.btPathFolder.Size = new System.Drawing.Size(28, 23);
            this.btPathFolder.TabIndex = 6;
            this.btPathFolder.Text = "...";
            this.btPathFolder.UseVisualStyleBackColor = true;
            // 
            // btRange
            // 
            this.btRange.Location = new System.Drawing.Point(312, 56);
            this.btRange.Name = "btRange";
            this.btRange.Size = new System.Drawing.Size(28, 23);
            this.btRange.TabIndex = 6;
            this.btRange.Text = "...";
            this.btRange.UseVisualStyleBackColor = true;
            // 
            // QRControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btCreate);
            this.Controls.Add(this.gbPicture);
            this.Controls.Add(this.gbOptions);
            this.Controls.Add(this.gbСhoice);
            this.Name = "QRControl";
            this.Size = new System.Drawing.Size(381, 540);
            this.gbСhoice.ResumeLayout(false);
            this.gbСhoice.PerformLayout();
            this.gbOptions.ResumeLayout(false);
            this.gbOptions.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbSize)).EndInit();
            this.gbPicture.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbPicture)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbСhoice;
        private System.Windows.Forms.RadioButton rbOne;
        private System.Windows.Forms.TextBox txtQRcode;
        private System.Windows.Forms.RadioButton rbMany;
        private System.Windows.Forms.TextBox txtQRcodes;
        private System.Windows.Forms.GroupBox gbOptions;
        private System.Windows.Forms.ComboBox cbBackground;
        private System.Windows.Forms.ComboBox cbColour;
        private System.Windows.Forms.Label lbBackground;
        private System.Windows.Forms.Label lbColour;
        private System.Windows.Forms.TrackBar tbSize;
        private System.Windows.Forms.GroupBox gbPicture;
        private System.Windows.Forms.PictureBox pbPicture;
        private System.Windows.Forms.Button btCreate;
        private System.Windows.Forms.Button btPathFolder;
        private System.Windows.Forms.TextBox txtPathFolder;
        private System.Windows.Forms.CheckBox cbPictureFile;
        private System.Windows.Forms.Button btRange;
    }
}
