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
            this.gbChoice = new System.Windows.Forms.GroupBox();
            this.btSpecifyRange = new System.Windows.Forms.Button();
            this.txtPost = new System.Windows.Forms.Label();
            this.rbSpecifyRange = new System.Windows.Forms.RadioButton();
            this.rbColumnRight = new System.Windows.Forms.RadioButton();
            this.btRange = new System.Windows.Forms.Button();
            this.txtQRcodes = new System.Windows.Forms.TextBox();
            this.txtQRcode = new System.Windows.Forms.TextBox();
            this.rbMany = new System.Windows.Forms.RadioButton();
            this.rbOne = new System.Windows.Forms.RadioButton();
            this.gbOptions = new System.Windows.Forms.GroupBox();
            this.btPathFolder = new System.Windows.Forms.Button();
            this.txtPathFolder = new System.Windows.Forms.TextBox();
            this.cbPictureFile = new System.Windows.Forms.CheckBox();
            this.tbSize = new System.Windows.Forms.TrackBar();
            this.cbBackground = new System.Windows.Forms.ComboBox();
            this.cbColour = new System.Windows.Forms.ComboBox();
            this.lbBackground = new System.Windows.Forms.Label();
            this.lbColour = new System.Windows.Forms.Label();
            this.gbPicture = new System.Windows.Forms.GroupBox();
            this.pbPicture = new System.Windows.Forms.PictureBox();
            this.btCreate = new System.Windows.Forms.Button();
            this.panelQR = new System.Windows.Forms.Panel();
            this.cbAddText = new System.Windows.Forms.CheckBox();
            this.gbСhoice.SuspendLayout();
            this.gbChoice.SuspendLayout();
            this.gbOptions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbSize)).BeginInit();
            this.gbPicture.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbPicture)).BeginInit();
            this.panelQR.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbСhoice
            // 
            this.gbСhoice.Controls.Add(this.gbChoice);
            this.gbСhoice.Controls.Add(this.btRange);
            this.gbСhoice.Controls.Add(this.txtQRcodes);
            this.gbСhoice.Controls.Add(this.txtQRcode);
            this.gbСhoice.Controls.Add(this.rbMany);
            this.gbСhoice.Controls.Add(this.rbOne);
            this.gbСhoice.Location = new System.Drawing.Point(3, 3);
            this.gbСhoice.Name = "gbСhoice";
            this.gbСhoice.Size = new System.Drawing.Size(350, 116);
            this.gbСhoice.TabIndex = 0;
            this.gbСhoice.TabStop = false;
            this.gbСhoice.Text = "Создать:";
            // 
            // gbChoice
            // 
            this.gbChoice.Controls.Add(this.btSpecifyRange);
            this.gbChoice.Controls.Add(this.txtPost);
            this.gbChoice.Controls.Add(this.rbSpecifyRange);
            this.gbChoice.Controls.Add(this.rbColumnRight);
            this.gbChoice.Location = new System.Drawing.Point(0, 79);
            this.gbChoice.Name = "gbChoice";
            this.gbChoice.Size = new System.Drawing.Size(350, 37);
            this.gbChoice.TabIndex = 11;
            this.gbChoice.TabStop = false;
            // 
            // btSpecifyRange
            // 
            this.btSpecifyRange.Location = new System.Drawing.Point(312, 9);
            this.btSpecifyRange.Name = "btSpecifyRange";
            this.btSpecifyRange.Size = new System.Drawing.Size(28, 23);
            this.btSpecifyRange.TabIndex = 10;
            this.btSpecifyRange.Text = "...";
            this.btSpecifyRange.UseVisualStyleBackColor = true;
            this.btSpecifyRange.Click += new System.EventHandler(this.btSpecifyRange_Click);
            // 
            // txtPost
            // 
            this.txtPost.AutoSize = true;
            this.txtPost.Location = new System.Drawing.Point(6, 13);
            this.txtPost.Name = "txtPost";
            this.txtPost.Size = new System.Drawing.Size(71, 13);
            this.txtPost.TabIndex = 8;
            this.txtPost.Text = "Разместить:";
            // 
            // rbSpecifyRange
            // 
            this.rbSpecifyRange.AutoSize = true;
            this.rbSpecifyRange.Location = new System.Drawing.Point(191, 12);
            this.rbSpecifyRange.Name = "rbSpecifyRange";
            this.rbSpecifyRange.Size = new System.Drawing.Size(119, 17);
            this.rbSpecifyRange.TabIndex = 9;
            this.rbSpecifyRange.TabStop = true;
            this.rbSpecifyRange.Text = "Указать диапазон";
            this.rbSpecifyRange.UseVisualStyleBackColor = true;
            // 
            // rbColumnRight
            // 
            this.rbColumnRight.AutoSize = true;
            this.rbColumnRight.Location = new System.Drawing.Point(83, 12);
            this.rbColumnRight.Name = "rbColumnRight";
            this.rbColumnRight.Size = new System.Drawing.Size(107, 17);
            this.rbColumnRight.TabIndex = 7;
            this.rbColumnRight.TabStop = true;
            this.rbColumnRight.Text = "Колонка справа";
            this.rbColumnRight.UseVisualStyleBackColor = true;
            // 
            // btRange
            // 
            this.btRange.Location = new System.Drawing.Point(312, 56);
            this.btRange.Name = "btRange";
            this.btRange.Size = new System.Drawing.Size(28, 23);
            this.btRange.TabIndex = 6;
            this.btRange.Text = "...";
            this.btRange.UseVisualStyleBackColor = true;
            this.btRange.Click += new System.EventHandler(this.btRange_Click);
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
            this.txtQRcode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtQRcode_KeyDown);
            // 
            // rbMany
            // 
            this.rbMany.AutoSize = true;
            this.rbMany.Location = new System.Drawing.Point(7, 56);
            this.rbMany.Name = "rbMany";
            this.rbMany.Size = new System.Drawing.Size(76, 17);
            this.rbMany.TabIndex = 1;
            this.rbMany.Text = "QR - коды";
            this.rbMany.UseVisualStyleBackColor = true;
            this.rbMany.CheckedChanged += new System.EventHandler(this.rbMany_CheckedChanged);
            // 
            // rbOne
            // 
            this.rbOne.AutoSize = true;
            this.rbOne.Checked = true;
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
            this.gbOptions.Controls.Add(this.cbAddText);
            this.gbOptions.Controls.Add(this.btPathFolder);
            this.gbOptions.Controls.Add(this.txtPathFolder);
            this.gbOptions.Controls.Add(this.cbPictureFile);
            this.gbOptions.Controls.Add(this.tbSize);
            this.gbOptions.Controls.Add(this.cbBackground);
            this.gbOptions.Controls.Add(this.cbColour);
            this.gbOptions.Controls.Add(this.lbBackground);
            this.gbOptions.Controls.Add(this.lbColour);
            this.gbOptions.Location = new System.Drawing.Point(3, 123);
            this.gbOptions.Name = "gbOptions";
            this.gbOptions.Size = new System.Drawing.Size(350, 166);
            this.gbOptions.TabIndex = 1;
            this.gbOptions.TabStop = false;
            this.gbOptions.Text = "Опции";
            // 
            // btPathFolder
            // 
            this.btPathFolder.Location = new System.Drawing.Point(298, 133);
            this.btPathFolder.Name = "btPathFolder";
            this.btPathFolder.Size = new System.Drawing.Size(28, 23);
            this.btPathFolder.TabIndex = 6;
            this.btPathFolder.Text = "...";
            this.btPathFolder.UseVisualStyleBackColor = true;
            this.btPathFolder.Click += new System.EventHandler(this.btPathFolder_Click);
            // 
            // txtPathFolder
            // 
            this.txtPathFolder.Location = new System.Drawing.Point(7, 136);
            this.txtPathFolder.Name = "txtPathFolder";
            this.txtPathFolder.Size = new System.Drawing.Size(285, 20);
            this.txtPathFolder.TabIndex = 5;
            // 
            // cbPictureFile
            // 
            this.cbPictureFile.AutoSize = true;
            this.cbPictureFile.Location = new System.Drawing.Point(10, 112);
            this.cbPictureFile.Name = "cbPictureFile";
            this.cbPictureFile.Size = new System.Drawing.Size(117, 17);
            this.cbPictureFile.TabIndex = 4;
            this.cbPictureFile.Text = "Сохранить в файл";
            this.cbPictureFile.UseVisualStyleBackColor = true;
            this.cbPictureFile.CheckedChanged += new System.EventHandler(this.cbPictureFile_CheckedChanged);
            // 
            // tbSize
            // 
            this.tbSize.Location = new System.Drawing.Point(10, 64);
            this.tbSize.Name = "tbSize";
            this.tbSize.Size = new System.Drawing.Size(326, 45);
            this.tbSize.TabIndex = 3;
            this.tbSize.Scroll += new System.EventHandler(this.tbSize_Scroll);
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
            this.gbPicture.Location = new System.Drawing.Point(3, 288);
            this.gbPicture.Name = "gbPicture";
            this.gbPicture.Size = new System.Drawing.Size(350, 132);
            this.gbPicture.TabIndex = 2;
            this.gbPicture.TabStop = false;
            this.gbPicture.Text = "Картинка";
            // 
            // pbPicture
            // 
            this.pbPicture.Location = new System.Drawing.Point(103, 18);
            this.pbPicture.Name = "pbPicture";
            this.pbPicture.Size = new System.Drawing.Size(133, 108);
            this.pbPicture.TabIndex = 0;
            this.pbPicture.TabStop = false;
            // 
            // btCreate
            // 
            this.btCreate.Location = new System.Drawing.Point(85, 424);
            this.btCreate.Name = "btCreate";
            this.btCreate.Size = new System.Drawing.Size(182, 23);
            this.btCreate.TabIndex = 3;
            this.btCreate.Text = "Создать";
            this.btCreate.UseVisualStyleBackColor = true;
            this.btCreate.Click += new System.EventHandler(this.btCreate_Click);
            // 
            // panelQR
            // 
            this.panelQR.AutoScroll = true;
            this.panelQR.Controls.Add(this.gbСhoice);
            this.panelQR.Controls.Add(this.btCreate);
            this.panelQR.Controls.Add(this.gbOptions);
            this.panelQR.Controls.Add(this.gbPicture);
            this.panelQR.Location = new System.Drawing.Point(3, 3);
            this.panelQR.Name = "panelQR";
            this.panelQR.Size = new System.Drawing.Size(368, 471);
            this.panelQR.TabIndex = 4;
            // 
            // cbAddText
            // 
            this.cbAddText.AutoSize = true;
            this.cbAddText.Location = new System.Drawing.Point(146, 112);
            this.cbAddText.Name = "cbAddText";
            this.cbAddText.Size = new System.Drawing.Size(180, 17);
            this.cbAddText.TabIndex = 7;
            this.cbAddText.Text = "разместить текст с QR-кодом";
            this.cbAddText.UseVisualStyleBackColor = true;
            // 
            // QRControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panelQR);
            this.Name = "QRControl";
            this.Size = new System.Drawing.Size(374, 479);
            this.Load += new System.EventHandler(this.QRControl_Load);
            this.BackColorChanged += new System.EventHandler(this.QRControl_BackColorChanged);
            this.gbСhoice.ResumeLayout(false);
            this.gbСhoice.PerformLayout();
            this.gbChoice.ResumeLayout(false);
            this.gbChoice.PerformLayout();
            this.gbOptions.ResumeLayout(false);
            this.gbOptions.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbSize)).EndInit();
            this.gbPicture.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbPicture)).EndInit();
            this.panelQR.ResumeLayout(false);
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
        private System.Windows.Forms.Panel panelQR;
        private System.Windows.Forms.RadioButton rbColumnRight;
        private System.Windows.Forms.Label txtPost;
        private System.Windows.Forms.Button btSpecifyRange;
        private System.Windows.Forms.RadioButton rbSpecifyRange;
        private System.Windows.Forms.GroupBox gbChoice;
        private System.Windows.Forms.CheckBox cbAddText;
    }
}
