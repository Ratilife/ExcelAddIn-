namespace ExcelAddInЭкспортДанных
{
    partial class TableСreationForm
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
            this.lbHeading = new System.Windows.Forms.Label();
            this.lbNameTable = new System.Windows.Forms.Label();
            this.lbCellAddress = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.lbCountColumns = new System.Windows.Forms.Label();
            this.lbRowCount = new System.Windows.Forms.Label();
            this.tbСolumns = new System.Windows.Forms.TextBox();
            this.tbRows = new System.Windows.Forms.TextBox();
            this.gbСreate = new System.Windows.Forms.GroupBox();
            this.rbActivSeheet = new System.Windows.Forms.RadioButton();
            this.rbNewSheet = new System.Windows.Forms.RadioButton();
            this.btCreate = new System.Windows.Forms.Button();
            this.gbСreate.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbHeading
            // 
            this.lbHeading.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lbHeading.AutoSize = true;
            this.lbHeading.Font = new System.Drawing.Font("Microsoft Tai Le", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lbHeading.Location = new System.Drawing.Point(64, 14);
            this.lbHeading.Name = "lbHeading";
            this.lbHeading.Size = new System.Drawing.Size(220, 27);
            this.lbHeading.TabIndex = 0;
            this.lbHeading.Text = "Создание таблицы";
            this.lbHeading.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // lbNameTable
            // 
            this.lbNameTable.AutoSize = true;
            this.lbNameTable.Location = new System.Drawing.Point(33, 73);
            this.lbNameTable.Name = "lbNameTable";
            this.lbNameTable.Size = new System.Drawing.Size(75, 13);
            this.lbNameTable.TabIndex = 1;
            this.lbNameTable.Text = "Имя таблицы";
            // 
            // lbCellAddress
            // 
            this.lbCellAddress.AutoSize = true;
            this.lbCellAddress.Location = new System.Drawing.Point(33, 108);
            this.lbCellAddress.Name = "lbCellAddress";
            this.lbCellAddress.Size = new System.Drawing.Size(274, 13);
            this.lbCellAddress.TabIndex = 2;
            this.lbCellAddress.Text = "Адрес ячейки с которой начать построение таблицы";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(115, 65);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(271, 20);
            this.textBox1.TabIndex = 3;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(334, 105);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(52, 20);
            this.textBox2.TabIndex = 4;
            // 
            // lbCountColumns
            // 
            this.lbCountColumns.AutoSize = true;
            this.lbCountColumns.Location = new System.Drawing.Point(33, 140);
            this.lbCountColumns.Name = "lbCountColumns";
            this.lbCountColumns.Size = new System.Drawing.Size(116, 13);
            this.lbCountColumns.TabIndex = 5;
            this.lbCountColumns.Text = "Количество столбцов";
            // 
            // lbRowCount
            // 
            this.lbRowCount.AutoSize = true;
            this.lbRowCount.Location = new System.Drawing.Point(33, 167);
            this.lbRowCount.Name = "lbRowCount";
            this.lbRowCount.Size = new System.Drawing.Size(98, 13);
            this.lbRowCount.TabIndex = 6;
            this.lbRowCount.Text = "Количество строк";
            // 
            // tbСolumns
            // 
            this.tbСolumns.Location = new System.Drawing.Point(334, 140);
            this.tbСolumns.Name = "tbСolumns";
            this.tbСolumns.Size = new System.Drawing.Size(51, 20);
            this.tbСolumns.TabIndex = 7;
            // 
            // tbRows
            // 
            this.tbRows.Location = new System.Drawing.Point(334, 167);
            this.tbRows.Name = "tbRows";
            this.tbRows.Size = new System.Drawing.Size(51, 20);
            this.tbRows.TabIndex = 7;
            // 
            // gbСreate
            // 
            this.gbСreate.Controls.Add(this.rbNewSheet);
            this.gbСreate.Controls.Add(this.rbActivSeheet);
            this.gbСreate.Location = new System.Drawing.Point(36, 204);
            this.gbСreate.Name = "gbСreate";
            this.gbСreate.Size = new System.Drawing.Size(349, 49);
            this.gbСreate.TabIndex = 8;
            this.gbСreate.TabStop = false;
            this.gbСreate.Text = "Создать:";
            // 
            // rbActivSeheet
            // 
            this.rbActivSeheet.AutoSize = true;
            this.rbActivSeheet.Location = new System.Drawing.Point(7, 26);
            this.rbActivSeheet.Name = "rbActivSeheet";
            this.rbActivSeheet.Size = new System.Drawing.Size(123, 17);
            this.rbActivSeheet.TabIndex = 0;
            this.rbActivSeheet.TabStop = true;
            this.rbActivSeheet.Text = "На активном листе";
            this.rbActivSeheet.UseVisualStyleBackColor = true;
            // 
            // rbNewSheet
            // 
            this.rbNewSheet.AutoSize = true;
            this.rbNewSheet.Location = new System.Drawing.Point(173, 26);
            this.rbNewSheet.Name = "rbNewSheet";
            this.rbNewSheet.Size = new System.Drawing.Size(106, 17);
            this.rbNewSheet.TabIndex = 1;
            this.rbNewSheet.TabStop = true;
            this.rbNewSheet.Text = "На новом листе";
            this.rbNewSheet.UseVisualStyleBackColor = true;
            // 
            // btCreate
            // 
            this.btCreate.Location = new System.Drawing.Point(36, 279);
            this.btCreate.Name = "btCreate";
            this.btCreate.Size = new System.Drawing.Size(75, 23);
            this.btCreate.TabIndex = 9;
            this.btCreate.Text = "Создать";
            this.btCreate.UseVisualStyleBackColor = true;
            // 
            // TableСreationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btCreate);
            this.Controls.Add(this.gbСreate);
            this.Controls.Add(this.tbRows);
            this.Controls.Add(this.tbСolumns);
            this.Controls.Add(this.lbRowCount);
            this.Controls.Add(this.lbCountColumns);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.lbCellAddress);
            this.Controls.Add(this.lbNameTable);
            this.Controls.Add(this.lbHeading);
            this.Name = "TableСreationForm";
            this.Size = new System.Drawing.Size(407, 328);
            this.gbСreate.ResumeLayout(false);
            this.gbСreate.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbHeading;
        private System.Windows.Forms.Label lbNameTable;
        private System.Windows.Forms.Label lbCellAddress;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label lbCountColumns;
        private System.Windows.Forms.Label lbRowCount;
        private System.Windows.Forms.TextBox tbСolumns;
        private System.Windows.Forms.TextBox tbRows;
        private System.Windows.Forms.GroupBox gbСreate;
        private System.Windows.Forms.RadioButton rbNewSheet;
        private System.Windows.Forms.RadioButton rbActivSeheet;
        private System.Windows.Forms.Button btCreate;
    }
}
