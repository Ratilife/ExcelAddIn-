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
            this.lbNameTable = new System.Windows.Forms.Label();
            this.lbCellAddress = new System.Windows.Forms.Label();
            this.tbNameTable = new System.Windows.Forms.TextBox();
            this.tbCellAddress = new System.Windows.Forms.TextBox();
            this.lbCountColumns = new System.Windows.Forms.Label();
            this.lbRowCount = new System.Windows.Forms.Label();
            this.tbСolumns = new System.Windows.Forms.TextBox();
            this.tbRows = new System.Windows.Forms.TextBox();
            this.gbСreate = new System.Windows.Forms.GroupBox();
            this.rbNewSheet = new System.Windows.Forms.RadioButton();
            this.rbActivSeheet = new System.Windows.Forms.RadioButton();
            this.btCreate = new System.Windows.Forms.Button();
            this.gbСreate.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbNameTable
            // 
            this.lbNameTable.AutoSize = true;
            this.lbNameTable.Location = new System.Drawing.Point(33, 37);
            this.lbNameTable.Name = "lbNameTable";
            this.lbNameTable.Size = new System.Drawing.Size(75, 13);
            this.lbNameTable.TabIndex = 1;
            this.lbNameTable.Text = "Имя таблицы";
            // 
            // lbCellAddress
            // 
            this.lbCellAddress.AutoSize = true;
            this.lbCellAddress.Location = new System.Drawing.Point(33, 72);
            this.lbCellAddress.Name = "lbCellAddress";
            this.lbCellAddress.Size = new System.Drawing.Size(274, 13);
            this.lbCellAddress.TabIndex = 2;
            this.lbCellAddress.Text = "Адрес ячейки с которой начать построение таблицы";
            // 
            // tbNameTable
            // 
            this.tbNameTable.Location = new System.Drawing.Point(115, 29);
            this.tbNameTable.Name = "tbNameTable";
            this.tbNameTable.Size = new System.Drawing.Size(271, 20);
            this.tbNameTable.TabIndex = 3;
            // 
            // tbCellAddress
            // 
            this.tbCellAddress.Location = new System.Drawing.Point(334, 69);
            this.tbCellAddress.Name = "tbCellAddress";
            this.tbCellAddress.Size = new System.Drawing.Size(52, 20);
            this.tbCellAddress.TabIndex = 4;
            // 
            // lbCountColumns
            // 
            this.lbCountColumns.AutoSize = true;
            this.lbCountColumns.Location = new System.Drawing.Point(33, 104);
            this.lbCountColumns.Name = "lbCountColumns";
            this.lbCountColumns.Size = new System.Drawing.Size(116, 13);
            this.lbCountColumns.TabIndex = 5;
            this.lbCountColumns.Text = "Количество столбцов";
            // 
            // lbRowCount
            // 
            this.lbRowCount.AutoSize = true;
            this.lbRowCount.Location = new System.Drawing.Point(33, 131);
            this.lbRowCount.Name = "lbRowCount";
            this.lbRowCount.Size = new System.Drawing.Size(98, 13);
            this.lbRowCount.TabIndex = 6;
            this.lbRowCount.Text = "Количество строк";
            // 
            // tbСolumns
            // 
            this.tbСolumns.Location = new System.Drawing.Point(334, 104);
            this.tbСolumns.Name = "tbСolumns";
            this.tbСolumns.Size = new System.Drawing.Size(51, 20);
            this.tbСolumns.TabIndex = 7;
            // 
            // tbRows
            // 
            this.tbRows.Location = new System.Drawing.Point(334, 131);
            this.tbRows.Name = "tbRows";
            this.tbRows.Size = new System.Drawing.Size(51, 20);
            this.tbRows.TabIndex = 7;
            // 
            // gbСreate
            // 
            this.gbСreate.Controls.Add(this.rbNewSheet);
            this.gbСreate.Controls.Add(this.rbActivSeheet);
            this.gbСreate.Location = new System.Drawing.Point(36, 168);
            this.gbСreate.Name = "gbСreate";
            this.gbСreate.Size = new System.Drawing.Size(349, 49);
            this.gbСreate.TabIndex = 8;
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
            // btCreate
            // 
            this.btCreate.Location = new System.Drawing.Point(36, 243);
            this.btCreate.Name = "btCreate";
            this.btCreate.Size = new System.Drawing.Size(231, 23);
            this.btCreate.TabIndex = 9;
            this.btCreate.Text = "Создать";
            this.btCreate.UseVisualStyleBackColor = true;
            this.btCreate.Click += new System.EventHandler(this.btCreate_Click);
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
            this.Controls.Add(this.tbCellAddress);
            this.Controls.Add(this.tbNameTable);
            this.Controls.Add(this.lbCellAddress);
            this.Controls.Add(this.lbNameTable);
            this.Name = "TableСreationForm";
            this.Size = new System.Drawing.Size(407, 328);
            this.Load += new System.EventHandler(this.TableСreationForm_Load);
            this.gbСreate.ResumeLayout(false);
            this.gbСreate.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lbNameTable;
        private System.Windows.Forms.Label lbCellAddress;
        private System.Windows.Forms.TextBox tbNameTable;
        private System.Windows.Forms.TextBox tbCellAddress;
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
