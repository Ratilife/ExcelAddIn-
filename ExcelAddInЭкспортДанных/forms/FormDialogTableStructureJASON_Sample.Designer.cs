namespace ExcelAddInЭкспортДанных.forms
{
    partial class FormDialogTableStructureJASON_Sample
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
            this.gbTable = new System.Windows.Forms.GroupBox();
            this.btDelete = new System.Windows.Forms.Button();
            this.btAdd = new System.Windows.Forms.Button();
            this.dgvTable = new System.Windows.Forms.DataGridView();
            this.labelTextKol = new System.Windows.Forms.Label();
            this.tbKolTable = new System.Windows.Forms.TextBox();
            this.rbCurrentSheet = new System.Windows.Forms.RadioButton();
            this.rbNewSheet = new System.Windows.Forms.RadioButton();
            this.gbWhere_to_place = new System.Windows.Forms.GroupBox();
            this.btOK = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.btnAddColumn = new System.Windows.Forms.Button();
            this.gbRow = new System.Windows.Forms.GroupBox();
            this.gbColumn = new System.Windows.Forms.GroupBox();
            this.btnDeleteColumn = new System.Windows.Forms.Button();
            this.gbTable.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTable)).BeginInit();
            this.gbWhere_to_place.SuspendLayout();
            this.gbRow.SuspendLayout();
            this.gbColumn.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbTable
            // 
            this.gbTable.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbTable.Controls.Add(this.gbColumn);
            this.gbTable.Controls.Add(this.gbRow);
            this.gbTable.Controls.Add(this.dgvTable);
            this.gbTable.Location = new System.Drawing.Point(12, 12);
            this.gbTable.Name = "gbTable";
            this.gbTable.Size = new System.Drawing.Size(490, 426);
            this.gbTable.TabIndex = 0;
            this.gbTable.TabStop = false;
            this.gbTable.Text = "Таблица для стркутуры";
            // 
            // btDelete
            // 
            this.btDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btDelete.Location = new System.Drawing.Point(98, 14);
            this.btDelete.Name = "btDelete";
            this.btDelete.Size = new System.Drawing.Size(75, 23);
            this.btDelete.TabIndex = 3;
            this.btDelete.Text = "Удалить";
            this.btDelete.UseVisualStyleBackColor = true;
            this.btDelete.Click += new System.EventHandler(this.btDelete_Click);
            // 
            // btAdd
            // 
            this.btAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btAdd.Location = new System.Drawing.Point(17, 15);
            this.btAdd.Name = "btAdd";
            this.btAdd.Size = new System.Drawing.Size(75, 23);
            this.btAdd.TabIndex = 1;
            this.btAdd.Text = "Добавить";
            this.btAdd.UseVisualStyleBackColor = true;
            this.btAdd.Click += new System.EventHandler(this.btAdd_Click);
            // 
            // dgvTable
            // 
            this.dgvTable.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTable.Location = new System.Drawing.Point(7, 20);
            this.dgvTable.Name = "dgvTable";
            this.dgvTable.Size = new System.Drawing.Size(477, 354);
            this.dgvTable.TabIndex = 0;
            // 
            // labelTextKol
            // 
            this.labelTextKol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelTextKol.AutoSize = true;
            this.labelTextKol.Location = new System.Drawing.Point(531, 27);
            this.labelTextKol.Name = "labelTextKol";
            this.labelTextKol.Size = new System.Drawing.Size(161, 13);
            this.labelTextKol.TabIndex = 1;
            this.labelTextKol.Text = "Количество объектов (таблиц)";
            // 
            // tbKolTable
            // 
            this.tbKolTable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.tbKolTable.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbKolTable.Location = new System.Drawing.Point(699, 22);
            this.tbKolTable.Name = "tbKolTable";
            this.tbKolTable.Size = new System.Drawing.Size(72, 20);
            this.tbKolTable.TabIndex = 2;
            // 
            // rbCurrentSheet
            // 
            this.rbCurrentSheet.AutoSize = true;
            this.rbCurrentSheet.Checked = true;
            this.rbCurrentSheet.Location = new System.Drawing.Point(6, 19);
            this.rbCurrentSheet.Name = "rbCurrentSheet";
            this.rbCurrentSheet.Size = new System.Drawing.Size(96, 17);
            this.rbCurrentSheet.TabIndex = 3;
            this.rbCurrentSheet.TabStop = true;
            this.rbCurrentSheet.Text = "Текущий лист";
            this.rbCurrentSheet.UseVisualStyleBackColor = true;
            // 
            // rbNewSheet
            // 
            this.rbNewSheet.AutoSize = true;
            this.rbNewSheet.Location = new System.Drawing.Point(137, 19);
            this.rbNewSheet.Name = "rbNewSheet";
            this.rbNewSheet.Size = new System.Drawing.Size(85, 17);
            this.rbNewSheet.TabIndex = 4;
            this.rbNewSheet.Text = "Новый лист";
            this.rbNewSheet.UseVisualStyleBackColor = true;
            // 
            // gbWhere_to_place
            // 
            this.gbWhere_to_place.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.gbWhere_to_place.Controls.Add(this.rbCurrentSheet);
            this.gbWhere_to_place.Controls.Add(this.rbNewSheet);
            this.gbWhere_to_place.Location = new System.Drawing.Point(534, 57);
            this.gbWhere_to_place.Name = "gbWhere_to_place";
            this.gbWhere_to_place.Size = new System.Drawing.Size(237, 43);
            this.gbWhere_to_place.TabIndex = 5;
            this.gbWhere_to_place.TabStop = false;
            this.gbWhere_to_place.Text = "Где расположить?";
            // 
            // btOK
            // 
            this.btOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btOK.Location = new System.Drawing.Point(606, 415);
            this.btOK.Name = "btOK";
            this.btOK.Size = new System.Drawing.Size(75, 23);
            this.btOK.TabIndex = 2;
            this.btOK.Text = "OK";
            this.btOK.UseVisualStyleBackColor = true;
            this.btOK.Click += new System.EventHandler(this.btOK_Click);
            // 
            // btCancel
            // 
            this.btCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btCancel.Location = new System.Drawing.Point(699, 415);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(75, 23);
            this.btCancel.TabIndex = 10;
            this.btCancel.Text = "Cancel";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnAddColumn
            // 
            this.btnAddColumn.Location = new System.Drawing.Point(21, 13);
            this.btnAddColumn.Name = "btnAddColumn";
            this.btnAddColumn.Size = new System.Drawing.Size(75, 23);
            this.btnAddColumn.TabIndex = 4;
            this.btnAddColumn.Text = "Добавить";
            this.btnAddColumn.UseVisualStyleBackColor = true;
            this.btnAddColumn.Click += new System.EventHandler(this.btnAddColumn_Click);
            // 
            // gbRow
            // 
            this.gbRow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.gbRow.Controls.Add(this.btAdd);
            this.gbRow.Controls.Add(this.btDelete);
            this.gbRow.Location = new System.Drawing.Point(7, 380);
            this.gbRow.Name = "gbRow";
            this.gbRow.Size = new System.Drawing.Size(195, 40);
            this.gbRow.TabIndex = 5;
            this.gbRow.TabStop = false;
            this.gbRow.Text = "Строки";
            // 
            // gbColumn
            // 
            this.gbColumn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.gbColumn.Controls.Add(this.btnDeleteColumn);
            this.gbColumn.Controls.Add(this.btnAddColumn);
            this.gbColumn.Location = new System.Drawing.Point(294, 381);
            this.gbColumn.Name = "gbColumn";
            this.gbColumn.Size = new System.Drawing.Size(190, 39);
            this.gbColumn.TabIndex = 6;
            this.gbColumn.TabStop = false;
            this.gbColumn.Text = "Колонки";
            // 
            // btnDeleteColumn
            // 
            this.btnDeleteColumn.Location = new System.Drawing.Point(102, 13);
            this.btnDeleteColumn.Name = "btnDeleteColumn";
            this.btnDeleteColumn.Size = new System.Drawing.Size(75, 23);
            this.btnDeleteColumn.TabIndex = 5;
            this.btnDeleteColumn.Text = "Удалить";
            this.btnDeleteColumn.UseVisualStyleBackColor = true;
            this.btnDeleteColumn.Click += new System.EventHandler(this.btnDeleteColumn_Click);
            // 
            // FormDialogTableStructureJASON_Sample
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btOK);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.gbWhere_to_place);
            this.Controls.Add(this.tbKolTable);
            this.Controls.Add(this.labelTextKol);
            this.Controls.Add(this.gbTable);
            this.Name = "FormDialogTableStructureJASON_Sample";
            this.gbTable.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTable)).EndInit();
            this.gbWhere_to_place.ResumeLayout(false);
            this.gbWhere_to_place.PerformLayout();
            this.gbRow.ResumeLayout(false);
            this.gbColumn.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox gbTable;
        private System.Windows.Forms.Label labelTextKol;
        private System.Windows.Forms.TextBox tbKolTable;
        private System.Windows.Forms.RadioButton rbCurrentSheet;
        private System.Windows.Forms.RadioButton rbNewSheet;
        private System.Windows.Forms.GroupBox gbWhere_to_place;
        private System.Windows.Forms.DataGridView dgvTable;
        private System.Windows.Forms.Button btDelete;
        private System.Windows.Forms.Button btOK;
        private System.Windows.Forms.Button btAdd;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.GroupBox gbRow;
        private System.Windows.Forms.Button btnAddColumn;
        private System.Windows.Forms.GroupBox gbColumn;
        private System.Windows.Forms.Button btnDeleteColumn;
    }
}