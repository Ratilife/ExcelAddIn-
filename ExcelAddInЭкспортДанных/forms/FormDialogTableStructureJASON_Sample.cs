using ExcelAddInЭкспортДанных.classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddInЭкспортДанных.forms
{
    public partial class FormDialogTableStructureJASON_Sample : Form
    {
        private DataTable dataTable;
        private QRControl qrControlInstance;
        public FormDialogTableStructureJASON_Sample(string vidTable, QRControl qrControl)
        {
            
            InitializeComponent();
            tbKolTable.Text = "1";
            if (vidTable == "ОсновныеСредства")
            {
                LoadData();
            }
            if (vidTable == "Пользователь")
            {
                
                CreateTableForData();
            }
            qrControlInstance = qrControl;

        }

        private void CreateTable() 
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("Название колонки");
            dataTable.Columns.Add("Название поля JSON");
            dataTable.Columns.Add("Значение");
            
        }
        // Общие настройки для таблице на форме
        private void generalConstruction_dgvTable() 
        {
            // Привязываем DataTable к dgTable
            dgvTable.DataSource = dataTable;
            // Настраиваем ширину столбцов
            dgvTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            // Настраиваем стиль заголовков столбцов
            dgvTable.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvTable.Font, System.Drawing.FontStyle.Bold);
        }
        #region СозданиеПоШаблонуОсновныеСредства
        private void LoadData()
        {
            // Создаем экземпляр класса WorkingJSON
            WorkingJSON workingJSON = new WorkingJSON();
            // Получаем данные таблицы с помощью метода generateTable()
            List<Dictionary<string, string>> tableData = workingJSON.generateTable();

            // Создаем DataTable для хранения данных
            CreateTable();

            // Заполняем DataTable данными из tableData
            foreach (var row in tableData)
            {
                var dataRow = dataTable.NewRow();
                dataRow["Название колонки"] = row["Название колонки"];
                dataRow["Название поля JSON"] = row["Название поля JSON"];
                dataRow["Значение"] = row["Значение"];
                dataTable.Rows.Add(dataRow);
            }
            // Настраиваем таблицу
            generalConstruction_dgvTable();


        }
        #endregion

        #region СозданиеТаблицыДляСтруктурыПользователем
        private void CreateTableForData()
        {
            // Создаем DataTable для хранения данных
            CreateTable();
            // Настраиваем таблицу
            generalConstruction_dgvTable();
        }
        #endregion

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btAdd_Click(object sender, EventArgs e)
        {
            //TODO определится с подскасками
            // Создаем новую строку в DataTable
            var dataRow = dataTable.NewRow();
            dataRow["Название колонки"] = "Новое значение(необязательно)";
            dataRow["Название поля JSON"] = "new_value(Обязательно)";
            dataRow["Значение"] = "значение(обязательно)";

            // Добавляем строку в DataTable
            dataTable.Rows.Add(dataRow);

            // Обновляем DataGridView
            dgvTable.DataSource = dataTable;
        }

        private void btDelete_Click(object sender, EventArgs e)
        {
            if (dgvTable.SelectedRows.Count > 0)
            {
                // Получаем индекс выбранной строки
                int rowIndex = dgvTable.SelectedRows[0].Index;

                // Удаляем строку из DataTable
                dataTable.Rows[rowIndex].Delete();

                // Обновляем DataGridView
                dgvTable.DataSource = dataTable;
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите строку для удаления.");
            }
        }

        private void btOK_Click(object sender, EventArgs e)
        {
            InformationAndHelp inf = new InformationAndHelp();

            
            updateLabelOnOpenForm(inf.inf1);
            int KolTable = int.Parse(tbKolTable.Text);
            WorkingWithTables wTables = new WorkingWithTables();
            List<string> cellCoordinates = wTables.CreateTableToJSON_СellCoordinates(dataTable, KolTable, rbCurrentSheet.Checked, rbNewSheet.Checked);

            qrControlInstance.cellCoordinates = cellCoordinates.ToArray();

            //Завершаем диалог
            DialogResult = DialogResult.OK;
            Close();
        }
        private void updateLabelOnOpenForm(string text)
        {
            // Проверка, что форма qrControlInstance открыта
            if (qrControlInstance != null && !qrControlInstance.IsDisposed)
            {
                // Обновление текста Label
                qrControlInstance.updateLabelText(text);
            }
        }

        /**
        * Метод AddColumnToDataGridView добавляет новую колонку в DataGridView.
        * 
        * Метод открывает диалоговое окно для ввода имени новой колонки. Если имя не пустое и уникальное,
        * создается и добавляется новая колонка в DataGridView. В противном случае, выводится сообщение об ошибке.
        * 
        * @return void
        */

        // Метод для создания колонки
        private void AddColumnToDataGridView()
        {
            // Окно ввода для имени колонки
            string columnName = Microsoft.VisualBasic.Interaction.InputBox(
        "Введите имя новой колонки:", "Добавить колонку", "");

            if (!string.IsNullOrEmpty(columnName))
            {
                if (dgvTable.Columns[columnName] == null) // Проверка на уникальность имени колонки
                {
                    // Создаем и добавляем новую колонку
                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                    newColumn.Name = columnName;
                    newColumn.HeaderText = columnName;
                    dgvTable.Columns.Add(newColumn);
                }
                else
                {
                    MessageBox.Show("Колонка с таким именем уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Имя колонки не может быть пустым.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /**
        * Метод RemoveColumnFromDataGridView удаляет выбранные колонки из DataGridView.
        * 
        * Метод проверяет, выбраны ли колонки в DataGridView. Если выбраны, он проходит по всем выбранным колонкам
        * и удаляет их, за исключением колонок с именами "Название поля JSON" и "Значение", для которых выводится предупреждающее сообщение.
        * Если ни одна колонка не выбрана, выводится сообщение об ошибке.
        * 
        * @return void
        */
        // Метод для удаления колонки
        private void RemoveColumnFromDataGridView()
        {
            if (dgvTable.SelectedColumns.Count > 0) // Проверка, выбрана ли колонка
            {
                foreach (DataGridViewColumn column in dgvTable.SelectedColumns)
                {
                    // Запрещаем удаление столбцов с именами "Название поля JSON" и "Значение"
                    if (column.Name == "Название поля JSON" || column.Name == "Значение")
                    {
                        MessageBox.Show($"Нельзя удалить столбец \"{column.HeaderText}\".", "Запрещено", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        dgvTable.Columns.Remove(column);
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите колонку для удаления.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAddColumn_Click(object sender, EventArgs e)
        {
            AddColumnToDataGridView();
        }

        private void btnDeleteColumn_Click(object sender, EventArgs e)
        {
            //TODO: не выделяется колонка перед удалением разобраться
            RemoveColumnFromDataGridView();
        }
    }
}
