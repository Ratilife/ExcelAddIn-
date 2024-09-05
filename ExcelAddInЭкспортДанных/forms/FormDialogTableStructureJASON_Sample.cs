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

            qrControlInstance.isJSON = true;
            updateLabelOnOpenForm(inf.inf4);
            int KolTable = int.Parse(tbKolTable.Text);
            WorkingWithTables wTables = new WorkingWithTables();
            wTables.CreateTableToJSON(dataTable, KolTable, rbCurrentSheet.Checked, rbNewSheet.Checked);
            //Завершаем диалог
            DialogResult = DialogResult.OK;
            Close();
        }

        static void updateLabelOnOpenForm(string text)
        {
            // Поиск открытой формы типа QRControl
            //форма передана в конструктор как этим фактом фоспользоваться
            QRControl openForm = Application.OpenForms.OfType<QRControl>().FirstOrDefault();
            if (openForm != null)
            {
                // Обновление текста Label
                openForm.updateLabelText(text);
            }
        }
    }
}
