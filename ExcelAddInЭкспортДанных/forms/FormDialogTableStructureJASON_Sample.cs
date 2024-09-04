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
        public FormDialogTableStructureJASON_Sample()
        {
            InitializeComponent();
            LoadData();
        }

        private void CreateTable() 
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("Название колонки");
            dataTable.Columns.Add("Название поля JSON");
            dataTable.Columns.Add("Значение");
            
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

            // Привязываем DataTable к dgTable
            dgTable.DataSource = dataTable;
            // Настраиваем ширину столбцов
            dgTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            // Настраиваем стиль заголовков столбцов
            dgTable.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgTable.Font, System.Drawing.FontStyle.Bold);
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
            dgTable.DataSource = dataTable;
        }
    }
}
