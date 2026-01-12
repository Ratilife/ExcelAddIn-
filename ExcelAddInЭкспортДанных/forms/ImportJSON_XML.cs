using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExcelAddInЭкспортДанных.classes;

namespace ExcelAddInЭкспортДанных.forms
{
    public partial class Import_JSON_XML : Form
    {
        public string FormatDefinition { get; set; }            // определение формата для конвертации
        public string formatFile { get; private set; }          // выбор формата
        private string fileFilter; // Сохраняем фильтр для диалога выбора файла

        public Import_JSON_XML(string formatDefinition, string filter)
        {
            InitializeComponent();
            //this.btCreate.Click += btCreate_Click;
            FormatDefinition = formatDefinition;
            fileFilter = filter;
        }

        

        private void butFilePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            // Используем переданный фильтр, если он есть, иначе определяем по выбранному формату
            if (!string.IsNullOrWhiteSpace(fileFilter))
            {
                dialog.Filter = fileFilter + "|All files (*.*)|*.*";
            }
            else
            {
                string selectedFormat = cmbSaveAs.SelectedItem?.ToString() ?? "";

                if (selectedFormat.Contains("JSON"))
                {
                    dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
                }
                else if (selectedFormat.Contains("XML"))
                {
                    dialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
                }
                else
                {
                    dialog.Filter = "JSON files (*.json)|*.json|XML files (*.xml)|*.xml|All files (*.*)|*.*";
                }
            }

            // Устанавливаем заголовок диалога в зависимости от формата
            if (FormatDefinition == "json")
            {
                dialog.Title = "Выберите JSON файл";
            }
            else if (FormatDefinition == "xml")
            {
                dialog.Title = "Выберите XML файл";
            }
            else
            {
                dialog.Title = "Выберите файл";
            }

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                tbFilePath.Text = dialog.FileName;
            }
        }

        private void btCreate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbFilePath.Text))
            {
                MessageBox.Show("Пожалуйста, выберите файл для импорта.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(tbFilePath.Text))
            {
                MessageBox.Show("Выбранный файл не существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string selectedFormat = cmbSaveAs.SelectedItem?.ToString() ?? "";
            bool createNewSheet = rbNewSheet.Checked;

            if (selectedFormat.Contains("JSON"))
            {
                ImportData importData = new ImportData();
                importData.ImportJsonToExcelInActiveWorkbook(tbFilePath.Text, createNewSheet);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else if (selectedFormat.Contains("XML"))
            {
                MessageBox.Show("Импорт XML будет реализован позже.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите формат файла для импорта.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Import_JSON_XML_Load(object sender, EventArgs e)
        {
            // Устанавливаем выбранный формат в комбобоксе на основе переданного formatDefinition
            if (FormatDefinition == "json")
            {
                // Ищем элемент, содержащий "JSON"
                for (int i = 0; i < cmbSaveAs.Items.Count; i++)
                {
                    if (cmbSaveAs.Items[i].ToString().Contains("JSON"))
                    {
                        cmbSaveAs.SelectedIndex = i;
                        formatFile = FormatDefinition;
                        break;
                    }
                }
            }
            else if (FormatDefinition == "xml")
            {
                // Ищем элемент, содержащий "XML"
                for (int i = 0; i < cmbSaveAs.Items.Count; i++)
                {
                    if (cmbSaveAs.Items[i].ToString().Contains("XML"))
                    {
                        cmbSaveAs.SelectedIndex = i;
                        formatFile = FormatDefinition;
                        break;
                    }
                }
            }
            else
            {
                // Если формат не определен, выбираем первый элемент
                if (cmbSaveAs.Items.Count > 0)
                {
                    cmbSaveAs.SelectedIndex = 0;
                }
            }
        }
    }
}
