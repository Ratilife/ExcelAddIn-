using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddInЭкспортДанных
{
    public partial class ExportXlsxToDF : Form
    {
        // Свойства для хранения выбранного диапазона, формата экспорта, разделителя, кодировки и опции открытия после экспорта
        public string ChoiceForExport { get; private set; }     // Выбор диапазона конвертации *
        public string SelectedRange { get; private set; }       // Выбранный диапазон *
        public string formatFile { get; private set; }          // выбор формата
        public bool OpenAfterExport { get; private set; }       // опции открытия *

        public string FormatDefinition { get; set; }            // определение формата для конвертации

        public ExportXlsxToDF()
        {
            InitializeComponent();
        }

        public ExportXlsxToDF(string formatDefinition)
        {
            InitializeComponent();
            FormatDefinition = formatDefinition;
        }


        private void ExportXlsxToDF_Load(object sender, EventArgs e)
        {
            btnSelectRange.Enabled = false;
            if (FormatDefinition == "xls") { 
                cmbSaveAs.SelectedIndex = 0;
                formatFile = FormatDefinition;
            }
            if (FormatDefinition == "xlsm") 
            { 
                cmbSaveAs.SelectedIndex = 1;
                formatFile = FormatDefinition;
            }
            if (FormatDefinition == "txt") 
            { 
                cmbSaveAs.SelectedIndex = 2;
                formatFile = FormatDefinition;
            }
            if (FormatDefinition == "xml") 
            { 
                cmbSaveAs.SelectedIndex = 3;
                formatFile = FormatDefinition;
            }
            if (FormatDefinition == "html")
            { 
                cmbSaveAs.SelectedIndex = 4;
                formatFile = FormatDefinition;
            }
            if (FormatDefinition == "pdf")
            { 
                cmbSaveAs.SelectedIndex = 5;
                formatFile = FormatDefinition;
            }
            if (FormatDefinition == "json")
            { 
                cmbSaveAs.SelectedIndex = 6;
                formatFile = FormatDefinition;
            }

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            //Определение способа конвертации

            if (rbRange.Checked)
            {
                ChoiceForExport = "Range";
            }
            if (rbActiveSheet.Checked)
            {
                ChoiceForExport = "ActiveSheet";
            }
            if (rdBook.Checked)
            {
                ChoiceForExport = "Book";
            }
            //Передаем открыть файл после создания 
            OpenAfterExport = chOpen.Checked;
            // Устанавливаем результат диалога как OK и закрываем форму
            DialogResult = DialogResult.OK;
            Close();
        }

        private void cmbSaveAs_SelectedIndexChanged(object sender, EventArgs e)
        {
            String selectedFormat = cmbSaveAs.Text;
            switch (selectedFormat)
            {
                case "Книга Excel 97-2003(*.xls)":
                    formatFile = "xls";  
                    break;

                case "Книга Excel с поддержкой макрасов (*.xlsm)":
                    formatFile = "xlsm";            
                    break;

                case "Текст Юникод (*.txt)":
                    formatFile = "txt"; 
                    break;

                case "XML- данные (*.xml)":
                    formatFile = "xml";
                    break;

                case "Веб-страница (*.html)":
                    formatFile = "html";
                    break;

                case "PDF(*.pdf)":
                    formatFile = "pdf";
                    break;

                case "JSON(*.json)":
                    formatFile = "json";
                    break;

                default:
                    MessageBox.Show("Неизвестная кодировка. Пожалуйста, выберите другую.");
                    break;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}






