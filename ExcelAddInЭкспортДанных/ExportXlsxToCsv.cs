using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddInЭкспортДанных
{
    public partial class ExportXlsxToCsv : Form
    {
        private Excel.Application excelApp;

        // Свойства для хранения выбранного диапазона, формата экспорта, разделителя, кодировки и опции открытия после экспорта
        public string ChoiceForExport { get; private set; }     // Выбор диапазона конвертации *
        public string SelectedRange { get; private set; }       // Выбранный диапазон *
        public bool ExportAsFormatted { get; private set; }     // Форма Экспорта ?
        public string CsvDelimiter { get; private set; }        // Разделитель *
        public Encoding CsvEncoding { get; private set; }       // Кодировка *
        public bool OpenAfterExport { get; private set; }       // опции открытия *
        
        public ExportXlsxToCsv()
        {
            InitializeComponent();
        }

        // Обработчик события для кнопки выбора диапазона
        private void btnSelectRange_Click(object sender, EventArgs e)
        {/*
            // Получение активного приложения Excel
            /*Excel.Application excelApp = Globals.ThisAddIn.Application;
            if (excelApp == null)
            {
                MessageBox.Show("Ошибка: не удалось получить доступ к приложению Excel.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Установка фокуса на Excel для выбора диапазона
            excelApp.ActiveWindow.Activate();
            MessageBox.Show("Выберите диапазон в Excel и нажмите Enter.", "Выбор диапазона", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // Подписка на событие KeyDown для отслеживания нажатия клавиши Enter
            excelApp.SheetSelectionChange += ExcelApp_SheetSelectionChange;
            try
            {
                // Ожидаем выбора диапазона
                excelApp.StatusBar = "Выберите диапазон ячеек и нажмите Enter";
                Excel.Range selectedRange = excelApp.InputBox("Выберите диапазон ячеек", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8) as Excel.Range;

                if (selectedRange != null)
                {
                    // Отображаем выбранный диапазон в текстовом поле
                    txtRange.Text = selectedRange.Address;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }*/
        }

        

        private void ExportXlsxToCsv_Load(object sender, EventArgs e)
        {
            btnSelectRange.Enabled = false;
            cmbEncoding.SelectedIndex = 0;
            cmbSeparator.SelectedIndex = 0;


        }

        private void cmbSeparator_SelectedIndexChanged(object sender, EventArgs e)
        {
            String selectedDelimiter = cmbSeparator.Text;
            switch (selectedDelimiter)
            {
                case "запятая":
                    CsvDelimiter = ",";
                    break;
                case "точка с запятой":
                    CsvDelimiter = ";";
                    break;

                case "табуляция)":
                    CsvDelimiter = "\t";
                    break;
                case "вертикальная черта":
                    CsvDelimiter = "|";
                    break;
            }
        }

        private void cmbEncoding_SelectedIndexChanged(object sender, EventArgs e)
        {
            String selectedEncoding = cmbEncoding.Text;
            switch (selectedEncoding)
            {
                case "Unicode(UTF-8)":
                    CsvEncoding = new UTF8Encoding(false);  // UTF-8 без BOM
                    break;

                case "Unicode(UTF-8-BOM)":
                    CsvEncoding = Encoding.UTF8;            //UTF-8-BOM
                    break;

                case "Кириллица(Windows)":
                    CsvEncoding = Encoding.GetEncoding(1251); // Windows-1251
                    break;

                case "Кириллица(ISO)":
                    CsvEncoding = Encoding.GetEncoding("ISO-8859-5");
                    break;

                case "Кириллица(KOI8-R)":
                    CsvEncoding = Encoding.GetEncoding("KOI8-R");
                    break;

                case "Кириллица(KOI8-U)":
                    CsvEncoding = Encoding.GetEncoding("KOI8-U");
                    break;

                case "Кириллица(Mac)":
                    CsvEncoding = Encoding.GetEncoding("x-mac-cyrillic");
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
        // Передаем диапазон ячеек для работы
        void funSelectedRange() 
        {
            SelectedRange = txtRange.Text;
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
    }
       /* private void ExcelApp_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            // Получение активного приложения Excel
            Excel.Application excelApp = Globals.ThisAddIn.Application;

            // Проверка нажатия клавиши Enter
            if (Control.ModifierKeys == Keys.None && Control.ModifierKeys != Keys.Shift)
            {
                SelectedRange = Target.Address;
                txtRange.Text = SelectedRange;

                // Отписка от события
                excelApp.SheetSelectionChange -= ExcelApp_SheetSelectionChange;
            }
        }
    }*/
}
