using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using Microsoft.Office.Interop.Excel;



namespace ExcelAddInЭкспортДанных
{
    public partial class MyCustomRibbonExportFeed
    {
        public string formatDefinition { get; private set; }
        private void MyCustomRibbonExportFeed_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //основной метод 
        /* void ExportXlsxToCsv(string xlsxPath, string csvPath, System.Text.Encoding encoding)
         {
             // Получение текущего экземпляра приложения Excel
             Excel.Application excelApp = Globals.ThisAddIn.Application;
             // Проверка, удалось ли получить доступ к приложению Excel
             if (excelApp == null)
             {
                 // Если не удалось, выводится сообщение об ошибке и выполнение функции прекращается
                 MessageBox.Show("Error: Не удается получить доступ к приложению Excel.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 return;
             }
             // Объявление переменной для книги Excel
             Excel.Workbook workbook = null;
             try
             {
                 // Открытие файла XLSX
                 workbook = excelApp.Workbooks.Open(xlsxPath);
                 // Получение первого листа в книге
                 Excel.Worksheet worksheet = workbook.Sheets[1];
                 // Получение количества строк и столбцов в используемом диапазоне листа
                 int rowCount = worksheet.UsedRange.Rows.Count;
                 int colCount = worksheet.UsedRange.Columns.Count;

                 // Открытие потока для записи в CSV файл
                 using (StreamWriter writer = new StreamWriter(csvPath, false, encoding))
                 {
                     // Проход по всем строкам используемого диапазона
                     for (int row = 1; row <= rowCount; row++)
                     {
                         // Объявление массива для хранения данных строки
                         string[] rowData = new string[colCount];
                         // Проход по всем столбцам строки
                         for (int col = 1; col <= colCount; col++)
                         {
                             // Запись значения ячейки в массив
                             rowData[col - 1] = worksheet.Cells[row, col].Text.ToString();
                         }
                         // Запись строки в CSV файл с разделителем ';'
                         writer.WriteLine(string.Join(";", rowData));
                     }
                 }
                 // Вывод сообщения об успешном экспорте
                 MessageBox.Show("Успешный экспорт!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
             }
             catch (Exception ex)
             {
                 // Обработка ошибок: вывод сообщения об ошибке
                 MessageBox.Show("Error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
             finally
             {
                 // Закрытие книги Excel, если она была открыта
                 if (workbook != null)
                 {
                     workbook.Close(false);
                 }
             }
         }*/

        void exportFormatSelection(string formatExport, string filter, string title)
        {
            formatDefinition = formatExport;
            using (ExportXlsxToDF form = new ExportXlsxToDF(formatDefinition))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    //Определяем способ экспорта
                    String ChoiceForExport = form.ChoiceForExport;
                    if (ChoiceForExport != null)
                    {
                        ExportData exportData = new ExportData();
                        if (ChoiceForExport == "Book")
                        {
                            FolderBrowserDialog fbd = new FolderBrowserDialog();
                            fbd.ShowNewFolderButton = false;
                            if (fbd.ShowDialog() == DialogResult.OK)
                            {
                                string filePath = fbd.SelectedPath + "\\";
                                if (form.bookToOneDoc == true)
                                {
                                    exportData.ExportXlsxToDifferentFormatsBook(filePath, form.formatFile, form.OpenAfterExport, form.bookToOneDoc);
                                }
                                else
                                {
                                    exportData.ExportEachSheetToHTML(filePath, form.OpenAfterExport);
                                }
                            }
                        }
                        else
                        {
                            // Сохранение файла PDF через диалоговое окно
                            SaveFileDialog saveFileDialog = new SaveFileDialog
                            {
                                // Установка фильтра для сохранения только в формате PDF
                                Filter = filter,
                                // Заголовок диалогового окна
                                Title = title
                            };
                            // Проверка, была ли нажата кнопка "OK" в диалоговом окне
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                // Получение пути для сохранения файла PDF
                                string filePath = saveFileDialog.FileName;

                                if (ChoiceForExport == "Range")
                                {
                                    string SelectedRange = form.SelectedRange;
                                    exportData.ExportSelectedRangeToDF(filePath, SelectedRange, form.formatFile, form.OpenAfterExport);
                                }
                                if (ChoiceForExport == "ActiveSheet")
                                {
                                    if (form.formatFile == "html")
                                    {
                                        exportData.ExportActiveSheetToHTML(filePath, form.OpenAfterExport);
                                    }
                                    else
                                    {
                                        exportData.ExportActiveSheetToDifferentFormats(filePath, form.formatFile, form.OpenAfterExport);
                                    }


                                }

                            }

                        }


                    }
                }
            }
        }

        private void butExportXLSXtoCSV_Click(object sender, RibbonControlEventArgs e)
        {
            //* 1 Открытие окна "ЭкспортВ_CSV"
            //  2 Передать данные с формы в метод ExportXlsxToCsv
            //*/ 
            using (ExportXlsxToCsv form = new ExportXlsxToCsv())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    //Определяем способ экспорта
                    String ChoiceForExport = form.ChoiceForExport;
                    if (ChoiceForExport != null)
                    {
                        ExportData exportData = new ExportData();
                        if (ChoiceForExport == "Book")
                        {
                            // Сохранение файла CSV через диалоговое окно
                            FolderBrowserDialog fbd = new FolderBrowserDialog();
                            fbd.ShowNewFolderButton = false;
                            if (fbd.ShowDialog() == DialogResult.OK)
                            {
                                string csvPath = fbd.SelectedPath;
                                exportData.ExportXlsxToCsvBook(csvPath, form.CsvEncoding, form.CsvDelimiter, form.OpenAfterExport);
                            }
                        }
                        else
                        {

                            // Сохранение файла CSV через диалоговое окно
                            SaveFileDialog saveFileDialog = new SaveFileDialog
                            {
                                // Установка фильтра для сохранения только в формате CSV
                                Filter = "CSV Files|*.csv",
                                // Заголовок диалогового окна
                                Title = "Save as CSV File"
                            };
                            // Проверка, была ли нажата кнопка "OK" в диалоговом окне
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                // Получение пути для сохранения файла CSV
                                string csvPath = saveFileDialog.FileName;

                                if (ChoiceForExport == "Range")
                                {
                                    //string SelectedRange = $"\"{form.SelectedRange}\"";
                                    string SelectedRange = form.SelectedRange;
                                    exportData.ExportSelectedRangeToCsv(csvPath, SelectedRange, form.CsvEncoding, form.CsvDelimiter, form.OpenAfterExport);
                                }
                                if (ChoiceForExport == "ActiveSheet")
                                {
                                    exportData.ExportActiveSheetToCsv(csvPath, form.CsvEncoding, form.CsvDelimiter, form.OpenAfterExport);
                                }

                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("Пожалуйста, выберите диапазон экспорта");
                    }

                }
            }
        }

        private void butExportXLSXtoPDF_Click(object sender, RibbonControlEventArgs e)
        {
            string formatExport = "pdf";
            string filter = "PDF Files|*.pdf";
            string title = "Save as PDF File";
            exportFormatSelection(formatExport, filter, title);

        }

        private void butExportXLSXtoTXT_Click(object sender, RibbonControlEventArgs e)
        {
            string formatExport = "txt";
            string filter = "TXT Files|*.txt";
            string title = "Save as TXT File";
            exportFormatSelection(formatExport, filter, title);
        }

        private void butExportXLSXtoJSON_Click(object sender, RibbonControlEventArgs e)
        {
            string formatExport = "json";
            string filter = "JSON Files|*.json";
            string title = "Save as JSON File";
            exportFormatSelection(formatExport, filter, title);
        }

        private void butExportXLSXtoXLS_Click(object sender, RibbonControlEventArgs e)
        {
            string formatExport = "xls";
            string filter = "XLS Files|*.xls";
            string title = "Save as XLS File";
            exportFormatSelection(formatExport, filter, title);
        }

        private void butExportXLSXtoXLSM_Click(object sender, RibbonControlEventArgs e)
        {
            string formatExport = "xlsm";
            string filter = "XLSM Files|*.xlsm";
            string title = "Save as XLSM File";
            exportFormatSelection(formatExport, filter, title);
        }

        private void butExportXLSXtoXML_Click(object sender, RibbonControlEventArgs e)
        {
            string formatExport = "xml";
            string filter = "XML Files|*.xml";
            string title = "Save as XML File";
            exportFormatSelection(formatExport, filter, title);
        }

        private void butExportXLSXtoHTML_Click(object sender, RibbonControlEventArgs e)
        {
            string formatExport = "html";
            string filter = "HTML Files|*.html";
            string title = "Save as HTML File";
            exportFormatSelection(formatExport, filter, title);
        }

        private void btCreateTable_Click(object sender, RibbonControlEventArgs e)
        {
            WorkingWithTables tables = new WorkingWithTables();

        }
    }
}
