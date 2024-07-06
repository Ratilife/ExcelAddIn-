using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInЭкспортДанных
{
    internal class ExportData
    {
        // Метод для экспорта всей книги Excel в несколько CSV файлов
        void ExportXlsxToCsvBook(string xlsxPath, string csvBasePath, Encoding encoding)
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

                // Проход по всем листам в книге
                for (int i = 1; i <= workbook.Sheets.Count; i++)
                {
                    // Получение текущего листа
                    Excel.Worksheet worksheet = workbook.Sheets[i];
                    // Получение имени текущего листа
                    string sheetName = worksheet.Name;
                    // Создание пути для сохранения текущего листа как CSV
                    string csvPath = Path.Combine(csvBasePath, $"{sheetName}.csv");

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
        }

        // Метод для экспорта активного листа Excel в CSV файл
        void ExportActiveSheetToCsv(string csvPath, Encoding encoding)
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
                // Получение активного листа в книге
                Excel.Worksheet worksheet = excelApp.ActiveSheet;
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
        }

        // Метод для экспорта выбранного диапазона Excel в CSV файл
        void ExportSelectedRangeToCsv(string xlsxPath, string csvPath, string rangeAddress, Encoding encoding)
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
                // Получение выбранного диапазона
                Excel.Range selectedRange = worksheet.Range[rangeAddress];
                // Получение количества строк и столбцов в выбранном диапазоне
                int rowCount = selectedRange.Rows.Count;
                int colCount = selectedRange.Columns.Count;

                // Открытие потока для записи в CSV файл
                using (StreamWriter writer = new StreamWriter(csvPath, false, encoding))
                {
                    // Проход по всем строкам выбранного диапазона
                    for (int row = 1; row <= rowCount; row++)
                    {
                        // Объявление массива для хранения данных строки
                        string[] rowData = new string[colCount];
                        // Проход по всем столбцам строки
                        for (int col = 1; col <= colCount; col++)
                        {
                            // Запись значения ячейки в массив
                            rowData[col - 1] = selectedRange.Cells[row, col].Text.ToString();
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
        }

    }
}
