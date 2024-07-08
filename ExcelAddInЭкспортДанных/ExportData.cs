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
        /*
         * Метод ExportXlsxToCsvBook экспортирует все листы активной книги Excel в отдельные CSV файлы.
         * 
         * Параметры:
         *  - string csvBasePath: Базовый путь для сохранения CSV файлов.
         *  - Encoding encoding: Кодировка для записи CSV файлов.
         *  - string delimiter: Символ, используемый в качестве разделителя значений в CSV файлах.
         * 
         * Процесс:
         *  1. Получение текущего экземпляра приложения Excel.
         *  2. Проверка наличия активной книги.
         *  3. Проход по всем листам активной книги.
         *  4. Для каждого листа:
         *     a. Получение имени листа.
         *     b. Создание пути для сохранения листа как CSV.
         *     c. Получение используемого диапазона на листе.
         *     d. Запись данных из ячеек листа в CSV файл с использованием указанного разделителя и кодировки.
         *  5. Обработка возможных ошибок и вывод сообщений пользователю.
         * 
         * Примечание:
         * Метод предполагает, что книга Excel уже открыта и активна.
         */
        public void ExportXlsxToCsvBook( string csvBasePath, Encoding encoding, string delimiter)
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
            
            try
            {
                // Получение активной книги
                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                if (workbook == null)
                {
                    MessageBox.Show("Error: Нет активной книги Excel.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

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
                            // Запись строки в CSV файл с разделителем 
                            writer.WriteLine(string.Join(delimiter, rowData));
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
            
        }

        // Метод для экспорта активного листа Excel в CSV файл
        /*
         * Метод ExportActiveSheetToCsv экспортирует все ячейки активного листа активной книги Excel в файл CSV.
         * 
         * Параметры:
         *  - string csvPath: Путь к CSV файлу, в который будут записаны данные.
         *  - Encoding encoding: Кодировка для записи CSV файла.
         *  - string delimiter: Символ, используемый в качестве разделителя значений в CSV файле.
         * 
         * Процесс:
         *  1. Получение текущего экземпляра приложения Excel.
         *  2. Проверка наличия активной книги и листа.
         *  3. Получение всех ячеек на активном листе.
         *  4. Запись данных из ячеек в CSV файл с использованием указанного разделителя и кодировки.
         *  5. Обработка возможных ошибок и вывод сообщений пользователю.
         * 
         * Примечание:
         * Метод предполагает, что книга Excel уже открыта и активна, и работает с активным листом этой книги.
         */
        public void ExportActiveSheetToCsv(string csvPath, Encoding encoding, string delimiter)
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

            try
            {
                // Получение активной книги
                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                if (workbook == null)
                {
                    MessageBox.Show("Error: Нет активной книги Excel.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Получение активного листа
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                if (worksheet == null)
                {
                    MessageBox.Show("Error: Нет активного листа Excel.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Получение используемого диапазона на активном листе
                Excel.Range usedRange = worksheet.UsedRange;
                // Получение количества строк и столбцов в используемом диапазоне
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

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
                            rowData[col - 1] = usedRange.Cells[row, col].Text.ToString();
                        }
                        // Запись строки в CSV файл с использованием указанного разделителя
                        writer.WriteLine(string.Join(delimiter, rowData));
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
        }


        // Метод для экспорта выбранного диапазона Excel в CSV файл
        /* Метод ExportSelectedRangeToCsv экспортирует выбранный диапазон ячеек активного листа активной книги Excel в файл CSV. 
         * Параметры:
         *  string csvPath:
         *          Это строка, представляющая путь к файлу CSV, который будет создан. 
         *          Этот путь указывает на местоположение и имя нового CSV файла, 
         *          в который будут записаны данные из выбранного диапазона Excel.
         *  string rangeAddress:
         *          Это строка, представляющая адрес диапазона в формате Excel (например, "A1"), 
         *          который нужно экспортировать. Диапазон указывает на область листа Excel, 
         *          данные из которой будут извлечены и записаны в CSV файл.
         *  Encoding encoding:
         *          Это объект типа Encoding, который определяет кодировку для записи CSV файла. 
         *          Примеры кодировок включают Encoding.UTF8, Encoding.ASCII, и т.д. 
         *          Кодировка определяет, как символы будут преобразованы в байты при записи в файл.
         *  string delimiter: нужно добавить параметр
         *          Это строка, представляет собой разделитель - символ, используемый для разделения значений (ячеек) в строке данных 
         * Процесс:
         *      1. Получение текущего экземпляра приложения Excel.
         *      2. Проверка наличия активной книги и листа.
         *      3. Получение указанного диапазона на активном листе.
         *      4. Запись данных из диапазона в CSV файл с использованием указанного разделителя и кодировки.
         *      5. Обработка возможных ошибок и вывод сообщений пользователю.
         * 
         * Примечание:
         *          Метод предполагает, что книга Excel уже открыта и активна, и работает с активным листом этой книги.
         */
        public void ExportSelectedRangeToCsv(string csvPath, string rangeAddress, Encoding encoding, string delimiter)
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

            try
            {
                // Получение активной книги
                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                if (workbook == null)
                {
                    MessageBox.Show("Error: Нет активной книги Excel.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Получение активного листа
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                if (worksheet == null)
                {
                    MessageBox.Show("Error: Нет активного листа Excel.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

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
                        // Запись строки в CSV файл с использованием указанного разделителя
                        writer.WriteLine(string.Join(delimiter, rowData));
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
        }

    }
}
