using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Formatting = Newtonsoft.Json.Formatting;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInЭкспортДанных
{
    internal class ExportData
    {

        void OpenFile(string filePath) 
        {
            System.Diagnostics.Process.Start(filePath);
        }

        #region экспорт данных в CSV

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
        public void ExportXlsxToCsvBook( string csvBasePath, Encoding encoding, string delimiter, bool OpenAfterExport)
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
                if (OpenAfterExport == true)
                {

                }
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
        public void ExportActiveSheetToCsv(string csvPath, Encoding encoding, string delimiter, bool OpenAfterExport)
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
                if (OpenAfterExport == true)
                {

                }
            }
            catch (Exception ex)
            {
                // Обработка ошибок: вывод сообщения об ошибке
                MessageBox.Show("Error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Метод для экспорта выбранного диапазона Excel в CSV файл
        /* 
         * Метод ExportSelectedRangeToCsv экспортирует выбранный диапазон ячеек активного листа активной книги Excel в файл CSV. 
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
        public void ExportSelectedRangeToCsv(string csvPath, string rangeAddress, Encoding encoding, string delimiter, bool OpenAfterExport)
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
                // отккрыть файл
                if ( OpenAfterExport == true) 
                {
                    OpenFile( csvPath);
                }

            }
            catch (Exception ex)
            {
                // Обработка ошибок: вывод сообщения об ошибке
                MessageBox.Show("Error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region экспорт данных в различные форматы
        /*
            Этот метод экспортирует данные из активного листа Excel в формате JSON. 

            Параметры:
                - usedRange: используемый диапазон на активном листе Excel.
                - filePath: путь к файлу, в который будут записаны данные.

            Процесс:
                1. Метод сначала получает количество строк и столбцов в используемом диапазоне.
                2. Затем он создает список словарей для хранения данных всех строк.
                3. Метод проходит по всем строкам используемого диапазона. Для каждой строки создается словарь, в который записываются данные всех столбцов этой строки.
                4. Каждый словарь добавляется в список.
                5. Наконец, список преобразуется в формат JSON и записывается в файл по указанному пути.

            Примечание: 
                - Значения ячеек записываются в словарь в виде строк.
                - Если файл с указанным путем уже существует, его содержимое будет перезаписано.
        */
        void ExportActiveSheetToJSON(Excel.Range usedRange, string filePath) 
        {
             // Получение количества строк и столбцов в используемом диапазоне
             int rowCount = usedRange.Rows.Count;
             int colCount = usedRange.Columns.Count;

             // Создание списка для хранения данных всех строк
             List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();

             // Проход по всем строкам используемого диапазона
             for (int row = 1; row <= rowCount; row++)
             {
                 // Создание словаря для хранения данных строки
                 Dictionary<string, object> rowData = new Dictionary<string, object>();
                 // Проход по всем столбцам строки
                 for (int col = 1; col <= colCount; col++)
                 {
                     // Запись значения ячейки в словарь
                     rowData["Column" + col] = usedRange.Cells[row, col].Text.ToString();
                 }
                 // Добавление словаря в список
                 rows.Add(rowData);
             }

             // Преобразование списка в JSON и запись в файл
             string json = JsonConvert.SerializeObject(rows, Formatting.Indented);
             File.WriteAllText(filePath, json);
        }

        /*
        Этот метод экспортирует данные из активного листа Excel в формате TXT.

        Параметры:
            - usedRange: используемый диапазон на активном листе Excel.
            - filePath: путь к файлу, в который будут записаны данные.

        Процесс:
            1. Метод сначала получает количество строк и столбцов в используемом диапазоне.
            2. Затем он открывает поток для записи в TXT файл.
            3. Метод проходит по всем строкам используемого диапазона. Для каждой строки создается массив, в который записываются данные всех столбцов этой строки.
            4. Каждая строка записывается в TXT файл. Данные в строке разделяются символом табуляции ("\t").
            5. После обработки всех строк поток записи закрывается.

        Примечание: 
            - Значения ячеек записываются в файл в виде строк.
            - Если файл с указанным путем уже существует, его содержимое будет перезаписано.
        */
        void ExportActiveSheetToTXT(Excel.Range usedRange, string filePath)
        {
            // Получение количества строк и столбцов в используемом диапазоне
            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            // Открытие потока для записи в TXT файл
            using (StreamWriter writer = new StreamWriter(filePath, false))
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
                    // Запись строки в TXT файл
                    writer.WriteLine(string.Join("\t", rowData));
                }
            }
        }

        void ExportActiveSheetToXML(Excel.Range usedRange, string filePath)
        {
            // Создание нового XML документа
            XmlDocument xmlDoc = new XmlDocument();

            // Создание корневого элемента
            XmlElement root = xmlDoc.CreateElement("root");
            xmlDoc.AppendChild(root);

            // Проход по всем строкам диапазона
            for (int r = 1; r <= usedRange.Rows.Count; r++)
            {
                // Создание элемента для строки
                XmlElement row = xmlDoc.CreateElement("row");
                root.AppendChild(row);

                // Проход по всем столбцам диапазона
                for (int c = 1; c <= usedRange.Columns.Count; c++)
                {
                    // Создание элемента для ячейки
                    XmlElement cell = xmlDoc.CreateElement("cell");

                    // Получение значения ячейки
                    object cellValue = (usedRange.Cells[r, c] as Excel.Range).Value2;

                    // Установка значения элемента
                    cell.InnerText = Convert.ToString(cellValue);

                    // Добавление элемента ячейки к элементу строки
                    row.AppendChild(cell);
                }
            }

            // Сохранение XML документа в файл
            xmlDoc.Save(filePath);
        }

        void ExportActiveSheetToHTML(Excel.Range usedRange, string filePath)
        {
            // Создание нового StringBuilder для хранения HTML
            StringBuilder html = new StringBuilder();

            // Добавление начала HTML документа
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<body>");
            html.AppendLine("<table>");

            // Проход по всем строкам диапазона
            for (int r = 1; r <= usedRange.Rows.Count; r++)
            {
                html.AppendLine("<tr>");

                // Проход по всем столбцам диапазона
                for (int c = 1; c <= usedRange.Columns.Count; c++)
                {
                    // Получение значения ячейки
                    object cellValue = (usedRange.Cells[r, c] as Excel.Range).Value2;

                    // Добавление значения ячейки в HTML
                    html.AppendLine("<td>" + Convert.ToString(cellValue) + "</td>");
                }

                html.AppendLine("</tr>");
            }

            // Добавление конца HTML документа
            html.AppendLine("</table>");
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            // Запись HTML в файл
            File.WriteAllText(filePath, html.ToString());

        }

        public void ExportSelectedRangeToDF(string filePath, string rangeAddress, string extension, bool OpenAfterExport)
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
                Excel.Range range = worksheet.get_Range(rangeAddress);
                if (range == null)
                {
                    MessageBox.Show("Error: Нет выбранного диапазона.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Экспорт выбранного диапазона в выбранном формате
                if (extension.ToLower() == "pdf")
                {
                    range.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, filePath);
                }
                else if (extension.ToLower() == "xls")
                {
                    // Сохранение выбранного диапазона в формате XLS
                    range.Copy();
                    Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
                    Excel.Worksheet newWorksheet = newWorkbook.ActiveSheet;
                    newWorksheet.Paste();
                    newWorkbook.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);
                    newWorkbook.Close();
                }
                else if (extension.ToLower() == "xlsm")
                {
                    // Сохранение выбранного диапазона в формате XLSM
                    range.Copy();
                    Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
                    Excel.Worksheet newWorksheet = newWorkbook.ActiveSheet;
                    newWorksheet.Paste();
                    newWorkbook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                    newWorkbook.Close();
                }
                else if (extension.ToLower() == "txt")
                {
                    ExportActiveSheetToTXT(range, filePath);
                }
                else if (extension.ToLower() == "json")
                {
                    ExportActiveSheetToJSON(range, filePath);
                }
                else if (extension.ToLower() == "xml")
                {
                    ExportActiveSheetToXML(range, filePath);
                }
                else if (extension.ToLower() == "html")
                {
                
                }

                    // Вывод сообщения об успешном экспорте
                    MessageBox.Show("Успешный экспорт!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);



                if (OpenAfterExport == true)
                {
                    OpenFile(filePath);
                }
            }
            catch (Exception ex)
            {
                // Обработка ошибок: вывод сообщения об ошибке
                MessageBox.Show("Error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        /*
        Этот метод экспортирует данные из активного листа Excel в различные форматы файлов.

        Параметры:
            - filePath: путь к файлу, в который будут записаны данные.
            - extension: расширение файла, определяющее формат файла для экспорта.

        Процесс:
            1. Метод сначала получает текущий экземпляр приложения Excel и проверяет, удалось ли получить доступ к нему.
            2. Затем он получает активную книгу и активный лист в этой книге.
            3. В зависимости от указанного расширения файла, метод экспортирует данные активного листа в соответствующий формат файла.
            4. Если процесс экспорта проходит успешно, выводится сообщение об успешном экспорте.
            5. Если в процессе экспорта возникает ошибка, выводится сообщение об ошибке.

        Примечание: 
            - Для экспорта в форматы TXT и JSON используются отдельные методы `ExportActiveSheetToTXT` и `ExportActiveSheetToJSON`.
            - Если файл с указанным путем уже существует, его содержимое будет перезаписано.
            - Если расширение файла не соответствует ни одному из поддерживаемых форматов, метод не будет делать ничего.
        */
        // Метод для экспорта активного листа Excel в разные форматы
        public void ExportActiveSheetToDifferentFormats(string filePath, string extension, bool OpenAfterExport)
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

                // Экспорт активного листа в выбранном формате
                if (extension.ToLower() == "pdf")
                {
                    worksheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, filePath);
                }
                else if (extension.ToLower() == "xls")
                {
                    // Сохранение активной книги в формате XLS
                    workbook.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);
                }
                else if (extension.ToLower() == "xlsm")
                {
                    // Сохранение активной книги в формате XLSM
                    workbook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                }
                // вывести в отдельный метод ExportActiveSheetToTXT(worksheet WS)
                else if (extension.ToLower() == "txt")
                {
                    // Получение используемого диапазона на активном листе
                    Excel.Range usedRange = worksheet.UsedRange;
                    ExportActiveSheetToTXT(usedRange, filePath);
                }
                else if (extension.ToLower() == "xml")
                {
                    // Сохранение активной книги в формате XML
                    workbook.SaveAs(filePath, Excel.XlFileFormat.xlXMLSpreadsheet);
                }
                
                else if (extension.ToLower() == "json")
                {
                    // Получение используемого диапазона на активном листе
                    Excel.Range usedRange = worksheet.UsedRange;
                    ExportActiveSheetToJSON(usedRange, filePath);
                }
                else if (extension.ToLower() == "html")
                {
                    // Сохранение активной книги в формате HTML
                    workbook.SaveAs(filePath, Excel.XlFileFormat.xlHtml);
                }

                // Вывод сообщения об успешном экспорте
                MessageBox.Show("Успешный экспорт!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (OpenAfterExport == true)
                {
                    OpenFile(filePath);
                }

            }
            catch (Exception ex)
            {
                // Обработка ошибок: вывод сообщения об ошибке
                MessageBox.Show("Error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ExportXlsxToDifferentFormatsBook(string filePath, string extension, bool OpenAfterExport)
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
                    // Создание пути для сохранения текущего листа в выбранном формате
                    string exportPath = Path.Combine(Path.GetDirectoryName(filePath), $"{sheetName}.{extension}");

                    // Экспорт текущего листа в выбранном формате
                    if (extension.ToLower() == "pdf")
                    {
                        worksheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, exportPath);
                    }
                    else if (extension.ToLower() == "xls")
                    {
                        // Сохранение активной книги в формате XLS
                        workbook.SaveAs(exportPath, Excel.XlFileFormat.xlExcel8);
                    }
                    else if (extension.ToLower() == "xlsm")
                    {
                        // Сохранение активной книги в формате XLSM
                        workbook.SaveAs(exportPath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                    }
                    else if (extension.ToLower() == "txt")
                    {
                        // Получение используемого диапазона на текущем листе
                        Excel.Range usedRange = worksheet.UsedRange;
                        ExportActiveSheetToTXT(usedRange, exportPath);
                    }
                    else if (extension.ToLower() == "xml")
                    {
                        // Сохранение активной книги в формате XML
                        workbook.SaveAs(exportPath, Excel.XlFileFormat.xlXMLSpreadsheet);
                    }
                    else if (extension.ToLower() == "json")
                    {
                        // Получение используемого диапазона на текущем листе
                        Excel.Range usedRange = worksheet.UsedRange;
                        ExportActiveSheetToJSON(usedRange, exportPath);
                    }
                    else if (extension.ToLower() == "html")
                    {
                        // Сохранение активной книги в формате HTML
                        workbook.SaveAs(exportPath, Excel.XlFileFormat.xlHtml);
                    }
                }

                // Вывод сообщения об успешном экспорте
                MessageBox.Show("Успешный экспорт!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (OpenAfterExport == true)
                {
                    OpenFile(filePath);
                }

            }
            catch (Exception ex)
            {
                // Обработка ошибок: вывод сообщения об ошибке
                MessageBox.Show("Error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}
