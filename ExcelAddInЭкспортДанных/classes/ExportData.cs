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
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using System.Windows.Shapes;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelAddInЭкспортДанных
{
    internal class ExportData
    {
        // Вспомогательный метод для открытия файла
        void OpenFile(string filePath)
        {
            try
            {
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Не удается открыть файл. " + ex.Message, "Open File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        public void ExportXlsxToCsvBook(string csvBasePath, Encoding encoding, string delimiter, bool OpenAfterExport)
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
                List<string> listFilePath = new List<string>();
                // Проход по всем листам в книге
                for (int i = 1; i <= workbook.Sheets.Count; i++)
                {
                    // Получение текущего листа
                    Excel.Worksheet worksheet = workbook.Sheets[i];
                    // Получение имени текущего листа
                    string sheetName = worksheet.Name;
                    // Создание пути для сохранения текущего листа как CSV
                    //string csvPath = Path.Combine(csvBasePath, $"{sheetName}.csv");
                    string csvPath = System.IO.Path.Combine(csvBasePath, $"{sheetName}.csv");
                    listFilePath.Add(csvPath);
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
                    foreach (string path in listFilePath)
                    {
                        OpenFile(path);
                    }
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
                    OpenFile(csvPath);
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
                if (OpenAfterExport == true)
                {
                    OpenFile(csvPath);
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

        /*
        * Метод ExportActiveSheetToXML экспортирует указанный диапазон Excel в файл XML.
        *
        * Параметры:
        *   - Excel.Range usedRange: Диапазон ячеек на листе Excel, который необходимо экспортировать.
        *   - string filePath: Путь, по которому будет сохранен файл XML.
        *
        * Процесс:
        *   1. Создание нового XML документа.
        *   2. Создание корневого элемента и добавление его в документ.
        *   3. Проход по всем строкам указанного диапазона:
        *    - Для каждой строки создается элемент "row".
        *    - Проход по всем столбцам указанного диапазона:
        *      - Для каждой ячейки создается элемент "cell".
        *      - Получается значение ячейки и устанавливается в элемент "cell".
        *      - Элемент "cell" добавляется в элемент "row".
        * 4. Сохранение XML документа в указанный файл.
        *
        * Примечание:
        *    - Метод предполагает, что диапазон usedRange содержит данные, которые могут быть преобразованы в строки.
        *    - Если файл с указанным именем уже существует, он будет перезаписан.
        */
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

        private string GetCellStyle(Excel.Range cell)
        {
            StringBuilder style = new StringBuilder();

            // Получение цвета фона ячейки
            int bgColor = (int)(double)cell.Interior.Color;
            style.AppendFormat("background-color: #{0:X6};", bgColor & 0xFFFFFF);

            // Получение цвета текста ячейки
            int fontColor = (int)(double)cell.Font.Color;
            style.AppendFormat("color: #{0:X6};", fontColor & 0xFFFFFF);

            // Получение информации о шрифте
            style.AppendFormat("font-family: {0};", cell.Font.Name);
            style.AppendFormat("font-size: {0}px;", cell.Font.Size);

            // Проверка стиля шрифта
            if (cell.Font.Bold)
                style.Append("font-weight: bold;");
            if (cell.Font.Italic)
                style.Append("font-style: italic;");
            if ((int)cell.Font.Underline != (int)Excel.XlUnderlineStyle.xlUnderlineStyleNone)
                style.Append("text-decoration: underline;");

            // Получение информации о выравнивании текста
            if ((int)cell.HorizontalAlignment == (int)Excel.XlHAlign.xlHAlignCenter)
                style.Append("text-align: center;");

            return style.ToString();
        }




        public void ExportSelectedRangeToHTML(Excel.Range usedRange, string filePath)
        {
            // Создание нового StringBuilder для хранения HTML
            StringBuilder html = new StringBuilder();

            // Добавление начала HTML документа
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("<style>");
            // Добавление стилей для таблицы
            html.AppendLine("table { border-collapse: collapse; width: 100%; }");
            html.AppendLine("th, td { border: 1px solid black; padding: 5px; }");
            html.AppendLine("</style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");
            html.AppendLine("<table>");

            // Проход по всем строкам диапазона
            for (int r = 1; r <= usedRange.Rows.Count; r++)
            {
                bool rowHasData = false;
                StringBuilder rowHtml = new StringBuilder();
                rowHtml.AppendLine("<tr>");

                // Проход по всем столбцам диапазона
                for (int c = 1; c <= usedRange.Columns.Count; c++)
                {
                    // Получение значения ячейки
                    Excel.Range cell = usedRange.Cells[r, c] as Excel.Range;
                    object cellValue = cell.Value2;

                    if (cellValue != null)
                    {
                        rowHasData = true;
                        // Получение стилей ячейки
                        string cellStyle = GetCellStyle(cell);

                        // Добавление значения ячейки в HTML с сохранением стилей
                        rowHtml.AppendLine($"<td style='{cellStyle}'>{Convert.ToString(cellValue)}</td>");
                    }
                }

                rowHtml.AppendLine("</tr>");

                if (rowHasData)
                {
                    html.Append(rowHtml.ToString());
                }
            }

            // Добавление конца HTML документа
            html.AppendLine("</table>");
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            // Запись HTML в файл
            File.WriteAllText(filePath, html.ToString());
        }

        public void ExportActiveSheetToHTML(string filePath, bool OpenAfterExport)
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
                Excel.Worksheet activeSheet = workbook.ActiveSheet;
                if (activeSheet == null)
                {
                    MessageBox.Show("Error: Нет активного листа Excel.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Создание новой временной книги
                Excel.Workbook tempWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet tempSheet = tempWorkbook.Sheets[1];

                // Копирование активного листа в временную книгу
                activeSheet.Copy(tempSheet);

                // Удаление первого пустого листа из временной книги
                tempSheet.Delete();

                // Сохранение временной книги в формате HTML
                tempWorkbook.SaveAs(filePath, Excel.XlFileFormat.xlHtml);
                tempWorkbook.Close(false);

                // Вывод сообщения об успешном экспорте
                MessageBox.Show("Успешный экспорт!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Открытие файла после экспорта, если требуется
                if (OpenAfterExport)
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

        public void ExportEachSheetToHTML(string directoryPath, bool OpenAfterExport)
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
                    Excel.Worksheet sheet = workbook.Sheets[i];
                    // Создание пути для сохранения текущего листа в формате HTML
                    string filePath = System.IO.Path.Combine(directoryPath, $"{sheet.Name}.html");
                    // Создание новой временной книги
                    Excel.Workbook tempWorkbook = excelApp.Workbooks.Add();
                    Excel.Worksheet tempSheet = tempWorkbook.Sheets[1];

                    // Копирование текущего листа в временную книгу
                    sheet.Copy(tempSheet);

                    // Удаление первого пустого листа из временной книги
                    tempSheet.Delete();

                    // Сохранение временной книги в формате HTML
                    tempWorkbook.SaveAs(filePath, Excel.XlFileFormat.xlHtml);
                    tempWorkbook.Close(false);

                    // Открытие файла после экспорта, если требуется
                    if (OpenAfterExport)
                    {
                        OpenFile(filePath);
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

        

        /*
        * Метод ExportSelectedRangeToDF экспортирует выбранный диапазон ячеек из активного листа Excel в указанный формат файла.
        *
        * Параметры:
        *   - string filePath: Путь, по которому будет сохранен экспортированный файл.
        *   - string rangeAddress: Адрес диапазона ячеек, который необходимо экспортировать (например, "A1:D10").
        *   - string extension: Формат, в котором необходимо экспортировать данные (например, "pdf", "xls", "xlsm", "txt", "json", "xml", "html").
        *   - bool OpenAfterExport: Флаг, указывающий, нужно ли открывать файл после экспорта.
        *
        * Процесс:
        *   1. Получение текущего экземпляра приложения Excel и проверка его доступности.
        *   2. Получение активной книги и листа, а также проверка их доступности.
        *   3. Получение указанного диапазона ячеек на активном листе.
        *   4. Экспорт данных из выбранного диапазона в указанный формат файла:
        *    - PDF: Используется метод ExportAsFixedFormat.
        *    - XLS и XLSM: Данные копируются в новую книгу, которая затем сохраняется в указанном формате.
        *    - TXT, JSON, XML: Используются специализированные методы экспорта данных в соответствующий формат.
        *    - HTML: (предполагается, что код будет добавлен позже).
        *   5. Вывод сообщения об успешном экспорте.
        *   6. Опциональное открытие экспортированного файла после завершения экспорта.
        *
        * Примечание:
        *   - Метод предполагает, что диапазон rangeAddress содержит допустимый адрес ячеек.
        *   - Если файл с указанным именем уже существует, он будет перезаписан.
        *   - Код для экспорта в HTML формат еще не реализован и требует дополнения.
        */
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
                    ExportSelectedRangeToHTML(range, filePath);
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
                    // Исполнение перенесено в отдельный метод ExportActiveSheetToHTML
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
        * Метод ExportXlsxToDifferentFormatsBook экспортирует каждый лист активной книги Excel в указанный формат файла.
        *
        * Параметры:
        *   - string filePath: Путь, по которому будут сохранены файлы экспорта.
        *   - string extension: Расширение файла, указывающее формат экспорта. Поддерживаются: pdf, xls, xlsm, txt, xml, json, html.
        *   - bool OpenAfterExport: Если true, файл будет открыт после экспорта.
        *
        * Процесс:
        *   1. Получение текущего экземпляра приложения Excel.
        *   2. Проверка доступности активной книги и листа.
        *   3. Проход по всем листам в активной книге.
        *   4. Для каждого листа:
        *    - Формируется путь для сохранения файла с учетом имени листа и указанного расширения.
        *    - Выполняется экспорт в зависимости от указанного расширения:
        *      - PDF: Используется метод ExportAsFixedFormat.
        *      - XLS: Сохранение книги в формате Excel 97-2003.
        *      - XLSM: Сохранение книги в формате Excel с поддержкой макросов.
        *      - TXT: Экспорт в текстовый файл (вызывается метод ExportActiveSheetToTXT).
        *      - XML: Сохранение книги в формате XML.
        *      - JSON: Экспорт в JSON (вызывается метод ExportActiveSheetToJSON).
        *      - HTML: Сохранение книги в формате HTML.
        *   5. Отображение сообщения об успешном экспорте.
        *   6. Открытие файла после экспорта, если параметр OpenAfterExport установлен в true.
        *
        * Примечание:
        * - Метод предполагает, что файл, в который экспортируется, не существует или может быть перезаписан.
        * - В случае ошибок выводится сообщение с описанием ошибки.
        */
        public void ExportXlsxToDifferentFormatsBook(string filePath, string extension, bool OpenAfterExport, bool bookToOneDoc)
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

                //Создадим список для открытия n-го количества файлов в новом формате
                List<string> listFilePath = new List<string>();
                if (bookToOneDoc == true)
                {
                    // Сохранение файла HTML через диалоговое окно
                    SaveFileDialog saveFileDialog = new SaveFileDialog
                    {
                        // Установка фильтра для сохранения только в формате HTML
                        Filter = "HTML Files|*.html",
                        // Заголовок диалогового окна
                        Title = "Save as HTML File"
                    };
                    // Проверка, была ли нажата кнопка "OK" в диалоговом окне
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Получение пути для сохранения файла PDF
                        string exportPath = saveFileDialog.FileName;

                        // Сохранение активной книги в формате HTML
                        workbook.SaveAs(exportPath, Excel.XlFileFormat.xlHtml);
                    }
                }
                else
                {
                    // Проход по всем листам в книге
                    for (int i = 1; i <= workbook.Sheets.Count; i++)
                    {
                        // Получение текущего листа
                        Excel.Worksheet worksheet = workbook.Sheets[i];
                        // Получение имени текущего листа
                        string sheetName = worksheet.Name;
                        // Создание пути для сохранения текущего листа в выбранном формате
                        string exportPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePath), $"{sheetName}.{extension}");
                        listFilePath.Add(exportPath);
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

                            
                        }
                    }

                    // Вывод сообщения об успешном экспорте
                    MessageBox.Show("Успешный экспорт!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (OpenAfterExport == true)
                {
                    foreach (string path in listFilePath)
                    {
                        OpenFile(path);
                    }
                }

            }
            catch (Exception ex)
            {
                // Обработка ошибок: вывод сообщения об ошибке
                MessageBox.Show("Error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region экспортИзCsv_в_Xlsx

        public void ExportCsvToXlsx(string csvFilePath, string xlsxDirectoryPath)
        {
            // Установка лицензионного контекста
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Проверяем, что файл CSV существует
            if (!File.Exists(csvFilePath))
            {
                Console.WriteLine("Файл CSV не найден.");
                return;
            }

            // Чтение и разбор данных из CSV-файла
            var (headers, data) = ParseCsvFile(csvFilePath);

            if (headers == null || data == null || data.Length == 0)
            {
                Console.WriteLine("Ошибка при разборе CSV-файла или он пуст.");
                return;
            }

            // Извлекаем имя файла без расширения
            string csvFileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(csvFilePath);

            // Создаем путь для файла XLSX, используя директорию и имя CSV файла
            string xlsxFilePath = System.IO.Path.Combine(xlsxDirectoryPath, csvFileNameWithoutExtension + ".xlsx");

            // Проверяем и создаем директорию, если она не существует
            EnsureDirectoryExists(xlsxDirectoryPath);

            // Создаем Excel файл и таблицу
            CreateExcelFileWithTable(xlsxFilePath, headers, data);
        }

        /* Метод для чтения и разбора данных CSV файла.
         * Возвращает кортеж, содержащий заголовки и данные.
         */
        private (string[] headers, string[][] data) ParseCsvFile(string csvFilePath)
        {
            // Чтение CSV-файла
            var csvLines = File.ReadAllLines(csvFilePath);

            // Проверка на наличие данных
            if (csvLines.Length == 0)
            {
                return (null, null);
            }

            // Разделители между столбцами
            char[] delimiters = new char[] { ';', ':', '/', '*', '\\', '|', '\'' };

            // Первая строка — это шапка таблицы
            var headers = csvLines[0].Split(delimiters);

            // Остальные строки — это данные
            var data = csvLines.Skip(1).Select(line => line.Split(delimiters)).ToArray();

            return (headers, data);
        }

        /* Метод для проверки и создания директории, если она не существует.
         * Если директорию создать не удается, выбрасывается исключение.
         */
        private void EnsureDirectoryExists(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                try
                {
                    Directory.CreateDirectory(directoryPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при создании директории: {ex.Message}");
                    throw;
                }
            }
        }

        /* Метод для создания Excel файла и заполнения его данными из CSV.
         * Создает таблицу с заголовками и данными.
         */
        private void CreateExcelFileWithTable(string xlsxFilePath, string[] headers, string[][] data)
        {
            using (var package = new ExcelPackage())
            {
                // Создаем лист в Excel файле
                var worksheet = package.Workbook.Worksheets.Add("Data");

                // Определяем количество строк и столбцов
                int rowCount = data.Length + 1; // +1 для шапки таблицы
                int colCount = headers.Length;

                // Заполняем заголовки (шапку) таблицы
                for (int colIndex = 0; colIndex < colCount; colIndex++)
                {
                    worksheet.Cells[1, colIndex + 1].Value = headers[colIndex].Trim();
                }

                // Заполняем данные в лист
                for (int rowIndex = 0; rowIndex < data.Length; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < colCount; colIndex++)
                    {
                        worksheet.Cells[rowIndex + 2, colIndex + 1].Value = data[rowIndex][colIndex].Trim();
                    }
                }

                // Определяем диапазон для таблицы
                ExcelRange tableRange = worksheet.Cells[1, 1, rowCount, colCount];

                // Создаем таблицу
                var excelTable = worksheet.Tables.Add(tableRange, "ExportCsvToXlsx");
                // excelTable.TableStyle = TableStyles.Medium2; // Применяем стиль таблицы

                // Применение базового форматирования
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Сохранение файла XLSX
                try
                {
                    var fileInfo = new FileInfo(xlsxFilePath);
                    package.SaveAs(fileInfo);
                    Console.WriteLine("Экспорт завершен успешно. Файл сохранен по пути: " + xlsxFilePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка при сохранении файла XLSX: " + ex.Message);
                }
            }
        }


        #endregion

        #region импорт данных из xml
        // проверить работоспособность метода
        public void ImportXmlToExcelInActiveWorkbook(string xmlFilePath)
        {
            // Получение текущего экземпляра приложения Excel
            Excel.Application excelApp = Globals.ThisAddIn.Application;

            // Проверка, удалось ли получить доступ к приложению Excel
            if (excelApp == null)
            {
                MessageBox.Show("Error: Не удается получить доступ к приложению Excel.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Получение активной книги
                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                if (workbook == null)
                {
                    MessageBox.Show("Error: Нет активной книги Excel.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Вопрос пользователю о создании нового листа или использовании активного листа
                DialogResult result = MessageBox.Show("Создать новый лист для импорта данных?", "Импорт данных из XML", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                Excel.Worksheet worksheet;

                if (result == DialogResult.Yes)
                {
                    // Создание нового листа
                    worksheet = (Excel.Worksheet)workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                    worksheet.Name = "ImportedData";
                }
                else
                {
                    // Использование активного листа
                    worksheet = workbook.ActiveSheet;
                    if (worksheet == null)
                    {
                        MessageBox.Show("Error: Нет активного листа Excel.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // Чтение XML-файла
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlFilePath);

                // Предполагаем, что структура XML-файла известна и данные находятся в узлах "Row"
                XmlNodeList rows = xmlDoc.SelectNodes("//Row");
                int rowIndex = 1;

                foreach (XmlNode row in rows)
                {
                    int colIndex = 1;
                    foreach (XmlNode cell in row.ChildNodes)
                    {
                        worksheet.Cells[rowIndex, colIndex] = cell.InnerText;
                        colIndex++;
                    }
                    rowIndex++;
                }

                // Вывод сообщения об успешном импорте
                MessageBox.Show("Импорт данных завершен успешно!", "Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Обработка ошибок: вывод сообщения об ошибке
                MessageBox.Show("Error: " + ex.Message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    
    
    }
}
