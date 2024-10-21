using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace ExcelAddInЭкспортДанных
{
    internal class WorkingWithTables
    {
    /*
    * Создает таблицу в Excel на основе заданных параметров.
    *
    * Параметры:
    * - startCellAddress: Адрес начальной ячейки для таблицы.
    * - columnCount: Количество столбцов в таблице.
    * - rowCount: Количество строк в таблице.
    * - onActiveSheet: Флаг, указывающий, использовать ли активный лист.
    * - onNewSheet: Флаг, указывающий, создать ли новый лист.
    * - tableName: Имя для создаваемой таблицы.
    *
    * Метод выполняет следующие действия:
    * - Создает новый лист или использует активный лист в зависимости от параметров.
    * - Определяет диапазон ячеек для таблицы.
    * - Проверяет наличие существующих таблиц в заданном диапазоне и предотвращает создание новой таблицы, если в диапазоне уже есть таблица.
    * - Создает новую таблицу в указанном диапазоне и задает ей имя.
    */

        public void CreateTable(string startCellAddress, int columnCount, int rowCount, bool onActiveSheet, bool onNewSheet, string tableName)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;

            if (onNewSheet)
            {
                // Создание нового листа
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Worksheets.Add();
            }
            else if (onActiveSheet)
            {
                // Использование активного листа
                worksheet = excelApp.ActiveSheet;
            }
            else
            {
                // Если не указано иное, используется первый лист
                worksheet = excelApp.Worksheets[1];
            }

            Microsoft.Office.Interop.Excel.Range startCell = worksheet.Range[startCellAddress];
            Microsoft.Office.Interop.Excel.Range endCell = startCell.get_Offset(rowCount, columnCount - 1);
            Microsoft.Office.Interop.Excel.Range tableRange = worksheet.Range[startCell, endCell];

            // Проверка наличия таблиц в заданном диапазоне
            foreach (Microsoft.Office.Interop.Excel.ListObject existingTable in worksheet.ListObjects)
            {
                if (excelApp.Intersect(existingTable.Range, tableRange) != null)
                {
                    // Если в заданном диапазоне уже есть таблица, вы можете выбрать другой диапазон
                    // или удалить существующую таблицу. Здесь мы просто прерываем выполнение метода.
                    Console.WriteLine("В заданном диапазоне уже есть таблица. Выберите другой диапазон.");
                    return;
                }
            }
            // Создание таблицы
            Microsoft.Office.Interop.Excel.ListObject table = worksheet.ListObjects.AddEx(
            SourceType: Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange,
            Source: tableRange,
            XlListObjectHasHeaders: Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes,
            TableStyleName: "TableStyleMedium2");
            // Установка имени таблицы
            table.Name = tableName;

        }
    /*
    * Создает таблицу в Excel на основе данных из DataTable и сохраняет её в формате JSON.
    *
    * Параметры:
    * - dataTable: Таблица данных для вставки в Excel.
    * - kolTable: Количество таблиц для создания.
    * - onActiveSheet: Флаг, указывающий, использовать ли активный лист.
    * - onNewSheet: Флаг, указывающий, создать ли новый лист.
    *
    * Метод выполняет следующие действия:
    * - Использует активный лист, создает новый лист или использует первый лист в зависимости от параметров.
    * - Вставляет заголовки столбцов и данные строк из DataTable в указанный диапазон ячеек.
    * - Применяет форматирование к ячейкам, включая жирный шрифт для заголовков и границы для всех ячеек.
    * - Автоматически изменяет ширину столбцов по содержимому.
    */

        public void CreateTableToJSON(DataTable dataTable, int kolTable, bool onActiveSheet, bool onNewSheet)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;

            if (onActiveSheet)
            {
                // Использование активного листа
                worksheet = excelApp.ActiveSheet;
            }
            else if (onNewSheet)
            {
                // Создание нового листа
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Worksheets.Add();
            }
            else
            {
                // Если не указано иное, используется первый лист
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Worksheets[1];
            }

            for (int i = 0; i < kolTable; i++)
            {
                int startRow = i * (dataTable.Rows.Count + 2) + 1;

                // Заголовки столбцов
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[startRow, col + 1].Value = dataTable.Columns[col].ColumnName;
                    worksheet.Cells[startRow, col + 1].Font.Bold = true;
                    worksheet.Cells[startRow, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[startRow, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[startRow, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[startRow, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                }

                // Данные строк
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[startRow + row + 1, col + 1].Value = dataTable.Rows[row][col];
                        worksheet.Cells[startRow + row + 1, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        worksheet.Cells[startRow + row + 1, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        worksheet.Cells[startRow + row + 1, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        worksheet.Cells[startRow + row + 1, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    }
                }

                // Автоматическое изменение ширины столбцов по содержимому
                worksheet.Columns.AutoFit();
            }
        }

    /*
    * Создает таблицу в Excel на основе данных из DataTable и возвращает список координат ячеек для привязки QR-кодов.
    *
    * Параметры:
    * - dataTable: Таблица данных для вставки в Excel.
    * - kolTable: Количество таблиц для создания.
    * - onActiveSheet: Флаг, указывающий, использовать ли активный лист.
    * - onNewSheet: Флаг, указывающий, создать ли новый лист.
    *
    * Возвращает:
    * - List<string>: Список координат ячеек, расположенных на одну колонку правее последнего столбца таблицы на уровне заголовка.
    *
    * Метод выполняет следующие действия:
    * - Использует активный лист, создает новый лист или использует первый лист в зависимости от параметров.
    * - Вставляет заголовки столбцов и данные строк из DataTable в указанный диапазон ячеек.
    * - Применяет форматирование к ячейкам, включая жирный шрифт для заголовков и границы для всех ячеек.
    * - Определяет координаты ячеек, расположенных на одну колонку правее последнего столбца таблицы на уровне заголовка, и добавляет их в список.
    * - Автоматически изменяет ширину столбцов по содержимому.
    */
        public List<string> CreateTableToJSON_СellCoordinates(DataTable dataTable, int kolTable, bool onActiveSheet, bool onNewSheet)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            // координаты ячеек куда будем привязывать QR-код
            List<string> cellCoordinates = new List<string>();

            if (onActiveSheet)
            {
                // Использование активного листа
                worksheet = excelApp.ActiveSheet;
            }
            else if (onNewSheet)
            {
                // Создание нового листа
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Worksheets.Add();
            }
            else
            {
                // Если не указано иное, используется первый лист
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.Worksheets[1];
            }

            for (int i = 0; i < kolTable; i++)
            {
                int startRow = i * (dataTable.Rows.Count + 2) + 1;

                // Заголовки столбцов
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[startRow, col + 1].Value = dataTable.Columns[col].ColumnName;
                    worksheet.Cells[startRow, col + 1].Font.Bold = true;
                    worksheet.Cells[startRow, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[startRow, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[startRow, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[startRow, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                }

                // Данные строк
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[startRow + row + 1, col + 1].Value = dataTable.Rows[row][col];
                        worksheet.Cells[startRow + row + 1, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        worksheet.Cells[startRow + row + 1, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        worksheet.Cells[startRow + row + 1, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        worksheet.Cells[startRow + row + 1, col + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    }
                }

                // Определение координаты ячейки на одну колонку правее последнего столбца таблицы на уровне заголовка
                int lastColumnIndex = dataTable.Columns.Count + 1; // Последний столбец + 1
                string cellAddress = worksheet.Cells[startRow, lastColumnIndex + 1].Address;
                cellCoordinates.Add(cellAddress);

                // Автоматическое изменение ширины столбцов по содержимому
                worksheet.Columns.AutoFit();
            }

            return cellCoordinates;
        }

    }
}
