using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInЭкспортДанных
{
    internal class WorkingWithTables
    {


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
            Microsoft.Office.Interop.Excel.Range endCell = startCell.get_Offset(rowCount , columnCount - 1);
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

    }
}
