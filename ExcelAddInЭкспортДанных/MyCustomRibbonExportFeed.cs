using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddInЭкспортДанных
{
    public partial class MyCustomRibbonExportFeed
    {
        private void MyCustomRibbonExportFeed_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //основной метод 
        void ExportXlsxToCsv(string xlsxPath, string csvPath, System.Text.Encoding encoding)
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
                    // Получение данных из формы
                }
            }
        }
    }
}
