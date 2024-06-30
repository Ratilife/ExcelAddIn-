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

        private void button1_ClickДляПроверкиВременный(object sender, RibbonControlEventArgs e)
        {
            // Открытие файла XLSX через диалоговое окно
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                // Установка фильтра для выбора только файлов Excel
                Filter = "Excel Files|*.xlsx",
                // Заголовок диалогового окна
                Title = "Select an Excel File"
            };
            // Проверка, была ли нажата кнопка "OK" в диалоговом окне
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Получение пути к выбранному файлу
                string xlsxPath = openFileDialog.FileName;

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
                    // Вызов метода для экспорта данных из XLSX в CSV
                    ExportXlsxToCsv(xlsxPath, csvPath, System.Text.Encoding.UTF8);
                }
            }
        }
        void ExportXlsxToCsv(string xlsxPath, string csvPath, System.Text.Encoding encoding)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;

            if (excelApp == null)
            {
                MessageBox.Show("Error: Unable to access Excel application.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Excel.Workbook workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Open(xlsxPath);
                Excel.Worksheet worksheet = workbook.Sheets[1];
                int rowCount = worksheet.UsedRange.Rows.Count;
                int colCount = worksheet.UsedRange.Columns.Count;

                using (StreamWriter writer = new StreamWriter(csvPath, false, encoding))
                {
                    for (int row = 1; row <= rowCount; row++)
                    {
                        string[] rowData = new string[colCount];
                        for (int col = 1; col <= colCount; col++)
                        {
                            rowData[col - 1] = worksheet.Cells[row, col].Text.ToString();
                        }
                        writer.WriteLine(string.Join(";", rowData));
                    }
                }

                MessageBox.Show("Export successful!", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Обработка ошибок
                MessageBox.Show("Error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                }
            }
        }
    }
}
