using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInЭкспортДанных.classes
{
    internal class ImportData
    {
        #region импорт данных из JSON

        /// <summary>
        /// Импортирует данные из JSON файла в активную книгу Excel.
        /// Автоматически определяет структуру JSON и создает таблицу с соответствующими колонками.
        /// </summary>
        /// <param name="jsonFilePath">Путь к JSON файлу</param>
        /// <param name="createNewSheet">true - создать новый лист, false - использовать активный лист</param>
        public void ImportJsonToExcelInActiveWorkbook(string jsonFilePath, bool createNewSheet = false)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;

            if (excelApp == null)
            {
                MessageBox.Show("Error: Не удается получить доступ к приложению Excel.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                if (!File.Exists(jsonFilePath))
                {
                    MessageBox.Show("Error: Файл не найден: " + jsonFilePath, "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                if (workbook == null)
                {
                    MessageBox.Show("Error: Нет активной книги Excel.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Excel.Worksheet worksheet;

                if (createNewSheet)
                {
                    worksheet = (Excel.Worksheet)workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                    worksheet.Name = "ImportedJSON";
                }
                else
                {
                    worksheet = workbook.ActiveSheet;
                    if (worksheet == null)
                    {
                        MessageBox.Show("Error: Нет активного листа Excel.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                string jsonContent = File.ReadAllText(jsonFilePath);
                List<Dictionary<string, object>> jsonData = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonContent);

                if (jsonData == null || jsonData.Count == 0)
                {
                    MessageBox.Show("Error: JSON файл пуст или имеет неверный формат.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                HashSet<string> allKeys = new HashSet<string>();
                foreach (var row in jsonData)
                {
                    foreach (var key in row.Keys)
                    {
                        allKeys.Add(key);
                    }
                }

                List<string> sortedKeys = allKeys.OrderBy(k => k).ToList();

                //int headerRow = 1;
                int colIndex = 1;
                //foreach (string key in sortedKeys)
                //{
                //    worksheet.Cells[headerRow, colIndex] = key;
                //    colIndex++;
                //}

                int dataRow = 2;
                foreach (var rowData in jsonData)
                {
                    colIndex = 1;
                    foreach (string key in sortedKeys)
                    {
                        object value = rowData.ContainsKey(key) ? rowData[key] : "";
                        worksheet.Cells[dataRow, colIndex] = value?.ToString() ?? "";
                        colIndex++;
                    }
                    dataRow++;
                }

                //Excel.Range headerRange = worksheet.Range[worksheet.Cells[headerRow, 1], worksheet.Cells[headerRow, sortedKeys.Count]];
                //headerRange.Font.Bold = true;
                //headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                worksheet.Columns.AutoFit();

                MessageBox.Show($"Импорт данных завершен успешно!\nИмпортировано строк: {jsonData.Count}\nКолонок: {sortedKeys.Count}",
                    "Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (JsonException ex)
            {
                MessageBox.Show("Error: Ошибка при чтении JSON файла. " + ex.Message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

    }
}
