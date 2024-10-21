using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;


namespace ExcelAddInЭкспортДанных
{
    internal class WorkingJSON
    {
        

        public List<Dictionary<string, string>> generateTable()
        {
            var table = new List<Dictionary<string, string>>
            {
                 new Dictionary<string, string>
                {
                    { "Название колонки", "Название основного средства" },
                    { "Название поля JSON", "name" },
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Инвентарный номер" },
                    { "Название поля JSON", "inventory_number" },
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Дата приобретения актива" },
                    { "Название поля JSON", "acquisition_date" },
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Дата последнего обслуживания" },
                    { "Название поля JSON", "last_maintenance_date" },
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Стоимость" },
                    { "Название поля JSON", "cost" }, 
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Фактическое наличие" },
                    { "Название поля JSON", "fakt_number" },
                    { "Значение", "" }

                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Год выпуска" },
                    { "Название поля JSON", "year_of_release" }, 
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Серийный номер" },
                    { "Название поля JSON", "serial_number" }, 
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Паспорт (документ о регистрации)" },
                    { "Название поля JSON", "registration_document" },
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Местонахождение" },
                    { "Название поля JSON", "location" }, 
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Ответственный" },
                    { "Название поля JSON", "responsible" },
                    { "Значение", "" }

                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Тип основного средства" },
                    { "Название поля JSON", "typeFA" }, 
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Производитель" },
                    { "Название поля JSON", "manufacturer" },
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Модель" },
                    { "Название поля JSON", "modelFA" }, 
                    { "Значение", "" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Дата следующего обслуживания" },
                    { "Название поля JSON", "next_maintenance_date" }, 
                    { "Значение", "" }
                }
            };

            return table;
        }


        /*
         * Метод для чтения данных из активного листа Excel, разделенных на таблицы, и преобразования их в JSON.
         * @param worksheet Активный лист Excel, с которым происходит работа.
         * @return Список JSON объектов, созданных на основе данных из таблиц.
         */
        public List<Dictionary<string, string>> createJSON(Excel.Worksheet worksheet)
        {
            //TODO: Проверить целесообразность этого метода. Устраивает ожидания
            // Список для хранения JSON объектов
            List<Dictionary<string, string>> jsonList = new List<Dictionary<string, string>>();

            // Получение диапазона используемых ячеек на листе
            Excel.Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            // Переменная для хранения текущей таблицы
            Dictionary<string, string> currentTable = null;

            // Проход по всем строкам листа
            for (int row = 1; row <= rowCount; row++)
            {
                // Чтение значений из колонок "Название поля JSON" и "Значение"
                var jsonFieldName = (string)(usedRange.Cells[row, 2] as Excel.Range).Text;
                var value         = (string)(usedRange.Cells[row, 3] as Excel.Range).Text;

                // Проверка на пустую строку, которая разделяет таблицы
                if (string.IsNullOrWhiteSpace(jsonFieldName) && string.IsNullOrWhiteSpace(value))
                {
                    // Если текущая таблица не пуста, добавляем ее в список JSON и сбрасываем переменную
                    if (currentTable != null)
                    {
                        jsonList.Add(currentTable);
                        currentTable = null;
                    }
                }
                else
                {
                    // Если текущая таблица пуста, инициализируем новую таблицу
                    if (currentTable == null)
                    {
                        currentTable = new Dictionary<string, string>();
                    }
                    // Добавляем данные в текущую таблицу
                    currentTable[jsonFieldName] = value;
                }
            }

            // Добавляем последнюю таблицу в список JSON, если она не пуста
            if (currentTable != null)
            {
                jsonList.Add(currentTable);
            }

            return jsonList;
        }

    }
}
