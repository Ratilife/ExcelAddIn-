using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelAddInЭкспортДанных
{
    internal class WorkingJSON
    {
        public string InventoryNumber { get; set; }
        public DateTime AcquisitionDate { get; set; }
        public DateTime LastMaintenanceDate { get; set; }
        public float Cost { get; set; }
        public int FaktNumber { get; set; }
        public int YearOfRelease { get; set; }
        public string SerialNumber { get; set; }
        public string RegistrationDocument { get; set; }
        public string Location { get; set; }
        public string Responsible { get; set; }
        public string TypeFA { get; set; }
        public string Manufacturer { get; set; }
        public string ModelFA { get; set; }
        public DateTime NextMaintenanceDate { get; set; }

        public List<Dictionary<string, string>> generateTable()
        {
            var table = new List<Dictionary<string, string>>
            {
                new Dictionary<string, string>
                {
                    { "Название колонки", "Инвентарный номер" },
                    { "Название поля JSON", "inventory_number" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Дата приобретения актива" },
                    { "Название поля JSON", "acquisition_date" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Дата последнего обслуживания" },
                    { "Название поля JSON", "last_maintenance_date" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Стоимость" },
                    { "Название поля JSON", "cost" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Фактическое наличие" },
                    { "Название поля JSON", "fakt_number" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Год выпуска" },
                    { "Название поля JSON", "year_of_release" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Серийный номер" },
                    { "Название поля JSON", "serial_number" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Паспорт (документ о регистрации)" },
                    { "Название поля JSON", "registration_document" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Местонахождение" },
                    { "Название поля JSON", "location" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Ответственный" },
                    { "Название поля JSON", "responsible" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Тип основного средства" },
                    { "Название поля JSON", "typeFA" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Производитель" },
                    { "Название поля JSON", "manufacturer" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Модель" },
                    { "Название поля JSON", "modelFA" }
                },
                new Dictionary<string, string>
                {
                    { "Название колонки", "Дата следующего обслуживания" },
                    { "Название поля JSON", "next_maintenance_date" }
                }
            };

            return table;
        }
    }
}
