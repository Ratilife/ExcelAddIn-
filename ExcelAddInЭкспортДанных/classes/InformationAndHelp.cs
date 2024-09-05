using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInЭкспортДанных.classes
{
    internal class InformationAndHelp
    {
        public string inf1 { get; set; } = "На лисне где сформированны таблицы для создания QR кода посторонние данные нужно убрать";
        public string inf2 { get; set; } = "QR код будет создаваться с активного листа.Перейдите на лист с данными для формирования QR кода";
        public string inf3 { get; set; } = "Будет создан QR код со структурой JSON";
        public string inf4 { get; set; } = "Будет создан QR код со структурой TEXT";

        public void InstructionJSON()
        {
            //TODO: Инструкция по формированию
            //Создать текст с инструкцией, как формировать данные под JSON для формирования QR кода
        }
    }
}
