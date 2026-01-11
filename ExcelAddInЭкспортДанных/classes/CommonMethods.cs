using Microsoft.Office.Core;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;
using System.Drawing;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;


namespace ExcelAddInЭкспортДанных
{
    /* Общие методы */
    internal class CommonMethods
    {
        public string dialogFolder()
        {
            string filePath;
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowNewFolderButton = true;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath = dialog.SelectedPath + "\\";
            }
            else
            {
                filePath = null;
            }
            return filePath;
        }
        public string dialogFile() 
        {
            string filePath;
            SaveFileDialog dialog = new SaveFileDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath = dialog.FileName;
            }
            else
            {
                filePath = null;
            }

                return filePath;
        }
        public string OpenCsvFile()
        {
            string filePath = null;
            System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog();

            // Устанавливаем фильтр для файлов CSV
            dialog.Filter = "CSV files (*.csv)|*.csv";
            dialog.Title = "Выберите CSV файл";

            // Показываем диалоговое окно
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePath = dialog.FileName; // Получаем путь к выбранному файлу
            }

            return filePath;
        }
        //TODO проверить в работе метод OpenFolder()
        public string OpenFolder()
        {
            string folderPath = null;
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Выберите папку";

                // Показываем диалоговое окно
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    folderPath = dialog.SelectedPath; // Получаем путь к выбранной папке
                }
            }

            return folderPath;
        }


        /*public string SelectRange()
        {
            Excel.Application excelApp;
            object inputBoxResult;

            try
            {
                // Получение текущего экземпляра приложения Excel
                excelApp = Globals.ThisAddIn.Application;
            }
            catch (COMException)
            {
                MessageBox.Show("Excel не открыт");
                return null;
            }

            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                MessageBox.Show("Нет активной книги");
                return null;
            }

            Excel.Worksheet worksheet = workbook.ActiveSheet;
            if (worksheet == null)
            {
                MessageBox.Show("Нет активного листа");
                return null;
            }

            inputBoxResult = excelApp.InputBox("Выберите диапазон", Type: 8);
            if (inputBoxResult is bool && (bool)inputBoxResult == false)
            {
                MessageBox.Show("Диапазон не выбран");
                return null;
            }

            Excel.Range selectedRange = (Excel.Range)inputBoxResult;
            string selectedAddress = selectedRange.get_Address();
            //MessageBox.Show("Выбранный диапазон: " + selectedAddress);
            return selectedAddress;
        }*/
        //удалить
        /*public bool IsDarkTheme()
        {
            const string registryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            const string registryValueName = "AppsUseLightTheme";

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(registryKeyPath))
            {
                if (key != null)
                {
                    object registryValueObject = key.GetValue(registryValueName);
                    if (registryValueObject != null)
                    {
                        int registryValue = (int)registryValueObject;

                        return registryValue == 0; // 0 означает темную тему
                    }
                }
            }

            // Если значение реестра отсутствует или недоступно, считаем тему светлой
            return false;
        }*/

        public string SelectRange()
        {
            Excel.Application excelApp;
            object inputBoxResult;

            try
            {
                // Получение текущего экземпляра приложения Excel
                excelApp = Globals.ThisAddIn.Application;
            }
            catch (COMException)
            {
                MessageBox.Show("Excel не открыт");
                return null;
            }

            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                MessageBox.Show("Нет активной книги");
                return null;
            }

            Excel.Worksheet activeSheet = workbook.ActiveSheet;
            if (activeSheet == null)
            {
                MessageBox.Show("Нет активного листа");
                return null;
            }

            inputBoxResult = excelApp.InputBox("Выберите диапазон", Type: 8);
            if (inputBoxResult is bool && (bool)inputBoxResult == false)
            {
                MessageBox.Show("Диапазон не выбран");
                return null;
            }

            Excel.Range selectedRange = (Excel.Range)inputBoxResult;

            // Проверка, что диапазон не превышает один столбец
            if (selectedRange.Columns.Count > 1)
            {
                MessageBox.Show("Нельзя выбирать диапазон больше одного столбца");
                return null;
            }

            string selectedAddress = selectedRange.get_Address();

            // Определяем имя листа, на котором выделен диапазон
            Excel.Worksheet selectedSheet = selectedRange.Worksheet;
            if (selectedSheet != activeSheet)
            {
                selectedAddress = selectedSheet.Name + "!" + selectedAddress;
            }

            return selectedAddress;
        }


        public void InitializeTrackBar(TrackBar tb)
        {
            tb.Minimum = 50;  // Минимальное значение
            tb.Maximum = 500;  // Максимальное значение
            tb.Value = 100;    // Начальное значение
            tb.TickFrequency = 20;  // Частота меток (опционально)
            tb.SmallChange = 10;    // Малое изменение (опционально)
            tb.LargeChange = 20;    // Большое изменение (опционально)
        }

        /**
       * Заменяет недопустимые символы в имени файла на указанный символ замены.
       *
       *   @param input Входная строка, которую нужно проверить и изменить.
       *   @param replacement Символ, на который будут заменены недопустимые символы. По умолчанию '_'.
       *   @return Новая строка с замененными недопустимыми символами.
       *
       * Этот метод проверяет входную строку на наличие символов, которые нельзя использовать в названиях файлов в Windows,
       * и заменяет их на указанный символ замены. Если символ замены не указан, используется нижнее подчеркивание ('_').
       */
        public string ReplaceInvalidChars(string input, char replacement = '_')
        {
            string invalidChars = new string(System.IO.Path.GetInvalidFileNameChars());
            foreach (char c in invalidChars)
            {
                input = input.Replace(c, replacement);
            }
            return input;
        }


        /**
        * Метод ShiftRange изменяет столбцы в диапазоне ячеек Excel и возвращает новый диапазон.
        *
        * @param range - исходный диапазон в формате "A1:B2".
        * @return Новый диапазон с измененными столбцами в формате "B1:C2".
        * @throws ArgumentException если формат исходного диапазона неверный.
        *
        * Метод выполняет следующие шаги:
        *   1. Разделяет строку диапазона на начальную и конечную ячейки.
        *   2. Проверяет корректность формата диапазона. Если строка диапазона не содержит ровно две ячейки,
        *      выбрасывается исключение ArgumentException.
        *   3. Изменяет столбцы начальной и конечной ячеек с помощью метода ShiftCellColumn.
        *   4. Формирует и возвращает новый диапазон, состоящий из измененных ячеек.
        *
        * Пример использования:
        *   string newRange = ShiftRange("G4:G71");
        *   // newRange будет содержать измененный диапазон, например, "H4:H71" если ShiftCellColumn
        *   // изменяет столбцы на один вправо.
        */
        public string ShiftRange(string range)
        {
            // Разделение диапазона на начальную и конечную ячейки
            var cells = range.Split(':');
            if (cells.Length != 2)
            {
                throw new ArgumentException("Неверный формат диапазона");
            }

            // Изменение столбцов начальной и конечной ячеек
            string newStartCell = ShiftCellColumn(cells[0]);
            string newEndCell = ShiftCellColumn(cells[1]);

            // Формирование нового диапазона
            return $"{newStartCell}:{newEndCell}";
        }

        /**
        * Метод ShiftCellColumn изменяет столбец указанной ячейки Excel, смещая его на одну позицию вправо.
        *
        * @param cell - исходный адрес ячейки в формате "A1", "B2" и т.д.
        * @return Новый адрес ячейки с измененным столбцом, например, "B1" для входного "A1".
        * @throws ArgumentException если формат исходного адреса ячейки неверный.
        *
        * Метод выполняет следующие шаги:
        *   1. Использует регулярное выражение для извлечения буквенного обозначения столбца и числового обозначения строки из адреса ячейки.
        *   2. Проверяет корректность формата адреса ячейки. Если регулярное выражение не находит совпадений, выбрасывается исключение ArgumentException.
        *   3. Извлекает части адреса ячейки: буквы столбца и цифры строки.
        *   4. Преобразует буквенное обозначение столбца в числовое значение с помощью метода ConvertColumnLettersToNumber.
        *   5. Смещает числовое значение столбца на один вправо.
        *   6. Преобразует измененное числовое значение столбца обратно в буквенное обозначение с помощью метода ConvertColumnNumberToLetters.
        *   7. Формирует и возвращает новый адрес ячейки с измененным столбцом и исходной строкой.
        *
        * Пример использования:
        *   string newCell = ShiftCellColumn("G4");
        *   // newCell будет содержать "H4", если "G" смещается на один столбец вправо.
        */
        public string ShiftCellColumn(string cell, bool colum=false)
        {
            // Регулярное выражение для извлечения букв и цифр из адреса ячейки
            var match = System.Text.RegularExpressions.Regex.Match(cell, @"\$?([A-Z]+)\$?(\d+)");
            if (!match.Success)
            {
                throw new ArgumentException("Неверный формат ячейки");
            }

            // Извлечение частей адреса ячейки
            string columnLetters = match.Groups[1].Value;
            string rowNumbers = match.Groups[2].Value;

            // Преобразование буквенного обозначения столбца в числовое значение
            int columnNumber = ConvertColumnLettersToNumber(columnLetters);

            // Смещение столбца на один вправо
            columnNumber++;

            // Преобразование числового значения столбца обратно в буквенное
            string newColumnLetters = ConvertColumnNumberToLetters(columnNumber);
            if(colum)
            {
                return newColumnLetters;
            }
            else 
            {
                // Возвращение нового адреса ячейки
                return $"{newColumnLetters}{rowNumbers}";
            }
        }

        /**
        * Метод ConvertColumnLettersToNumber преобразует буквенное обозначение столбца Excel в числовое значение.
        *
        * @param columnLetters - буквенное обозначение столбца, например, "A", "B", "AA".
        * @return Числовое значение столбца, где "A" = 1, "B" = 2, "AA" = 27 и т.д.
        *
        * Метод выполняет следующие шаги:
        *   1. Инициализирует переменную sum для хранения итогового числового значения.
        *   2. Проходит по каждому символу в строке columnLetters.
        *   3. Умножает текущее значение sum на 26, чтобы сдвинуть разрядность на одну позицию влево.
        *   4. Добавляет числовое значение текущего символа, вычитая 'A' и прибавляя 1, чтобы получить правильное значение (например, 'A' -> 1, 'B' -> 2).
        *   5. Возвращает итоговое числовое значение столбца.
        *
        * Пример использования:
        *   int columnNumber = ConvertColumnLettersToNumber("AB");
        *   // columnNumber будет содержать 28, так как "AB" соответствует столбцу 28.
        */
        private int ConvertColumnLettersToNumber(string columnLetters)
        {
            int sum = 0;
            foreach (char c in columnLetters)
            {
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum;
        }

        /**
        * Метод ConvertColumnNumberToLetters преобразует числовое значение столбца Excel в буквенное обозначение.
        *
        * @param columnNumber - числовое значение столбца, например, 1, 2, 27.
        * @return Буквенное обозначение столбца, где 1 = "A", 2 = "B", 27 = "AA" и т.д.
        *
        * Метод выполняет следующие шаги:
        *   1. Инициализирует пустую строку columnLetters для хранения результирующего буквенного обозначения столбца.
        *   2. Использует цикл while для обработки columnNumber до тех пор, пока его значение больше нуля.
        *   3. Уменьшает значение columnNumber на 1, чтобы учесть смещение, так как 'A' соответствует 1, а не 0.
        *   4. Вычисляет текущий символ столбца с помощью остатка от деления columnNumber на 26 и добавляет его к началу строки columnLetters.
        *   5. Делит columnNumber на 26, чтобы перейти к следующему разряду.
        *   6. Возвращает итоговое буквенное обозначение столбца.
        *
        * Пример использования:
        *   string columnLetters = ConvertColumnNumberToLetters(28);
        *   // columnLetters будет содержать "AB", так как 28 соответствует столбцу "AB".
        */
        private string ConvertColumnNumberToLetters(int columnNumber)
        {
            string columnLetters = "";
            while (columnNumber > 0)
            {
                columnNumber--;
                columnLetters = (char)('A' + (columnNumber % 26)) + columnLetters;
                columnNumber /= 26;
            }
            return columnLetters;
        }

        private bool IsSingleCellSelected(object inputBoxResult)
        {
            // Проверяем, что inputBoxResult не является булевым значением false
            if (inputBoxResult is bool && (bool)inputBoxResult == false)
            {
                return false;
            }

            Excel.Range selectedRange = inputBoxResult as Excel.Range;

            // Проверяем, что выбранная область состоит из одной ячейки
            if (selectedRange != null && selectedRange.Cells.Count == 1)
            {
                return true;
            }

            return false;
        }


        //Перенести в конспект
        /**
        * Метод ExtractNumbers извлекает все числа из входной строки и возвращает их в виде списка целых чисел.
        *
        * @param input - входная строка, из которой будут извлекаться числа.
        * @return Список целых чисел, найденных в входной строке.
        *
        * Метод выполняет следующие шаги:
        * 1.    Инициализирует пустой список result для хранения найденных чисел.
        * 2.    Определяет шаблон регулярного выражения pattern для поиска одной или более цифр.
        * 3.    Создает объект Regex с указанным шаблоном.
        * 4.    Использует метод Matches объекта Regex для поиска всех совпадений в строке input и возвращает коллекцию MatchCollection.
        * 5.    Проходит по всем совпадениям в коллекции matches и добавляет каждое совпадение в список result, предварительно преобразовав его в целое число с помощью int.Parse.
        * 6.    Возвращает список result, содержащий все найденные числа.
        *
        * Пример использования:
        *   List<int> numbers = ExtractNumbers("abc123def456ghi789");
        *   // numbers будет содержать [123, 456, 789], так как они являются единственными числами в строке "abc123def456ghi789".
        */

        public List<int> ExtractNumbers(string input)
        {
            List<int> result = new List<int>();

            // Регулярное выражение для поиска чисел
            string pattern = @"\d+";
            Regex regex = new Regex(pattern);

            // Поиск всех совпадений в строке
            MatchCollection matches = regex.Matches(input);
           
                // Добавление найденных чисел в список
                foreach (Match match in matches)
            {
                result.Add(int.Parse(match.Value));
            }


            return result;
        }
        
        /**
        * Метод ExtractLetters извлекает все последовательности букв из входной строки и возвращает их в виде списка.
        *
        * @param input - входная строка, из которой будут извлекаться буквы.
        * @return Список строк, каждая из которых содержит последовательность букв, найденную в входной строке.
        *
        * Метод выполняет следующие шаги:
        *   1. Инициализирует пустой список result для хранения найденных последовательностей букв.
        *   2. Определяет шаблон регулярного выражения pattern для поиска одной или более букв (как заглавных, так и строчных).
        *   3. Создает объект Regex с указанным шаблоном.
        *   4. Использует метод Matches объекта Regex для поиска всех совпадений в строке input и возвращает коллекцию MatchCollection.
        *   5. Проходит по всем совпадениям в коллекции matches и добавляет каждое совпадение в список result.
        *   6. Возвращает список result, содержащий все найденные последовательности букв.
        *
        * Пример использования:
        *   List<string> letters = ExtractLetters("Hello123World");
        *   // letters будет содержать ["Hello", "World"], так как они являются единственными последовательностями букв в строке "Hello123World".
        */
        public List<string> ExtractLetters(string input)
        {
            List<string> result = new List<string>();

            // Регулярное выражение для поиска букв
            string pattern = @"[A-Za-z]+";
            Regex regex = new Regex(pattern);

            // Поиск всех совпадений в строке
            MatchCollection matches = regex.Matches(input);

            // Добавление найденных букв в список
            foreach (Match match in matches)
            {
                result.Add(match.Value);
            }

            return result;
        }

        /**
        * Метод defineСellsQRcode определяет координаты ячеек для привязки QR-кодов на активном листе Excel.
        * 
        * Метод проходит по всем строкам активного листа, определяет таблицы по жирным заголовкам и записывает
        * адреса ячеек, находящихся на одну колонку правее последнего столбца таблицы на уровне заголовка, в список cellCoordinates.
        * 
        * @return List<string> - список адресов ячеек для привязки QR-кодов.
        */

        public List<string> defineСellsQRcode()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = excelApp.ActiveSheet;
            List<string> cellCoordinates = new List<string>();

            // Определение существующих таблиц на активном листе
            int existingTablesCount = 0;
            int lastRow = worksheet.UsedRange.Rows.Count;   // Последняя заполненная строка

            for (int row = 1; row <= lastRow; row++)
            {
                if (worksheet.Cells[row, 1].Value != null && worksheet.Cells[row, 1].Font.Bold)
                {
                    existingTablesCount++;
                    int tableRowCount = 0;

                    // Определение количества строк в таблице
                    for (int r = row + 1; r <= lastRow; r++)
                    {
                        if (worksheet.Cells[r, 1].Value == null || worksheet.Cells[r, 1].Font.Bold)
                        {
                            break;
                        }
                        tableRowCount++;
                    }

                    // Определение координаты ячейки на одну колонку правее последнего столбца таблицы на уровне заголовка
                    int lastColumnIndex = worksheet.UsedRange.Columns.Count;
                    string cellAddress = worksheet.Cells[row, lastColumnIndex + 2].Address;
                    cellCoordinates.Add(cellAddress);

                    // Пропуск строк таблицы
                    row += tableRowCount;
                }
            }

            return cellCoordinates;
        }

        

    }
}
