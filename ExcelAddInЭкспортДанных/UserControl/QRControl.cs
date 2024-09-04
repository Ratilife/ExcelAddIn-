using ExcelAddInЭкспортДанных.forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media;
using Color = System.Drawing.Color;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddInЭкспортДанных
{
    public partial class QRControl : UserControl
    {
        private ColorComboBox colorComboBox;
        private string filePath { get;  set; }
        public string QRToDate { get; private set; }
        public string QRToDateMany { get; private set; }
        string RangeSelection; 
        CommonMethods cm = new CommonMethods();
        private int size;
        public QRControl()
        {
            InitializeComponent();

            setColorComboBox();
            pbPicture.SizeMode = PictureBoxSizeMode.Zoom;
            ContextMenuPictureBox();
            button();
        }

        #region контекстное меню к PictureBox
        /**
        * Создает и привязывает контекстное меню к PictureBox.
        * В контекстное меню добавляется пункт "Копировать", который вызывает метод CopyImageToClipboard при нажатии.
        */
        private void ContextMenuPictureBox() 
        {
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem copyItem = new ToolStripMenuItem("Копировать");
            copyItem.Click += CopyImageToClipboard;
            contextMenu.Items.Add(copyItem);
            pbPicture.ContextMenuStrip = contextMenu; // Привязываем контекстное меню к PictureBox

        }
        /**
        * Копирует изображение из PictureBox в буфер обмена.
        * Если изображение отсутствует, выводится сообщение об ошибке.
        *
        *   @param sender Объект, вызвавший событие.
        *   @param e Аргументы события.
        */
        private void CopyImageToClipboard(object sender, EventArgs e)
        {
            if (pbPicture.Image != null)
            {
                Clipboard.SetImage(pbPicture.Image); // Копируем изображение в буфер обмена
            }
            else
            {
                MessageBox.Show("Изображение отсутствует в PictureBox.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        void setColorComboBox()
       {
            colorComboBox = new ColorComboBox();
            colorComboBox.InitializeColorComboBox(cbColour);
            colorComboBox.InitializeColorComboBox(cbBackground);

            // Устанавливаем значения по умолчанию
            cbColour.SelectedItem = Color.Black;
            cbBackground.SelectedItem = Color.White;
        }

        private void btPathFolder_Click(object sender, EventArgs e)
        {
           
            //CommonMethods cm = new CommonMethods();
            filePath = cm.dialogFolder();
            txtPathFolder.Text = filePath;

        }

        private void QRControl_Load(object sender, EventArgs e)
        {
            txtPathFolder.Visible = false;
            btPathFolder.Visible=false;
            cm.InitializeTrackBar(tbSize);
            visibilityFor_rbMany();
        }

        private void visibilityFor_rbMany()
        {
            if (rbMany.Checked == true)
            {
                txtPost.Visible = true;
                rbColumnRight.Visible = true;
                rbSpecifyRange.Visible = true;
                btSpecifyRange.Visible = true;
                gbChoice.Visible = true;
            }
            else
            {
                txtPost.Visible = false;
                rbColumnRight.Visible = false;
                rbSpecifyRange.Visible = false;
                btSpecifyRange.Visible = false;
                gbChoice.Visible = false;
            }
        }
        private void cbPictureFile_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPictureFile.Checked)
            {
                txtPathFolder.Visible = true;
                btPathFolder.Visible = true;

            }
            else
            {
                txtPathFolder.Visible = false;
                btPathFolder.Visible = false;
            }

        }

        private void btCreate_Click(object sender, EventArgs e)
        {
             QRToDate = txtQRcode.Text;
             QRcode qr = new QRcode();
             Bitmap qrBitmap;
             // Получаем выбранные цвета из ComboBox
             Color qrColor = (Color)cbColour.SelectedItem;
             Color backgroundColor = (Color)cbBackground.SelectedItem;
             int firstNumber=0;
             int secondNumber=0;
             string newCol="";
             string startCell =null;
             string endCell =null;
             string fp = null;
             string tempFilePath = null;
            List<int> numbers = new List<int>();
            Excel.Worksheet targetSheet = null;
            string fileName = null;

            bool addText = cbAddText.Checked;
            // Получаем текущее значение TrackBar
            size = tbSize.Value;
            //QR - код 
            if (rbOne.Checked)
             {
               
               
                if (cbPictureFile.Checked) 
                { 

                    qrBitmap = qr.CreateQRCodePicture(QRToDate, filePath, qrColor, backgroundColor,size);
                    // Отображаем QR-код в PictureBox
                    pbPicture.Image = qrBitmap;
                }else
                {

                    qrBitmap = qr.CreateQRCode(QRToDate, qrColor, backgroundColor,size, addText);
                    // Отображаем QR-код в PictureBox
                    pbPicture.Image = qrBitmap;
                }

                
             }
            //QR - коды
            if (rbMany.Checked)
             {
               
                string input = txtQRcodes.Text;


                // Получаем текущий рабочий лист
                Excel.Worksheet worksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

                // Разделяем строку диапазона на начальную и конечную ячейки
                string[] cells = input.Split(':');
                // Проверка, что в массиве действительно два элемента
                if (cells.Length == 2)
                {
                     startCell = cells[0];
                     endCell = cells[1];
                }
                else if(cells.Length == 1)
                {
                     startCell = cells[0];
                     endCell = startCell;
                }
                
                //колонка справа
                if (rbColumnRight.Checked)
                {
                    
                    //получаем Имя колонки справа
                     newCol = cm.ShiftCellColumn(cells[0],true);
                    //Получаем номера строк 
                    numbers = cm.ExtractNumbers(input);
                    if(numbers.Count == 2) 
                    { 
                     firstNumber = numbers[0]; // Получаем первый элемент
                     secondNumber = numbers[1]; // Получаем второй элемент
                    }else if(numbers.Count == 1) 
                    {
                        firstNumber = numbers[0];
                    }
                }
                //указать диапазон
                if (rbSpecifyRange.Checked)
                {
                    if (RangeSelection == null)
                    {
                        MessageBox.Show("Диапазон ячеек не может быть неопределенным.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    List<string> rs = cm.ExtractLetters(RangeSelection);
                    newCol = rs[0];
                    numbers = cm.ExtractNumbers(RangeSelection);
                    firstNumber = numbers[0];

                }

                // Определяем диапазон строк
                Excel.Range range = worksheet.get_Range(startCell, endCell);

                int index = 1;
                // Проходим по всем ячейкам в диапазоне
                foreach (Excel.Range cell in range.Cells)
                {
                    // Если значение ячейки пустое, пропустить эту итерацию
                    if (cell.Value2 == null || string.IsNullOrWhiteSpace(cell.Value2.ToString()))
                    {
                        firstNumber = firstNumber + 1;
                        continue;
                    }
                        
                   
                    // Создание QR-кода 
                    qrBitmap = qr.CreateQRCode(cell.Value2, qrColor, backgroundColor, size, addText);
                    //Сохранить в файл
                    if (cbPictureFile.Checked)
                    {
                        fileName = cm.ReplaceInvalidChars(cell.Value2.ToString());
                        fp = System.IO.Path.Combine(filePath, fileName  + ".png");
                        qrBitmap.Save(fp, System.Drawing.Imaging.ImageFormat.Png);
                    }
                    else
                    {
                        //qrBitmap = qr.CreateQRCode(cell.Value2, qrColor, backgroundColor, size);
                        // Создайте временный файл с уникальным именем
                        string tempFileName = "qrcode_temp_" + index.ToString() + ".png";
                        fp = System.IO.Path.Combine(System.IO.Path.GetTempPath(), tempFileName);

                        // Сохраните ваш Bitmap во временный файл
                        qrBitmap.Save(fp, System.Drawing.Imaging.ImageFormat.Png);
                        tempFilePath = fp;
                    }
                    index = index + 1;
                    if (rbSpecifyRange.Checked) 
                    {                            
                        // Проверка, указан ли лист
                        if (RangeSelection.Contains("!"))
                        {
                            // Разделите строку адреса на части
                            string[] addressParts = RangeSelection.Split('!');
                            string sheetName = addressParts[0];
                            // Получить лист по имени
                            targetSheet = worksheet.Parent.Worksheets[sheetName];
                        }
                        else
                        {
                            // Использовать активный лист
                            targetSheet = worksheet;       
                        }
                    }
                    if (rbColumnRight.Checked) 
                    {
                        // Использовать активный лист
                        targetSheet = worksheet;
                    }

                    // Проверка, если ячейка справа пустая
                    //Excel.Range cellRight = targetSheet.Cells[cell.Row, cell.Column + 1];
                    //if (cellRight.Value2 == null || string.IsNullOrWhiteSpace(cellRight.Value2.ToString()))
                    //{
                        Excel.Range cell2 = targetSheet.Range[newCol + firstNumber.ToString()];

                        InsertQRCodeIntoCell(cell2, qrBitmap, fp);
                    //}

                    firstNumber = firstNumber + 1;
                }

                
            }
            

        }
        private void InsertQRCodeIntoCell(Excel.Range cell, Bitmap qrBitmap, string filePath)
        {
            // Получаем размеры изображения
            float imageWidth = qrBitmap.Width * 3;
            float imageHeight = qrBitmap.Height * 3;

            // Преобразование размеров изображения в пункты (1 пункт = 1/72 дюйма)
            float imageWidthInPoints = imageWidth * 72 / qrBitmap.HorizontalResolution;
            float imageHeightInPoints = imageHeight * 72 / qrBitmap.VerticalResolution;

            // Установка ширины и высоты ячейки в соответствии с размером изображения
            cell.ColumnWidth = imageWidthInPoints / 7.0; // Примерное преобразование пунктов в ширину колонки
                                                         // Ограничение высоты строки максимальным значением 409
            cell.RowHeight = Math.Min(imageHeightInPoints, 409);

            // Вставка изображения
            Excel.Shape picture = cell.Worksheet.Shapes.AddPicture(
                filePath,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoCTrue,
                (float)cell.Left,
                (float)cell.Top,
                imageWidthInPoints,
                imageHeightInPoints
            );

            // Если временный файл был создан, удалить его после вставки изображения
            if (filePath.Contains("qrcode_temp_"))
            {
                // Проверьте, существует ли файл
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }
                else
                {
                    // Логирование или сообщение, если файл не был найден
                    Console.WriteLine($"Временный файл {filePath} не найден.");
                }
            }
        }
        private void txtQRcode_KeyDown(object sender, KeyEventArgs e)
        {
            rbOne.Checked = true;
        }

        private void QRControl_BackColorChanged(object sender, EventArgs e)
        {
            
        }
       

        private void btRange_Click(object sender, EventArgs e)
        {
            rbMany.Checked = true;
            string range =  cm.SelectRange();
            txtQRcodes.Text= range;
            if (rbColumnRight.Checked== true) 
            {
               // RangeSelection = cm.ShiftRange(range);
                btSpecifyRange.Enabled = true;
            }
        }

        private void tbSize_Scroll(object sender, EventArgs e)
        {
            // Получаем текущее значение TrackBar
            size = tbSize.Value;
        }

        private void rbMany_CheckedChanged(object sender, EventArgs e)
        {
            visibilityFor_rbMany();
        }

        private void btSpecifyRange_Click(object sender, EventArgs e)
        {
            string range = cm.SelectRange();
            RangeSelection = range;
        }

        #region КнопкаС_ВыпадающимСписком
        private void button() 
        {
            // Создаем ContextMenuStrip для выпадающего списка
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("Структура для основного средства", null, PrintOption1_Click);
            contextMenu.Items.Add("Структура созданная пользователем", null, PrintOption2_Click);


            // Создаем кнопку "Создать структуру JSON"

            bt_JSON.Text = "Создать структуру JSON";
            bt_JSON.Click += (sender, e) => contextMenu.Show(bt_JSON, new System.Drawing.Point(0, bt_JSON.Height));
            
        }
        private void openForm(string text,string parameter)
        {
            //TODO:  определится как передать форме какую структуру заполнять по шаблону или пользователь сам оформит
            FormDialogTableStructureJASON_Sample form = new FormDialogTableStructureJASON_Sample(parameter);
            form.Text = text;
            form.ShowDialog();
        }
        private void PrintOption1_Click(object sender, EventArgs e)
        {

            string text = "Форма диалога для формирования структуры json по шаблону - основные средства";

            openForm(text,"ОсновныеСредства");

        }

        private void PrintOption2_Click(object sender, EventArgs e)
        {
            string text = "Форма диалога для формирования структуры json сформированная пользователем";
            openForm(text, "Пользователь");
        }
        #endregion
    }
}
