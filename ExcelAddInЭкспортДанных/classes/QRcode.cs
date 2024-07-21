using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZXing;

namespace ExcelAddInЭкспортДанных
{
    internal class QRcode
    {
        public QRcode() { }
        public QRcode(string code) { }
        private void createQR_code() 
        {
        
        }

        public void CreateQRCodePicture(string data, string pathFolder, string counter="")
        {
            // Настройка параметров генерации
            var options = new ZXing.QrCode.QrCodeEncodingOptions
            {
                Width = 300,
                Height = 300,
                Margin = 1,
                // Устанавливаем кодировку в UTF-8 для поддержки кириллицы
                CharacterSet = "UTF-8"
            };

            // Создание объекта кодировщика
            var writer = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE,
                Options = options
            };

            // Проверка существования папки и создание, если она не существует
            if (!System.IO.Directory.Exists(pathFolder))
            {
                System.IO.Directory.CreateDirectory(pathFolder);
            }

            // Генерация QR кода
            using (var bitmap = writer.Write(data))
            {
                // Сформировать полный путь к файлу
                string filePath = System.IO.Path.Combine(pathFolder, "qrcode"+counter+".png");

                // Сохранение QR кода как изображения
                bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
            }
        }

    }
}
