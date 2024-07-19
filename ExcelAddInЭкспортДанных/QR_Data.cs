using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using ZXing;

namespace ExcelAddInЭкспортДанных
{
    internal class QR_Data
    {
        // проверить код
        void createQRcode()
        {
            
                // Данные для кодирования
                string data = "текст для кодирования";

                // Настройка параметров генерации
                var options = new ZXing.Common.EncodingOptions
                {
                    Width = 300,
                    Height = 300,
                    Margin = 1
                };

                // Создание объекта кодировщика
                var writer = new BarcodeWriter
                {
                    Format = BarcodeFormat.QR_CODE,
                    Options = options
                };

                // Генерация QR кода
                using (var bitmap = writer.Write(data))
                {
                    // Сохранение QR кода как изображения
                    bitmap.Save("qrcode.png", System.Drawing.Imaging.ImageFormat.Png);
                }

                Console.WriteLine("QR код создан и сохранен как qrcode.png");
           
        }
        void createQRcodeText() 
        {
            // Данные для кодирования
            string data = "https://www.example.com";

            // Настройка параметров генерации
            var options = new ZXing.Common.EncodingOptions
            {
                Width = 300,
                Height = 300,
                Margin = 1
            };

            // Создание объекта кодировщика
            var writer = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE,
                Options = options
            };

            // Генерация QR кода
            using (var qrCodeBitmap = writer.Write(data))
            {
                // Определение размеров изображения с текстом и QR кодом
                int width = qrCodeBitmap.Width;
                int height = qrCodeBitmap.Height + 50; // Дополнительное место для текста
                using (var bitmap = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.Clear(Color.White);

                        // Отрисовка текста
                        using (Font font = new Font("Arial", 20))
                        using (Brush brush = new SolidBrush(Color.Black))
                        {
                            var textSize = g.MeasureString(data, font);
                            g.DrawString(data, font, brush, (width - textSize.Width) / 2, 10);
                        }

                        // Отрисовка QR кода
                        g.DrawImage(qrCodeBitmap, 0, 50);
                    }

                    // Сохранение изображения
                    bitmap.Save("qrcode_with_text.png", System.Drawing.Imaging.ImageFormat.Png);
                }
            }

            Console.WriteLine("QR код с текстом создан и сохранен как qrcode_with_text.png");
        }
    }
    }
}
