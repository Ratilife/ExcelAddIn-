using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using ZXing;
using Color = System.Drawing.Color;
using Font =  System.Drawing.Font;
using Point = System.Drawing.Point;
using ZXing.QrCode;
using ZXing.Rendering;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInЭкспортДанных
{
    internal class QRcode
    {
        public QRcode() { }
        //public QRcode(string code) { }

        public Bitmap CreateQRCodePicture(string data, string pathFolder, Color qrColor, Color backgroundColor, int size, string counter="")
        {
            Bitmap bitmap;
            // Настройка параметров генерации
            var options = new ZXing.QrCode.QrCodeEncodingOptions
            
            {
                Width = size,
                Height = size,
                Margin = 1,
                // Устанавливаем кодировку в UTF-8 для поддержки кириллицы
                CharacterSet = "UTF-8"
            };

            // Создание объекта кодировщика
            var writer = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE,
                Options = options,
                Renderer = new ZXing.Rendering.BitmapRenderer
                {
                    Foreground = qrColor,
                    Background = backgroundColor
                }
            };

            // Проверка существования папки и создание, если она не существует
            if (!System.IO.Directory.Exists(pathFolder))
            {
                System.IO.Directory.CreateDirectory(pathFolder);
            }

            
            // Генерация QR кода
            using (bitmap = writer.Write(data))
            {
                // Сформировать полный путь к файлу
                string filePath = System.IO.Path.Combine(pathFolder, "qrcode"+counter+".png");

                // Сохранение QR кода как изображения
                bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
            }
            return bitmap;
        }
        
        public Bitmap CreateQRCode(string data, Color qrColor, Color backgroundColor, int size, bool addText)
        {
            if (string.IsNullOrWhiteSpace(data))
            {
                throw new ArgumentException("Данные для генерации QR-кода не могут быть нулевыми или пустыми.", nameof(data));
            }

            // Настройка параметров генерации
            var options = new QrCodeEncodingOptions
            {
                Width = size,
                Height = size,
                Margin = 1,
                CharacterSet = "UTF-8"
            };

            // Создание объекта кодировщика
            var writer = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE,
                Options = options,
                Renderer = new BitmapRenderer
                {
                    Foreground = qrColor,
                    Background = backgroundColor
                }
            };

            // Генерация QR-кода
            Bitmap qrBitmap = writer.Write(data);

            // Если нужно добавить текст, рисуем его на изображении
            if (addText)
            {
                // Определяем шрифт и параметры для текста
                Font textFont = new Font("Arial", 6, FontStyle.Bold);
                List<string> lines = new List<string>();

                using (Graphics g = Graphics.FromImage(qrBitmap))
                {
                    float maxWidth = qrBitmap.Width; // Максимальная ширина текста

                    // Разбиваем текст на строки, если он не помещается по ширине
                    string[] words = data.Split(' ');
                    StringBuilder currentLine = new StringBuilder();
                    foreach (var word in words)
                    {
                        if (g.MeasureString(currentLine + word, textFont).Width > maxWidth)
                        {
                            lines.Add(currentLine.ToString());
                            currentLine.Clear();
                        }
                        currentLine.Append(word + " ");
                    }
                    lines.Add(currentLine.ToString().Trim()); // добавляем последнюю строку

                    // Определяем размеры нового изображения с местом для текста
                    int newHeight = qrBitmap.Height + (int)(g.MeasureString("A", textFont).Height * lines.Count) + 5;

                    // Создаем новое изображение с увеличенной высотой
                    Bitmap bitmapWithText = new Bitmap(qrBitmap.Width, newHeight);
                    g.Dispose();

                    using (Graphics g2 = Graphics.FromImage(bitmapWithText))
                    {
                        // Закрашиваем фон
                        g2.Clear(backgroundColor);

                        // Рисуем QR-код в верхней части нового изображения
                        g2.DrawImage(qrBitmap, new Point(0, 0));

                        // Рисуем каждую строку текста под QR-кодом
                        float yPosition = qrBitmap.Height + 5;
                        foreach (var line in lines)
                        {
                            SizeF textSize = g2.MeasureString(line, textFont);
                            PointF textPosition = new PointF((bitmapWithText.Width - textSize.Width) / 2, yPosition);
                            g2.DrawString(line, textFont, new SolidBrush(qrColor), textPosition);
                            yPosition += textSize.Height;
                        }

                        // Возвращаем изображение с текстом
                        return bitmapWithText;
                    }
                }
            }

            // Возвращаем QR-код без текста
            return qrBitmap;
        }




    }
}
