using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelAddInЭкспортДанных
{
    internal class ColorComboBox : ComboBox
    {
        public ColorComboBox()
        {
         
        }
        public void InitializeColorComboBox(ComboBox comboBox)
        {
            // Установите стиль рисования для ComboBox
            comboBox.DrawMode = DrawMode.OwnerDrawFixed;
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList;

            // Добавьте стандартные цвета
            comboBox.Items.Add(Color.Black);
            comboBox.Items.Add(Color.White);
            comboBox.Items.Add(Color.Red);
            comboBox.Items.Add(Color.Green);
            comboBox.Items.Add(Color.Blue);

            // Подключите обработчик события DrawItem
            comboBox.DrawItem += new DrawItemEventHandler(ComboBox_DrawItem);
        }

        private void ComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            // Вызываем базовый метод для рисования фона
            e.DrawBackground();

            // Если нет элементов, выходим
            if (e.Index < 0) return;

            // Получаем цвет из элементов
            Color color = (Color)((ComboBox)sender).Items[e.Index];

            // Создаем кисть с нужным цветом
            using (SolidBrush brush = new SolidBrush(color))
            {
                // Рисуем квадрат цвета
                e.Graphics.FillRectangle(brush, e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Height - 4, e.Bounds.Height - 4);
            }

            // Рисуем рамку вокруг квадрата
            e.Graphics.DrawRectangle(Pens.Black, e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Height - 4, e.Bounds.Height - 4);

            // Если элемент выделен, рисуем рамку вокруг всего элемента
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                e.Graphics.DrawRectangle(Pens.Blue, e.Bounds.X, e.Bounds.Y, e.Bounds.Width - 1, e.Bounds.Height - 1);
            }

            // Вызываем базовый метод для рисования текста
            e.DrawFocusRectangle();
        }

        //Удалить
        // Переопределите метод для отрисовки элементов
        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            // Вызываем базовый метод для рисования фона
            e.DrawBackground();

            // Если нет элементов, выходим
            if (e.Index < 0) return;

            // Получаем цвет из элементов
            Color color = (Color)this.Items[e.Index];

            // Создаем кисть с нужным цветом
            using (SolidBrush brush = new SolidBrush(color))
            {
                // Рисуем квадрат цвета
                e.Graphics.FillRectangle(brush, e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Height - 4, e.Bounds.Height - 4);
            }

            // Рисуем рамку вокруг квадрата
            e.Graphics.DrawRectangle(Pens.Black, e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Height - 4, e.Bounds.Height - 4);

            // Если элемент выделен, рисуем рамку вокруг всего элемента
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                e.Graphics.DrawRectangle(Pens.Blue, e.Bounds.X, e.Bounds.Y, e.Bounds.Width - 1, e.Bounds.Height - 1);
            }

            // Вызываем базовый метод для рисования текста
            base.OnDrawItem(e);
        }
    }
}
