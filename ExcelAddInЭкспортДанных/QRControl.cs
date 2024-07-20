using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddInЭкспортДанных
{
    public partial class QRControl : UserControl
    {
        private ColorComboBox colorComboBox;

        public QRControl()
        {
            InitializeComponent();

            setColorComboBox();
            
        }

       void setColorComboBox()
       {
            colorComboBox = new ColorComboBox();
            colorComboBox.InitializeColorComboBox(cbColour);
            colorComboBox.InitializeColorComboBox(cbBackground);

            // Устанавливаем значения по умолчанию
            cbColour.SelectedItem = Color.Black;
            cbBackground.SelectedItem = Color.White;
        }
    }
}
