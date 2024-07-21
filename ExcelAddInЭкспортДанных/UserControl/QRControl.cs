using Microsoft.Office.Interop.Excel;
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
        private string filePath { get;  set; }
        public string QRToDate { get; private set; }
        public string QRToDateMany { get; private set; }
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

        private void btPathFolder_Click(object sender, EventArgs e)
        {
           
            CommonMethods cm = new CommonMethods();
            filePath = cm.dialogFolder();
            txtPathFolder.Text = filePath;

        }

        private void QRControl_Load(object sender, EventArgs e)
        {
            txtPathFolder.Visible = false;
            btPathFolder.Visible=false;
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
            if(rbOne.Checked)
            {
                QRToDate = txtQRcode.Text;
            }
            if(rbMany.Checked)
            {

            }
            QRcode qr = new QRcode();
            qr.CreateQRCodePicture(QRToDate, filePath);
            
        }

        private void txtQRcode_KeyDown(object sender, KeyEventArgs e)
        {
            txtQRcode.Enabled = true;
        }
    }
}
