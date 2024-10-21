using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddInЭкспортДанных.forms
{
    public partial class ExportCsvToXlsx : Form
    {
        CommonMethods cm = new CommonMethods();
        public ExportCsvToXlsx()
        {
            InitializeComponent();
            cbActiveWorkbook.Enabled = false;
        }

        private void butOK_Click(object sender, EventArgs e)
        {
            ExportData ed = new ExportData();
            ed.ExportCsvToXlsx(tbCsvFilePath.Text, tbXlsxFilePath.Text);
            // Устанавливаем результат диалога как OK и закрываем форму
            DialogResult = DialogResult.OK;
            Close();
        }

        private void butCsvFilePath_Click(object sender, EventArgs e)
        {
            String filePath = cm.OpenCsvFile();
            tbCsvFilePath.Text = filePath;
        }

        private void butXlsxFilePath_Click(object sender, EventArgs e)
        {
            String filePath = cm.dialogFolder();
            tbXlsxFilePath.Text = filePath;
        }

        private void cbActiveWorkbook_CheckedChanged(object sender, EventArgs e)
        {
            if (cbActiveWorkbook.Checked == true) 
            {
                tbXlsxFilePath.Enabled = true;
                butXlsxFilePath.Enabled=true;
            }
        }
    }
}
