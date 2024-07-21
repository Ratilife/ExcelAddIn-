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
        }
    }
}
