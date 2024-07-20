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
    public partial class TableСreationForm : UserControl
    {
        public string NameTable { get; private set; }
        public string CellAddress { get; private set; }
        public int KolСolumns { get; private set; }
        public int KolRows { get; private set; }
        public bool ActivSeheet { get; private set; }
        public bool NewSheet { get; private set; }

        public TableСreationForm()
        {
            InitializeComponent();
        }

        

        private void btCreate_Click(object sender, EventArgs e)
        {
            if (tbNameTable.Text.Length > 0)
            {
                NameTable = tbNameTable.Text;
            }
            else 
            {
                MessageBox.Show("Присвойте имя таблице.");
            }

            if (tbCellAddress.Text.Length > 0)
            {
                CellAddress = tbCellAddress.Text;
            }
            if (tbСolumns.Text.Length > 0)
            {
                KolСolumns = Convert.ToInt16(tbСolumns.Text);
                
            }
            if (tbRows.Text.Length > 0)
            {
                KolRows = Convert.ToInt16(tbRows.Text);
            }

            if(rbActivSeheet.Checked == true) 
            {
                ActivSeheet = true;
                NewSheet = false;
            }
            else 
            {
                NewSheet = true;
                ActivSeheet = false;
            }

            WorkingWithTables tables = new WorkingWithTables();

            tables.CreateTable(CellAddress, KolСolumns, KolRows,ActivSeheet,NewSheet,NameTable);


        }

        private void btClose_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void TableСreationForm_Load(object sender, EventArgs e)
        {
            if (tbCellAddress.Text.Length == 0)
            {
                tbCellAddress.Text = "A1";
            }
        }
    }
}
