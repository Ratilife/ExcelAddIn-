using System;

using System.Windows.Forms;

using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddInЭкспортДанных
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }
        // Реализация метода InternalStartup
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

    }
}
