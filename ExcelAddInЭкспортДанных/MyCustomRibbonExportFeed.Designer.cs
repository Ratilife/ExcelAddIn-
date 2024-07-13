namespace ExcelAddInЭкспортДанных
{
    partial class MyCustomRibbonExportFeed : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyCustomRibbonExportFeed()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tabRatiTools = this.Factory.CreateRibbonTab();
            this.grExportFeed = this.Factory.CreateRibbonGroup();
            this.butExportXLSXtoCSV = this.Factory.CreateRibbonButton();
            this.butExportXLSXtoPDF = this.Factory.CreateRibbonButton();
            this.butExportXLSXtoJSON = this.Factory.CreateRibbonButton();
            this.butExportXLSXtoTXT = this.Factory.CreateRibbonButton();
            this.butExportXLSXtoXLS = this.Factory.CreateRibbonButton();
            this.butExportXLSXtoXLSM = this.Factory.CreateRibbonButton();
            this.butExportXLSXtoXML = this.Factory.CreateRibbonButton();
            this.butExportXLSXtoHTML = this.Factory.CreateRibbonButton();
            this.tabRatiTools.SuspendLayout();
            this.grExportFeed.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabRatiTools
            // 
            this.tabRatiTools.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabRatiTools.Groups.Add(this.grExportFeed);
            this.tabRatiTools.Label = "RatiTools";
            this.tabRatiTools.Name = "tabRatiTools";
            // 
            // grExportFeed
            // 
            this.grExportFeed.DialogLauncher = ribbonDialogLauncherImpl1;
            this.grExportFeed.Items.Add(this.butExportXLSXtoCSV);
            this.grExportFeed.Items.Add(this.butExportXLSXtoPDF);
            this.grExportFeed.Items.Add(this.butExportXLSXtoJSON);
            this.grExportFeed.Items.Add(this.butExportXLSXtoTXT);
            this.grExportFeed.Items.Add(this.butExportXLSXtoXLS);
            this.grExportFeed.Items.Add(this.butExportXLSXtoXLSM);
            this.grExportFeed.Items.Add(this.butExportXLSXtoXML);
            this.grExportFeed.Items.Add(this.butExportXLSXtoHTML);
            this.grExportFeed.Label = "Экспорт данных";
            this.grExportFeed.Name = "grExportFeed";
            // 
            // butExportXLSXtoCSV
            // 
            this.butExportXLSXtoCSV.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.butExportXLSXtoCSV.Image = global::ExcelAddInЭкспортДанных.Properties.Resources.csv_filetype_icon_177543;
            this.butExportXLSXtoCSV.Label = "Экспорт в CSV";
            this.butExportXLSXtoCSV.Name = "butExportXLSXtoCSV";
            this.butExportXLSXtoCSV.ShowImage = true;
            this.butExportXLSXtoCSV.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butExportXLSXtoCSV_Click);
            // 
            // butExportXLSXtoPDF
            // 
            this.butExportXLSXtoPDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.butExportXLSXtoPDF.Image = global::ExcelAddInЭкспортДанных.Properties.Resources.pdf_filetype_icon_177525;
            this.butExportXLSXtoPDF.Label = "Экспорт в PDF";
            this.butExportXLSXtoPDF.Name = "butExportXLSXtoPDF";
            this.butExportXLSXtoPDF.ShowImage = true;
            this.butExportXLSXtoPDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butExportXLSXtoPDF_Click);
            // 
            // butExportXLSXtoJSON
            // 
            this.butExportXLSXtoJSON.Image = global::ExcelAddInЭкспортДанных.Properties.Resources.json_filetype_icon_177531;
            this.butExportXLSXtoJSON.Label = "Экспорт в JSON";
            this.butExportXLSXtoJSON.Name = "butExportXLSXtoJSON";
            this.butExportXLSXtoJSON.ShowImage = true;
            this.butExportXLSXtoJSON.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butExportXLSXtoJSON_Click);
            // 
            // butExportXLSXtoTXT
            // 
            this.butExportXLSXtoTXT.Image = global::ExcelAddInЭкспортДанных.Properties.Resources.txt_filetype_icon_177515;
            this.butExportXLSXtoTXT.Label = "Экспорт в TXT";
            this.butExportXLSXtoTXT.Name = "butExportXLSXtoTXT";
            this.butExportXLSXtoTXT.ShowImage = true;
            this.butExportXLSXtoTXT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butExportXLSXtoTXT_Click);
            // 
            // butExportXLSXtoXLS
            // 
            this.butExportXLSXtoXLS.Image = global::ExcelAddInЭкспортДанных.Properties.Resources.xls_filetype_icon_177510;
            this.butExportXLSXtoXLS.Label = "Экспорт в XLS";
            this.butExportXLSXtoXLS.Name = "butExportXLSXtoXLS";
            this.butExportXLSXtoXLS.ShowImage = true;
            this.butExportXLSXtoXLS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butExportXLSXtoXLS_Click);
            // 
            // butExportXLSXtoXLSM
            // 
            this.butExportXLSXtoXLSM.Image = global::ExcelAddInЭкспортДанных.Properties.Resources.free_icon_file_14421731;
            this.butExportXLSXtoXLSM.Label = "Экспорт в XLSM";
            this.butExportXLSXtoXLSM.Name = "butExportXLSXtoXLSM";
            this.butExportXLSXtoXLSM.ShowImage = true;
            this.butExportXLSXtoXLSM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butExportXLSXtoXLSM_Click);
            // 
            // butExportXLSXtoXML
            // 
            this.butExportXLSXtoXML.Image = global::ExcelAddInЭкспортДанных.Properties.Resources.xml_filetype_icon_177509;
            this.butExportXLSXtoXML.Label = "Экспорт в XML";
            this.butExportXLSXtoXML.Name = "butExportXLSXtoXML";
            this.butExportXLSXtoXML.ShowImage = true;
            this.butExportXLSXtoXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butExportXLSXtoXML_Click);
            // 
            // butExportXLSXtoHTML
            // 
            this.butExportXLSXtoHTML.Image = global::ExcelAddInЭкспортДанных.Properties.Resources.html_filetype_icon_177535;
            this.butExportXLSXtoHTML.Label = "Экспорт в HTML";
            this.butExportXLSXtoHTML.Name = "butExportXLSXtoHTML";
            this.butExportXLSXtoHTML.ShowImage = true;
            this.butExportXLSXtoHTML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butExportXLSXtoHTML_Click);
            // 
            // MyCustomRibbonExportFeed
            // 
            this.Name = "MyCustomRibbonExportFeed";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabRatiTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyCustomRibbonExportFeed_Load);
            this.tabRatiTools.ResumeLayout(false);
            this.tabRatiTools.PerformLayout();
            this.grExportFeed.ResumeLayout(false);
            this.grExportFeed.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabRatiTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grExportFeed;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butExportXLSXtoCSV;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butExportXLSXtoPDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butExportXLSXtoTXT;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butExportXLSXtoJSON;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butExportXLSXtoXLS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butExportXLSXtoXLSM;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butExportXLSXtoXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butExportXLSXtoHTML;
    }

    partial class ThisRibbonCollection
    {
        internal MyCustomRibbonExportFeed MyCustomRibbonExportFeed
        {
            get { return this.GetRibbon<MyCustomRibbonExportFeed>(); }
        }
    }
}
