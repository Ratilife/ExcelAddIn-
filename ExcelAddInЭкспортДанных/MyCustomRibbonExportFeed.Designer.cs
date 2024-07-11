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
            this.butExportXLSXtoTXT = this.Factory.CreateRibbonButton();
            this.butExportXLSXtoCSV = this.Factory.CreateRibbonButton();
            this.butExportXLSXtoPDF = this.Factory.CreateRibbonButton();
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
            this.grExportFeed.Items.Add(this.butExportXLSXtoTXT);
            this.grExportFeed.Items.Add(this.butExportXLSXtoCSV);
            this.grExportFeed.Items.Add(this.butExportXLSXtoPDF);
            this.grExportFeed.Label = "Экспорт данных";
            this.grExportFeed.Name = "grExportFeed";
            // 
            // butExportXLSXtoTXT
            // 
            this.butExportXLSXtoTXT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.butExportXLSXtoTXT.Image = global::ExcelAddInЭкспортДанных.Properties.Resources.txt_filetype_icon_177515;
            this.butExportXLSXtoTXT.Label = "Экспорт в TXT";
            this.butExportXLSXtoTXT.Name = "butExportXLSXtoTXT";
            this.butExportXLSXtoTXT.ShowImage = true;
            this.butExportXLSXtoTXT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butExportXLSXtoTXT_Click);
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
    }

    partial class ThisRibbonCollection
    {
        internal MyCustomRibbonExportFeed MyCustomRibbonExportFeed
        {
            get { return this.GetRibbon<MyCustomRibbonExportFeed>(); }
        }
    }
}
