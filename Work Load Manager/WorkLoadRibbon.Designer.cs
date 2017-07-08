namespace WorkLoadManager
{
    partial class WorkLoadRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WorkLoadRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabWorkLoad = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnConnect = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.tabWorkLoad.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabWorkLoad
            // 
            this.tabWorkLoad.Groups.Add(this.group1);
            this.tabWorkLoad.Groups.Add(this.group2);
            this.tabWorkLoad.Label = "Work Load Manager";
            this.tabWorkLoad.Name = "tabWorkLoad";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnConnect);
            this.group1.Label = "Rally";
            this.group1.Name = "group1";
            // 
            // btnConnect
            // 
            this.btnConnect.ImageName = "imgRally";
            this.btnConnect.Label = "Connect";
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.OfficeImageId = "ChartLines";
            this.btnConnect.ScreenTip = "Connect to Rally";
            this.btnConnect.ShowImage = true;
            this.btnConnect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RallyConnect);
            // 
            // group2
            // 
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // WorkLoadRibbon
            // 
            this.Name = "WorkLoadRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabWorkLoad);
            this.tabWorkLoad.ResumeLayout(false);
            this.tabWorkLoad.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabWorkLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConnect;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
    }

    partial class ThisRibbonCollection
    {
        internal WorkLoadRibbon WorkLoadRibbon
        {
            get { return this.GetRibbon<WorkLoadRibbon>(); }
        }
    }
}
