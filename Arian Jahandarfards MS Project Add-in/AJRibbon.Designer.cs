namespace Arian_Jahandarfards_MS_Project_Add_in
{
    partial class AJRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public AJRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AJRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Update = this.Factory.CreateRibbonGroup();
            this.btnCheckUpdates = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnCapture = this.Factory.CreateRibbonButton();
            this.btnReset = this.Factory.CreateRibbonButton();
            this.btnRun = this.Factory.CreateRibbonButton();
            this.btnStartAuto = this.Factory.CreateRibbonButton();
            this.btnStopAuto = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnGoToUID = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Update.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Update);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Arian Jahandarfard\'s Tools";
            this.tab1.Name = "tab1";
            // 
            // Update
            // 
            this.Update.Items.Add(this.btnCheckUpdates);
            this.Update.Label = "Update";
            this.Update.Name = "Update";
            // 
            // btnCheckUpdates
            // 
            this.btnCheckUpdates.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCheckUpdates.Image = ((System.Drawing.Image)(resources.GetObject("btnCheckUpdates.Image")));
            this.btnCheckUpdates.Label = "Check For Update";
            this.btnCheckUpdates.Name = "btnCheckUpdates";
            this.btnCheckUpdates.ShowImage = true;
            this.btnCheckUpdates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCheckUpdates_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnCapture);
            this.group1.Items.Add(this.btnReset);
            this.group1.Items.Add(this.btnRun);
            this.group1.Items.Add(this.btnStartAuto);
            this.group1.Items.Add(this.btnStopAuto);
            this.group1.Items.Add(this.btnSettings);
            this.group1.Label = "Milestone Tracker";
            this.group1.Name = "group1";
            // 
            // btnCapture
            // 
            this.btnCapture.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCapture.Image = ((System.Drawing.Image)(resources.GetObject("btnCapture.Image")));
            this.btnCapture.Label = "   Capture Snapshot";
            this.btnCapture.Name = "btnCapture";
            this.btnCapture.ShowImage = true;
            this.btnCapture.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCapture_Click);
            // 
            // btnReset
            // 
            this.btnReset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReset.Image = ((System.Drawing.Image)(resources.GetObject("btnReset.Image")));
            this.btnReset.Label = "Reset Snapshot   ";
            this.btnReset.Name = "btnReset";
            this.btnReset.ShowImage = true;
            this.btnReset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReset_Click);
            // 
            // btnRun
            // 
            this.btnRun.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRun.Image = ((System.Drawing.Image)(resources.GetObject("btnRun.Image")));
            this.btnRun.Label = "Run Tracker";
            this.btnRun.Name = "btnRun";
            this.btnRun.ShowImage = true;
            this.btnRun.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRun_Click);
            // 
            // btnStartAuto
            // 
            this.btnStartAuto.Image = ((System.Drawing.Image)(resources.GetObject("btnStartAuto.Image")));
            this.btnStartAuto.Label = "Start Auto Track";
            this.btnStartAuto.Name = "btnStartAuto";
            this.btnStartAuto.ShowImage = true;
            this.btnStartAuto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStartAuto_Click);
            // 
            // btnStopAuto
            // 
            this.btnStopAuto.Image = ((System.Drawing.Image)(resources.GetObject("btnStopAuto.Image")));
            this.btnStopAuto.Label = "Stop Auto Track";
            this.btnStopAuto.Name = "btnStopAuto";
            this.btnStopAuto.ShowImage = true;
            this.btnStopAuto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStopAuto_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.Image = ((System.Drawing.Image)(resources.GetObject("btnSettings.Image")));
            this.btnSettings.Label = "Settings";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.ShowImage = true;
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettings_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnGoToUID);
            this.group2.Label = "-----";
            this.group2.Name = "group2";
            // 
            // btnGoToUID
            // 
            this.btnGoToUID.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGoToUID.Image = ((System.Drawing.Image)(resources.GetObject("btnGoToUID.Image")));
            this.btnGoToUID.Label = "Search UID";
            this.btnGoToUID.Name = "btnGoToUID";
            this.btnGoToUID.ScreenTip = "Navigate to task by UniqueID";
            this.btnGoToUID.ShowImage = true;
            this.btnGoToUID.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGoToUID_Click);
            // 
            // AJRibbon
            // 
            this.Name = "AJRibbon";
            this.RibbonType = "Microsoft.Project.Project";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AJRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Update.ResumeLayout(false);
            this.Update.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCapture;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReset;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRun;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStartAuto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStopAuto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGoToUID;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Update;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckUpdates;
    }

    partial class ThisRibbonCollection
    {
        internal AJRibbon AJRibbon
        {
            get { return this.GetRibbon<AJRibbon>(); }
        }
    }
}