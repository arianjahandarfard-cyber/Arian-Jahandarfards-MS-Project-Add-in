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
            this.btnLowerOutlineLevel = this.Factory.CreateRibbonButton();
            this.btnIncreaseOutlineLevel = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnDynamicStatusSheet = this.Factory.CreateRibbonButton();
            this.menuProjectLinker = this.Factory.CreateRibbonMenu();
            this.btnProjectLinkerExcel = this.Factory.CreateRibbonButton();
            this.btnProjectLinkerBoth = this.Factory.CreateRibbonButton();
            this.menuHighlighterOptions = this.Factory.CreateRibbonMenu();
            this.chkHighlighterOff = this.Factory.CreateRibbonCheckBox();
            this.separatorHighlighter = this.Factory.CreateRibbonSeparator();
            this.btnHighlightYellow = this.Factory.CreateRibbonButton();
            this.btnHighlightGreen = this.Factory.CreateRibbonButton();
            this.btnHighlightBlue = this.Factory.CreateRibbonButton();
            this.btnHighlightOrange = this.Factory.CreateRibbonButton();
            this.btnHighlightRed = this.Factory.CreateRibbonButton();
            this.btnHighlightPurple = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Update.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Update);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Arian Jahandarfard\'s Tools";
            this.tab1.Name = "tab1";
            // 
            // Update
            // 
            this.Update.Items.Add(this.btnCheckUpdates);
            this.Update.Name = "Update";
            // 
            // btnCheckUpdates
            // 
            this.btnCheckUpdates.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCheckUpdates.Image = ((System.Drawing.Image)(resources.GetObject("btnCheckUpdates.Image")));
            this.btnCheckUpdates.Label = "Check For Updates";
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
            this.btnCapture.Label = "Capture Snapshot";
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
            this.group2.Items.Add(this.btnLowerOutlineLevel);
            this.group2.Items.Add(this.btnIncreaseOutlineLevel);
            this.group2.Label = "Views";
            this.group2.Name = "group2";
            // 
            // btnGoToUID
            // 
            this.btnGoToUID.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGoToUID.Image = ((System.Drawing.Image)(resources.GetObject("btnGoToUID.Image")));
            this.btnGoToUID.Label = "UID Search";
            this.btnGoToUID.Name = "btnGoToUID";
            this.btnGoToUID.ScreenTip = "Navigate to task by UniqueID";
            this.btnGoToUID.ShowImage = true;
            this.btnGoToUID.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGoToUID_Click);
            // 
            // btnLowerOutlineLevel
            // 
            this.btnLowerOutlineLevel.Label = "Lower Outline Level";
            this.btnLowerOutlineLevel.Name = "btnLowerOutlineLevel";
            this.btnLowerOutlineLevel.ScreenTip = "Collapse the current schedule down one WBS outline level";
            this.btnLowerOutlineLevel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLowerOutlineLevel_Click);
            // 
            // btnIncreaseOutlineLevel
            // 
            this.btnIncreaseOutlineLevel.Label = "Increase Outline Level";
            this.btnIncreaseOutlineLevel.Name = "btnIncreaseOutlineLevel";
            this.btnIncreaseOutlineLevel.ScreenTip = "Expand the current schedule up one WBS outline level";
            this.btnIncreaseOutlineLevel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnIncreaseOutlineLevel_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnDynamicStatusSheet);
            this.group3.Items.Add(this.menuProjectLinker);
            this.group3.Label = "Status Tools";
            this.group3.Name = "group3";
            // 
            // btnDynamicStatusSheet
            // 
            this.btnDynamicStatusSheet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDynamicStatusSheet.Label = "Create Dynamic Status Sheet 5.6";
            this.btnDynamicStatusSheet.Name = "btnDynamicStatusSheet";
            this.btnDynamicStatusSheet.ScreenTip = "Create a dynamic status sheet from an open Excel workbook";
            this.btnDynamicStatusSheet.ShowImage = true;
            this.btnDynamicStatusSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDynamicStatusSheet_Click);
            // 
            // menuProjectLinker
            // 
            this.menuProjectLinker.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuProjectLinker.Image = ((System.Drawing.Image)(resources.GetObject("menuProjectLinker.Image")));
            this.menuProjectLinker.Items.Add(this.btnProjectLinkerExcel);
            this.menuProjectLinker.Items.Add(this.btnProjectLinkerBoth);
            this.menuProjectLinker.Items.Add(this.menuHighlighterOptions);
            this.menuProjectLinker.Label = "Project Linker";
            this.menuProjectLinker.Name = "menuProjectLinker";
            this.menuProjectLinker.ScreenTip = "Choose how the Project Linker should run";
            this.menuProjectLinker.ShowImage = true;
            // 
            // btnProjectLinkerExcel
            // 
            this.btnProjectLinkerExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnProjectLinkerExcel.Image")));
            this.btnProjectLinkerExcel.Label = "Excel";
            this.btnProjectLinkerExcel.Name = "btnProjectLinkerExcel";
            this.btnProjectLinkerExcel.ScreenTip = "Click an Excel row to jump to the matching Project task";
            this.btnProjectLinkerExcel.ShowImage = true;
            this.btnProjectLinkerExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProjectLinkerExcel_Click);
            // 
            // btnProjectLinkerBoth
            // 
            this.btnProjectLinkerBoth.Image = ((System.Drawing.Image)(resources.GetObject("btnProjectLinkerBoth.Image")));
            this.btnProjectLinkerBoth.Label = "Excel + Project";
            this.btnProjectLinkerBoth.Name = "btnProjectLinkerBoth";
            this.btnProjectLinkerBoth.ScreenTip = "Link Excel rows to Project and Project tasks back to Excel";
            this.btnProjectLinkerBoth.ShowImage = true;
            this.btnProjectLinkerBoth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProjectLinkerBoth_Click);
            // 
            // menuHighlighterOptions
            // 
            this.menuHighlighterOptions.Image = ((System.Drawing.Image)(resources.GetObject("menuHighlighterOptions.Image")));
            this.menuHighlighterOptions.Items.Add(this.chkHighlighterOff);
            this.menuHighlighterOptions.Items.Add(this.separatorHighlighter);
            this.menuHighlighterOptions.Items.Add(this.btnHighlightYellow);
            this.menuHighlighterOptions.Items.Add(this.btnHighlightGreen);
            this.menuHighlighterOptions.Items.Add(this.btnHighlightBlue);
            this.menuHighlighterOptions.Items.Add(this.btnHighlightOrange);
            this.menuHighlighterOptions.Items.Add(this.btnHighlightRed);
            this.menuHighlighterOptions.Items.Add(this.btnHighlightPurple);
            this.menuHighlighterOptions.Label = "Highlighter Options";
            this.menuHighlighterOptions.Name = "menuHighlighterOptions";
            this.menuHighlighterOptions.ScreenTip = "Choose whether matched tasks should stay highlighted and what color to use";
            this.menuHighlighterOptions.ShowImage = true;
            // 
            // chkHighlighterOff
            // 
            this.chkHighlighterOff.Checked = true;
            this.chkHighlighterOff.Label = "Off";
            this.chkHighlighterOff.Name = "chkHighlighterOff";
            this.chkHighlighterOff.ScreenTip = "Turn task highlighting off";
            this.chkHighlighterOff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkHighlighterOff_Click);
            // 
            // separatorHighlighter
            // 
            this.separatorHighlighter.Name = "separatorHighlighter";
            // 
            // btnHighlightYellow
            // 
            this.btnHighlightYellow.Label = "Yellow";
            this.btnHighlightYellow.Name = "btnHighlightYellow";
            this.btnHighlightYellow.ScreenTip = "Highlight matching tasks in yellow";
            this.btnHighlightYellow.ShowImage = true;
            this.btnHighlightYellow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightYellow_Click);
            // 
            // btnHighlightGreen
            // 
            this.btnHighlightGreen.Label = "Green";
            this.btnHighlightGreen.Name = "btnHighlightGreen";
            this.btnHighlightGreen.ScreenTip = "Highlight matching tasks in green";
            this.btnHighlightGreen.ShowImage = true;
            this.btnHighlightGreen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightGreen_Click);
            // 
            // btnHighlightBlue
            // 
            this.btnHighlightBlue.Label = "Blue";
            this.btnHighlightBlue.Name = "btnHighlightBlue";
            this.btnHighlightBlue.ScreenTip = "Highlight matching tasks in blue";
            this.btnHighlightBlue.ShowImage = true;
            this.btnHighlightBlue.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightBlue_Click);
            // 
            // btnHighlightOrange
            // 
            this.btnHighlightOrange.Label = "Orange";
            this.btnHighlightOrange.Name = "btnHighlightOrange";
            this.btnHighlightOrange.ScreenTip = "Highlight matching tasks in orange";
            this.btnHighlightOrange.ShowImage = true;
            this.btnHighlightOrange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightOrange_Click);
            // 
            // btnHighlightRed
            // 
            this.btnHighlightRed.Label = "Red";
            this.btnHighlightRed.Name = "btnHighlightRed";
            this.btnHighlightRed.ScreenTip = "Highlight matching tasks in red";
            this.btnHighlightRed.ShowImage = true;
            this.btnHighlightRed.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightRed_Click);
            // 
            // btnHighlightPurple
            // 
            this.btnHighlightPurple.Label = "Purple";
            this.btnHighlightPurple.Name = "btnHighlightPurple";
            this.btnHighlightPurple.ScreenTip = "Highlight matching tasks in purple";
            this.btnHighlightPurple.ShowImage = true;
            this.btnHighlightPurple.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightPurple_Click);
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
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLowerOutlineLevel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnIncreaseOutlineLevel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Update;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckUpdates;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDynamicStatusSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuProjectLinker;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnProjectLinkerExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnProjectLinkerBoth;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuHighlighterOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkHighlighterOff;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separatorHighlighter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightYellow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightGreen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightBlue;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightOrange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightRed;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightPurple;
    }

    partial class ThisRibbonCollection
    {
        internal AJRibbon AJRibbon
        {
            get { return this.GetRibbon<AJRibbon>(); }
        }
    }
}
