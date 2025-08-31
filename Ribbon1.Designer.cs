namespace SmartMailAssistant1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.SummaryButton = this.Factory.CreateRibbonButton();
            this.TranslateButton = this.Factory.CreateRibbonButton();
            this.SuggestReplyButton = this.Factory.CreateRibbonButton();
            this.ClassifyButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "SmartMailAssistant";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.SummaryButton);
            this.group1.Items.Add(this.TranslateButton);
            this.group1.Items.Add(this.SuggestReplyButton);
            this.group1.Items.Add(this.ClassifyButton);
            this.group1.Label = "options";
            this.group1.Name = "group1";
            // 
            // SummaryButton
            // 
            this.SummaryButton.Label = "Summarize";
            this.SummaryButton.Name = "SummaryButton";
            this.SummaryButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // TranslateButton
            // 
            this.TranslateButton.Label = "Translate";
            this.TranslateButton.Name = "TranslateButton";
            this.TranslateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // SuggestReplyButton
            // 
            this.SuggestReplyButton.Label = "Suggest reply";
            this.SuggestReplyButton.Name = "SuggestReplyButton";
            this.SuggestReplyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // ClassifyButton
            // 
            this.ClassifyButton.Label = "Classify Emails";
            this.ClassifyButton.Name = "ClassifyButton";
            this.ClassifyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SummaryButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SuggestReplyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TranslateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ClassifyButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}