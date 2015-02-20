namespace MacroExtender
{
    partial class MacroExtenderRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MacroExtenderRibbon()
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
            this.MacroExtenderGroup = this.Factory.CreateRibbonGroup();
            this.ExecuteMacroButton = this.Factory.CreateRibbonButton();
            this.InsertMacrosSheetButton = this.Factory.CreateRibbonButton();
            this.OptionsButton = this.Factory.CreateRibbonButton();
            this.RefreshMacrosButton = this.Factory.CreateRibbonButton();
            this.ScopeSelectionComboBox = this.Factory.CreateRibbonComboBox();
            this.MacroSelectionComboBox = this.Factory.CreateRibbonComboBox();
            this.tab1.SuspendLayout();
            this.MacroExtenderGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.MacroExtenderGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // MacroExtenderGroup
            // 
            this.MacroExtenderGroup.Items.Add(this.ScopeSelectionComboBox);
            this.MacroExtenderGroup.Items.Add(this.MacroSelectionComboBox);
            this.MacroExtenderGroup.Items.Add(this.ExecuteMacroButton);
            this.MacroExtenderGroup.Items.Add(this.OptionsButton);
            this.MacroExtenderGroup.Items.Add(this.InsertMacrosSheetButton);
            this.MacroExtenderGroup.Items.Add(this.RefreshMacrosButton);
            this.MacroExtenderGroup.Label = "Macro Extender";
            this.MacroExtenderGroup.Name = "MacroExtenderGroup";
            // 
            // ExecuteMacroButton
            // 
            this.ExecuteMacroButton.Label = "             Execute Macro     ";
            this.ExecuteMacroButton.Name = "ExecuteMacroButton";
            // 
            // InsertMacrosSheetButton
            // 
            this.InsertMacrosSheetButton.Label = "Insert Macros Sheet";
            this.InsertMacrosSheetButton.Name = "InsertMacrosSheetButton";
            // 
            // OptionsButton
            // 
            this.OptionsButton.Label = "Options";
            this.OptionsButton.Name = "OptionsButton";
            // 
            // RefreshMacrosButton
            // 
            this.RefreshMacrosButton.Label = "Refresh Macros";
            this.RefreshMacrosButton.Name = "RefreshMacrosButton";
            // 
            // ScopeSelectionComboBox
            // 
            this.ScopeSelectionComboBox.Label = "Scope:";
            this.ScopeSelectionComboBox.Name = "ScopeSelectionComboBox";
            // 
            // MacroSelectionComboBox
            // 
            this.MacroSelectionComboBox.Label = "Macro:";
            this.MacroSelectionComboBox.Name = "MacroSelectionComboBox";
            // 
            // MacroExtenderRibbon
            // 
            this.Name = "MacroExtenderRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.MacroExtenderGroup.ResumeLayout(false);
            this.MacroExtenderGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup MacroExtenderGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExecuteMacroButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox ScopeSelectionComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox MacroSelectionComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OptionsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertMacrosSheetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RefreshMacrosButton;
    }

    partial class ThisRibbonCollection
    {
        internal MacroExtenderRibbon Ribbon1
        {
            get { return this.GetRibbon<MacroExtenderRibbon>(); }
        }
    }
}
