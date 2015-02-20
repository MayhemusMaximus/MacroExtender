using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Excel = Microsoft.Office.Interop.Excel;

using System.Windows.Forms;

namespace MacroExtender
{

    public partial class MacroExtenderRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        #region Ribbon Events Region

        private void MacroExtenderRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Preset objects to not be available to user.
            MacroSelectionComboBox.Enabled = false;
            ScopeSelectionComboBox.Enabled = false;
            RefreshMacrosButton.Enabled = false;
            InsertMacrosSheetButton.Enabled = false;
            ExecuteMacroButton.Enabled = false;
        }

        private void InsertMacrosSheetButton_Click(object sender, RibbonControlEventArgs e)
        {
            //Template template = new Template();
            //template.InsertMacrosSheet();
            //OptionsButtonEnabledState(true);

            //APIEventsManager eventsManager = new APIEventsManager();

            //Sheet.Change += new Excel.DocEvents_ChangeEventHandler(eventsManager.excelEvents_CellsChange);
        }

        private void WorksheetSelectionComboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            //setMacroSelectionComboBox();
        }

        private void MacroSelectionComboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            ExecuteMacroButtonEnabledState(true);
        }

        #endregion

        private void ExecuteMacroButtonEnabledState(bool enabled)
        {
            //            Globals.Ribbons.MacroExtenderRibbon.ExecuteMacroButton.Enabled = enabled;
        }
    }
}
