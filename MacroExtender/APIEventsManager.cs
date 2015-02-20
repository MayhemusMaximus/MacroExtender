using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
//using MacroExtender.Formatting;

namespace MacroExtender
{
    public class APIEventsManager : MacroExtenderRibbon
    {

        private const int DISP_E_BADINDEX = unchecked((int)0x800200B);

        private static Excel.Range lastSelection = null;
        private static double lastSelectionColor = 0;

        /// <summary>
        /// This method is called by the WorkbookBeforeCloseEvent
        /// </summary>
        /// <param name="wb">The workbook that is closing.</param>
        /// <param name="Cancel"></param>
        public void excelEvents_WorkbookBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {

            // THIS EVENT HANDLER IS BEING USED TO CHANGE THE BUTTON STATES OF THE MACRO
            // EXTENDER TO DEFAULT IF THE LAST WORKBOOK IS CLOSED.

            // CREATES AN INSTANCE OF THE MACROEXTENDER RIBBON
            MacroExtenderRibbon thisInstance = new MacroExtenderRibbon();

            // XLSB IS THE FILE EXTENSION OF "PERSONAL FILES" CREATED IN EXCEL BY
            // VBA/MACRO USERS. THEY ARE OPEN WHILE EXCEL IS RUNNING AND THEREFORE
            // SHOW UP IN THE WORKBOOK COUNT. IN ORDER TO GET AN ACCURE COUNT OF
            // WORKBOOKS THE USER IS WORKING IN THE FOREACH COUNTS THE XLSB FILES.
            int openXLSBCount = 0;
            //////foreach (Excel.Workbook openWB in ExcelObj.Workbooks)
            //////{
            //////    if (stringCompare("xlsb", StringExt.Right(openWB.Name, 4))) // VBA PERSONAL FILES
            //////    {
            //////        openXLSBCount++;
            //////    }
            //////}

            // AFTER GETTING THE XLSB COUNT SUBTRACT IT FROM THE WORKBOOK COUNT.
            // IF THE RESULT IS ONE THEN THE USER IS CLOSING THE LAST USER WORKBOOK
            // THAT IS OPEN; DISABLE THE BUTTONS THAT WILL NOT WORK IN THIS STATE.
            if ((ExcelObj.Workbooks.Count - openXLSBCount) == 1)
            {
                thisInstance.ScopeSelectionCBoxEnabledState(false);
                thisInstance.MacroSelectionCBoxEnabledState(false);

                thisInstance.InsertMacrosSheetButtonEnabledState(false);
                thisInstance.RefreshMacrosButtonEnabledState(false);
                thisInstance.ExecuteMacroButtonEnabledState(false);
                thisInstance.OptionsButtonEnabledState(true);
            }
        }

        /// <summary>This function returns true if the parameters match.
        /// It is used when case sensitivity is not necessary.</summary>
        /// <param name="string1">First case-insensitive string to be compared.</param>
        /// <param name="string2">Second case-insensitive string to be compared.</param>
        private bool stringCompare(string string1, string string2)
        {
            bool comparison = false;

            string1 = string1.ToUpper();
            string2 = string2.ToUpper();

            if (string1 == string2)
                comparison = true;

            return comparison;
        }

        /// <summary>
        /// This event handler is called with the SheetActivateEvent.
        /// </summary>
        /// <param name="Sh">The sheet that is activating.</param>
        public void excelEvents_SheetActivate(object Sh)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Sh;
            MacroExtenderRibbon thisInstance = new MacroExtenderRibbon();

            // SINCE DELETING A SHEET WILL FIRE THIS EVENT
            // USE TO DETERMINE IF IT WAS FIRED BECAUSE 
            // THE MACRO SHEET HAS BEEN DELETED.
            if (SheetExists("Macros"))
            {
                // THE SHEET EXISTS SO UPDATE THE CBOX
                // TO MATCH THE ACTIVE SHEET.
                ActivatedSheet = sheet;
                thisInstance.setMacroSelectionComboBox();
            }
            else
            {
                // THE "Macros" SHEET DOES NOT EXISTS SO ENSURE THAT
                // RIBBON LOSES FUNCTIONALITY WHERE IT IS USELESS.
                thisInstance.ScopeSelectionCBoxEnabledState(false);
                thisInstance.MacroSelectionCBoxEnabledState(false);

                thisInstance.InsertMacrosSheetButtonEnabledState(false);
                thisInstance.RefreshMacrosButtonEnabledState(false);
                thisInstance.ExecuteMacroButtonEnabledState(false);

                thisInstance.OptionsButtonEnabledState(true);
            }
        }

        public void excelEvents_WorkbookOpen(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            // THIS HANDLER METHOD CHECKS OPENNING WORKBOOKS FOR THE "Macros" SHEET
            // IF IT IS FOUND IT CREATES A SHEET CHANGE HANDLER TO ENABLE THE "Refresh"
            // BUTTON.

            // GET A HANDLE ON THE RIBBON CLASS
            MacroExtenderRibbon thisInstance = new MacroExtenderRibbon();
            ActivatedSheet = wb.ActiveSheet;

            // IF THE "Macros" SHEET EXISTS IT WILL BE HELD HERE.
            Excel.Worksheet sheet = null;


            try
            {
                // TRY TO ASSIGN sheet TO THE "Macros" SHEET (WON'T WORK IF IT DOESN'T EXIST).
                sheet = (Excel.Worksheet)wb.Sheets["Macros"];
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // THE EXCEOPTION IS ABOUT SOMETHING OTHER THAN THE SHEET NOT EXISTING; RETHROW THE EXCEPTION
                if (ex.ErrorCode != DISP_E_BADINDEX)
                {
                    throw;
                }
                else
                {
                    // THE SHEET DOESN'T EXIST.
                }
            }

            if (sheet != null)
            {
                sheet.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(excelEvents_CellsChange);
                sheet.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(ActiveSheet_SelectionChange);
            }
        }

        public void ActiveSheet_SelectionChange(Excel.Range target)
        {
            //////InteractiveFormatting interForm = new InteractiveFormatting();
            //////lastSelectionColor = interForm.SelectionColor(target, lastSelection, lastSelectionColor);
            //////lastSelection = target;
        }

        public void excelEvents_WorkbookNewSheet(Excel.Workbook wb, object sh)
        {
            Excel.Worksheet worksheet = (Excel.Worksheet)sh;
            if (worksheet != null)
            {
                MessageBox.Show(wb.Name, worksheet.Name);
            }
            MessageBox.Show("A new sheet has been added: " + sh);
        }

        public void excelEvents_WorkbookActivate(Excel.Workbook Wb)
        {
            MacroExtenderRibbon thisInstance = new MacroExtenderRibbon();

            UpdateExcelHandles();

            if (SheetExists("Macros"))
            {

                thisInstance.ScopeSelectionCBoxEnabledState(true);
                thisInstance.MacroSelectionCBoxEnabledState(true);

                thisInstance.InsertMacrosSheetButtonEnabledState(false);
                thisInstance.RefreshMacrosButtonEnabledState(false);
                thisInstance.ExecuteMacroButtonEnabledState(false);
                thisInstance.OptionsButtonEnabledState(true);


                thisInstance.UpdateCBoxes();
                thisInstance.BuildMacroList();
            }
            else
            {

                thisInstance.InsertMacrosSheetButtonEnabledState(true);
                thisInstance.ScopeSelectionCBoxEnabledState(false);

                thisInstance.MacroSelectionCBoxEnabledState(false);
                thisInstance.RefreshMacrosButtonEnabledState(false);
                thisInstance.OptionsButtonEnabledState(false);
            }
        }

        public void excelEvents_CellsChange(Excel.Range target)
        {
            if (target.Worksheet.Name.Equals("Macros"))
            {
                // CODE FOR Macros SHEET INTERACTION GOES HERE.
                MacroExtenderRibbon thisInstance = new MacroExtenderRibbon();
                RefreshMacrosButtonEnabledState(true);

                //////InteractiveFormatting interform = new InteractiveFormatting();
                //////interform.delegateCell(target);
            }
        }
    }
}
