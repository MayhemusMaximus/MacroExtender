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

        public struct MacrosStruct
        {
            private readonly int row;
            private readonly string macroName;
            private readonly string scope;
            private readonly string worksheet;

            public MacrosStruct(int row,
                            string macroName,
                            string scope,
                            string worksheet)
            {
                this.row = row;
                this.macroName = macroName;
                this.scope = scope;
                this.worksheet = worksheet;
            }

            public int Row { get { return row; } }
            public string MacroName { get { return macroName; } }
            public string Scope { get { return scope; } }
            public string Worksheet { get { return worksheet; } }
        }

        #region Fields and Properties Region

        //ExcelBase excelBase = new ExcelBase();

        private static Excel.Application excelObj;

        private static Excel.Workbook wb;

        private static Excel.Sheets sheets;

        private static Excel.Worksheet sheet;

        private static Excel.Worksheet activatedSheet;

        public static List<MacrosStruct> MacrosList = new List<MacrosStruct>();

        public const string ScopeListColumn = "A";

        public const string MethodListColumn = "B";

        public const string ScopeColumn = "D";

        public const string WorksheetColumn = "E";

        public const string MacroNameColumn = "F";

        public const string MethodColumn = "G";

        public const string Arg1Column = "H";

        public const string Arg2Column = "I";

        public const string Arg3Column = "J";

        public const string Arg4Column = "K";

        public const string Arg5Column = "L";

        public const string Arg6Column = "M";

        public const string Arg7Column = "N";

        public const string Arg8Column = "O";

        public const string Arg9Column = "P";

        public const string Arg10Column = "Q";

        // USED TO STOP BUILDING THE MacrosList
        private const int maxSearchRow = 1000;

        // USER VARIABLES
        public List<String> userStringList = new List<string>();
        public List<Int32> userIntegerList = new List<Int32>();

        public Dictionary<int, int> userIntIndexRef = new Dictionary<int, int>();
        public Dictionary<int, int> userStrIndexRef = new Dictionary<int, int>();

        public Excel.Application ExcelObj
        {
            get { return excelObj; }
            set { excelObj = value; }
        }

        public Excel.Workbook WB
        {
            get { return wb; }
            set { wb = value; }
        }

        public Excel.Sheets Sheets
        {
            get { return sheets; }
            set { sheets = value; }
        }

        public Excel.Worksheet Sheet
        {
            get { return sheet; }
            set { sheet = value; }
        }

        public Excel.Worksheet ActivatedSheet
        {
            get { return activatedSheet; }
            set { activatedSheet = value; }
        }

        public int MaxSearchRow
        {
            get { return maxSearchRow; }
        }

        #endregion

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

        #region Excel Behavior Region

        public void UpdateExcelHandles()
        {
            int iSection = 0, iTries = 0;
        tryAgain:
            try
            {

                // ATTEMPT TO USE GetObject TO REFERENCE THE RUNNING OFFICE APPLICATION

                // ASSIGN THE ACTIVE OBJECT (Excel.Application) TO THE excelObj OBJECT
                iSection = 1;
                ExcelObj = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                iSection = 0;

                // GET ACTIVE WORKBOOK
                WB = ExcelObj.ActiveWorkbook;

                // GET THE WORKSHEET
                sheets = (Excel.Sheets)WB.Worksheets;
            }
            catch (Exception err)
            {
                if (iSection == 1)
                {
                    //GetObject MAY HAVE FAILED BECAUSE THE 
                    //Shell FUNCTION IS ASYNCHRONOUS; ENOUGH TIME HAS NOT ELAPSED
                    //FOR GetObject TO FIND THE RUNNING Office APPLICATION. WAIT
                    //1/2 SECONDS AND RETRY THE GetObject. IF YOU TRY 20 TIMES
                    //AND GETOBJECT STILL FAILS, ASSUME SOME OTHER REASON FOR GETOBJECT FAILING AND EXIT THE PROCEDURE.
                    iTries++;
                    if (iTries < 20)
                    {
                        System.Threading.Thread.Sleep(500); // WAIT 1/2 SECONDS.
                        //this.Activate();
                        goto tryAgain; //RESUME CODE AT THE getObject LINE.
                    }
                    else
                        MessageBox.Show("GetObject still failing. Process ended.");
                }
                else
                {
                    MessageBox.Show(err.Message);
                }
            }
        }

        public Boolean SheetExists(string sheetName)
        {
            Boolean foundSheet = false;

            // TEST CHANGED USING TO USING EXCEL = MICROSOFT.OFFICE......
            foreach (Excel.Worksheet findSheet in WB.Sheets)
            {
                if (findSheet.Name.Equals(sheetName))
                {
                    foundSheet = true;
                }
            }

            return foundSheet;
        }

        #endregion // Excel Behavior Region

        #region Control Behavior Region

        public void InsertMacrosSheetButtonEnabledState(Boolean enabled)
        {
            Globals.Ribbons.MacroExtenderRibbon.InsertMacrosSheetButton.Enabled = enabled;
        }

        public void ScopeSelectionCBoxEnabledState(Boolean enabled)
        {
            Globals.Ribbons.MacroExtenderRibbon.ScopeSelectionComboBox.Enabled = enabled;
        }

        public void MacroSelectionCBoxEnabledState(Boolean enabled)
        {
            Globals.Ribbons.MacroExtenderRibbon.MacroSelectionComboBox.Enabled = enabled;
        }

        public void RefreshMacrosButtonEnabledState(Boolean enabled)
        {
            Globals.Ribbons.MacroExtenderRibbon.RefreshMacrosButton.Enabled = enabled;
        }

        public void ExecuteMacroButtonEnabledState(Boolean enabled)
        {
            Globals.Ribbons.MacroExtenderRibbon.ExecuteMacroButton.Enabled = enabled;
        }

        public void OptionsButtonEnabledState(Boolean enabled)
        {
            Globals.Ribbons.MacroExtenderRibbon.OptionsButton.Enabled = enabled;
        }

        public void setScopeSelectionCBox()
        {
            string previousText = Globals.Ribbons.MacroExtenderRibbon.ScopeSelectionComboBox.Text;
            Globals.Ribbons.MacroExtenderRibbon.ScopeSelectionComboBox.Items.Clear();

            List<String> scopeList = new List<string>();

            bool foundMatch = false;
            scopeList.Add("All");

            for (int x = 0; x < MacrosList.Count; x++)
            {
                foundMatch = false;
                for (int y = 0; y < scopeList.Count; y++)
                {
                    if (MacrosList[x].Scope == scopeList[y])
                    {
                        foundMatch = true;
                    }
                }
                if (foundMatch == false)
                {
                    scopeList.Add(MacrosList[x].Scope);
                }
            }

            scopeList.Sort();

            for (int i = 0; i < scopeList.Count; i++)
            {
                RibbonDropDownItem item = makeRibbonDropDownItem(scopeList[i], null);
                Globals.Ribbons.MacroExtenderRibbon.ScopeSelectionComboBox.Items.Add(item);
            }

            Globals.Ribbons.MacroExtenderRibbon.ScopeSelectionComboBox.Text = previousText;
        }

        public void setMacroSelectionCBox()
        {
            //MacroSelectionCbox.Items.Clear();

            Globals.Ribbons.MacroExtenderRibbon.MacroSelectionComboBox.Items.Clear();

            List<string> macroList = new List<string>();

            switch (Globals.Ribbons.MacroExtenderRibbon.ScopeSelectionComboBox.Text)
            {
                case "All":
                    buildAllScopeMacroList(macroList);
                    break;

                case "Application":
                    buildApplicationScopeMacroList(macroList);
                    break;

                case "Workbook":
                    buildWorkbookScopeMacroList(macroList);
                    break;

                default:
                    buildWorksheetScopeMacroListWithScopeSelection(macroList);
                    break;
            }

            macroList.Sort();

            for (int x = 0; x < macroList.Count; x++)
            {
                RibbonDropDownItem item = makeRibbonDropDownItem(macroList[x], null);
                Globals.Ribbons.MacroExtenderRibbon.MacroSelectionComboBox.Items.Add(item);
            }
        }
        public void UpdateCBoxes()
        {
            BuildMacroList();
            setScopeSelectionCBox();
        }
        #endregion

        #region Helper Methods Region


        public void BuildMacroList()
        {
            MacrosList.Clear();

            int row = 2;

            //ExcelBase excelBase = new ExcelBase();

            Sheet = (Excel.Worksheet)Sheets.get_Item("Macros");

            Excel.Range scopeCell;

            Excel.Range worksheetCell;

            Excel.Range macroNameCell;

            do // LOOPING THROUGH THE ROWS (2-1000) LOOKING FOR INSTANCES
            //WHERE ALL THREE VARIABLES ARE PROVIDED; ONCE FOUND THE 
            //VARIABLES ARE ADDED TO A LIST OF THE MacrosStruct.
            {

                scopeCell = (Excel.Range)Sheet.get_Range(ScopeColumn + row);
                worksheetCell = (Excel.Range)Sheet.get_Range(WorksheetColumn + row);
                macroNameCell = (Excel.Range)Sheet.get_Range(MacroNameColumn + row);

                if ((scopeCell.Value2 != null)
                    && (macroNameCell.Value2 != null))
                {
                    var item = new MacrosStruct(row, macroNameCell.Value2, scopeCell.Value2, worksheetCell.Value2);
                    MacrosList.Add(item);
                }

                row++;
            } while (row < MaxSearchRow);
        }

        private List<string> buildAllScopeMacroList(List<string> macroList)
        {
            macroList = buildApplicationScopeMacroList(macroList);
            macroList = buildWorkbookScopeMacroList(macroList);
            macroList = buildWorksheetScopeMacroListWithoutScopeSelection(macroList);
            return macroList;
        }

        private List<string> buildApplicationScopeMacroList(List<string> macroList)
        {
            for (int x = 0; x < MacrosList.Count; x++)
            {
                if (MacrosList[x].Scope == "Application")
                {
                    macroList.Add(MacrosList[x].MacroName);
                }
            }
            return macroList;
        }

        private List<string> buildWorkbookScopeMacroList(List<string> macroList)
        {
            for (int x = 0; x < MacrosList.Count; x++)
            {
                if (MacrosList[x].Scope == "Workbook")
                {
                    macroList.Add(MacrosList[x].MacroName);
                }
            }
            return macroList;
        }

        private List<string> buildWorksheetScopeMacroListWithoutScopeSelection(List<string> macroList)
        {

            //ExcelBase excelBase = new ExcelBase();
            for (int x = 0; x < MacrosList.Count; x++)
            {
                if (MacrosList[x].Worksheet == ActivatedSheet.Name)
                {
                    macroList.Add(MacrosList[x].MacroName);
                }
            }
            return macroList;
        }

        private List<string> buildWorksheetScopeMacroListWithScopeSelection(List<string> macroList)
        {

            //ExcelBase excelBase = new ExcelBase();
            for (int x = 0; x < MacrosList.Count; x++)
            {
                if (MacrosList[x].Worksheet == ActivatedSheet.Name
                    && MacrosList[x].Scope == ScopeSelectionComboBox.Text)
                {
                    macroList.Add(MacrosList[x].MacroName);
                }
            }
            return macroList;
        }

        private RibbonDropDownItem makeRibbonDropDownItem(string Label, System.Drawing.Image Image)
        {
            RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            tmp.Label = Label;
            tmp.Image = Image;
            return tmp;
        }

        private string getBeginningRowScope(int beginningRow)
        {
            Boolean foundScope = false;

            int curIndex = 0;

            do
            {
                if (MacrosList[curIndex].Row == beginningRow)
                {
                    foundScope = true;
                }
                curIndex++;
            } while (foundScope == false);

            return MacrosList[curIndex].Scope;
        }

        public string getBeginningRowMacroName(int beginningRow)
        {
            Boolean foundMacroName = false;

            int curIndex = 0;

            do
            {
                if (MacrosList[curIndex].Row == beginningRow)
                {
                    foundMacroName = true;
                }
                curIndex++;
            } while (foundMacroName == false);

            return MacrosList[curIndex].MacroName;
        }

        #endregion // Helper Methods Region
    }
}
