using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace MacroExtender
{
    public partial class ThisAddIn
    {

        #region FIELDS AND PROPERTIES

        private Excel.AppEvents_Event EventDel_WorkbookActivate;
        private Excel.AppEvents_Event EventDel_WorkbookOpen;
        private Excel.AppEvents_Event EventDel_WorkbookBeforeClose;

        //////APIEventsManager eventsManager = new APIEventsManager();
        
        #endregion

        #region

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            MacroExtenderRibbon thisInstance = new MacroExtenderRibbon();

            EventDel_WorkbookActivate = (Excel.AppEvents_Event)this.Application;
            EventDel_WorkbookActivate.WorkbookActivate +=
                new Excel.AppEvents_WorkbookActivateEventHandler(thisInstance.excelEvents_WorkbookActivate);

            EventDel_WorkbookOpen = (Excel.AppEvents_Event)this.Application;
            EventDel_WorkbookOpen.WorkbookOpen += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(thisInstance.excelEvents_WorkbookOpen);

            EventDel_WorkbookBeforeClose = (Excel.AppEvents_Event)this.Application;
            EventDel_WorkbookBeforeClose.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(thisInstance.excelEvents_WorkbookBeforeClose);
            this.Application.SheetActivate +=
                new Excel.AppEvents_SheetActivateEventHandler(thisInstance.excelEvents_SheetActivate);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
