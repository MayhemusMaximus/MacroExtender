using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MacroExtender
{
    public class EngineBase : MacroExtenderRibbon
    {

        public string methodName;

        public Microsoft.Office.Interop.Excel.Range methodCell;
        public Microsoft.Office.Interop.Excel.Range arg1Cell;
        public Microsoft.Office.Interop.Excel.Range arg2Cell;
        public Microsoft.Office.Interop.Excel.Range arg3Cell;
        public Microsoft.Office.Interop.Excel.Range arg4Cell;
        public Microsoft.Office.Interop.Excel.Range arg5Cell;
        public Microsoft.Office.Interop.Excel.Range arg6Cell;
        public Microsoft.Office.Interop.Excel.Range arg7Cell;
        public Microsoft.Office.Interop.Excel.Range arg8Cell;
        public Microsoft.Office.Interop.Excel.Range arg9Cell;
        public Microsoft.Office.Interop.Excel.Range arg10Cell;
        public String arg1;
        public String arg2;
        public String arg3;
        public String arg4;
        public String arg5;
        public String arg6;
        public String arg7;
        public String arg8;
        public String arg9;
        public String arg10;


        //Microsoft.Office.Interop.Excel.Worksheet macroSheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(getBeginningRowWorksheet(beginningRow));

        protected void UpdateRanges(int curRow)
        {
            methodCell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(MethodColumn + curRow);
            arg1Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg1Column + curRow);
            arg2Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg2Column + curRow);
            arg3Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg3Column + curRow);
            arg4Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg4Column + curRow);
            arg5Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg5Column + curRow);
            arg6Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg6Column + curRow);
            arg7Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg7Column + curRow);
            arg8Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg8Column + curRow);
            arg9Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg9Column + curRow);
            arg10Cell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(Arg10Column + curRow);
            methodCell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(MethodColumn + curRow);

            methodName = methodCell.Value2;
            arg1 = userVariableCheck(Convert.ToString(arg1Cell.Value2), methodName);
            arg2 = userVariableCheck(Convert.ToString(arg2Cell.Value2), methodName);
            arg3 = userVariableCheck(Convert.ToString(arg3Cell.Value2), methodName);
            arg4 = userVariableCheck(Convert.ToString(arg4Cell.Value2), methodName);
            arg5 = userVariableCheck(Convert.ToString(arg5Cell.Value2), methodName);
            arg6 = userVariableCheck(Convert.ToString(arg6Cell.Value2), methodName);
            arg7 = userVariableCheck(Convert.ToString(arg7Cell.Value2), methodName);
            arg8 = userVariableCheck(Convert.ToString(arg8Cell.Value2), methodName);
            arg9 = userVariableCheck(Convert.ToString(arg9Cell.Value2), methodName);
            arg10 = userVariableCheck(Convert.ToString(arg10Cell.Value2), methodName);
        }

        //TEST - BREAKOUT CHANGED FROM PRIVATE TO PUBLIC
        public string userVariableCheck(string arg, string methodName)
        {
            int stringLength = 0;
            if (arg != null)
            {
                stringLength = arg.Length;
            }

            string returnString = arg;

            if (methodName != "Input Box")
            {
                if (stringLength > 3)
                {
                    if (userIntIndexRef.Count() != 0
                            && StringExt.Left(arg, 3) == "Int")
                    {
                        int key = Convert.ToInt32(StringExt.Right(arg, 1));
                        foreach (KeyValuePair<int, int> pair in userIntIndexRef)
                        {
                            if (pair.Key == key)
                            {
                                returnString = Convert.ToString(userIntegerList[pair.Value]);
                            }
                        }
                    }
                    else if (userStrIndexRef.Count() != 0
                                && StringExt.Left(arg, 3) == "Str")
                    {
                        int key = Convert.ToInt32(StringExt.Right(arg, 1));
                        foreach (KeyValuePair<int, int> pair in userStrIndexRef)
                        {
                            if (pair.Key == key)
                            {
                                returnString = Convert.ToString(userStringList[pair.Value]);
                            }
                        }
                    }
                }
            }

            return returnString;
        }
    }
}
