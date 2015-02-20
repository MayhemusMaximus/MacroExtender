using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace MacroExtender
{
    class InteractiveFormatting : MacroExtenderRibbon
    {
        #region FIELDS AND PROPERTIES

        #endregion

        #region "MAIN" METHODS

        /// <summary>
        /// This method loops through cells in the provided range passing each to the delegateColumn method.
        /// </summary>
        /// <param name="target">A range to consider for formatting.</param>
        public void delegateCell(Excel.Range target)
        {
            // THE FUNCTIONALITY OF THIS CLASS STARTS HERE.
            // CELL IN THE PASSED RANGE IS PASSED TO THE
            // delegateColumn METHOD FOR FURTHER DELEGATION.
            foreach (Excel.Range cell in target)
            {
                delegateColumn(cell);
            }
        }

        private void delegateColumn(Microsoft.Office.Interop.Excel.Range target)
        {
            // THE COLUMN THAT THE PASSED CELL IS IN
            // DETERMINES HOW WE WANT TO APPROACH FORMATTING.
            string targetColumnName = getExcelColumnName(target.Column);

            switch (targetColumnName)
            {
                case ScopeListColumn:
                    break;

                case MethodListColumn:
                    break;

                case ScopeColumn:
                    break;

                case WorksheetColumn:
                    break;

                case MacroNameColumn:
                    break;

                case MethodColumn:

                    // CREATE AN INSTANCE OF MethodFormatting.cs
                    MethodFormatting methodForm = new MethodFormatting();

                    // PASS THE INSTANCE THE TARGET FOR REVIEW.
                    methodForm.delegateMethodFormatting(target);
                    break;

                case Arg1Column:
                    break;

                case Arg2Column:
                    break;

                case Arg3Column:
                    break;

                case Arg4Column:
                    break;

                case Arg5Column:
                    break;

                case Arg6Column:
                    break;

                case Arg7Column:
                    break;

                case Arg8Column:
                    break;

                case Arg9Column:
                    break;

                case Arg10Column:
                    break;

                default:
                    break;
            }
        }

        #endregion

        public void protectWorksheet()
        {
            //Sheet.Protect("Robert");//,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false);
        }

        public void unProtectWorksheet()
        {
            //Sheet.Protect("Robert");//,
            //Sheet.Unprotect("Robert");
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false,
            //false);
        }

        public double SelectionColor(Excel.Range target, Excel.Range lastTarget, double lastTargetColor)
        {

            double initialColor = target.Interior.Color;

            if (lastTarget != null)
            {
                if (lastTargetColor == 16777215.0) // THIS COLOR REPRESENTS ORIGINAL WHITE, HOWEVER IF YOU USE THIS VALUE TO CHANGE THE CELL TO WHITE YOU LOOSE THE GRID.
                    lastTarget.Interior.ColorIndex = 0;
                else
                    lastTarget.Interior.Color = lastTargetColor;
            }

            target.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Aqua);

            return initialColor;
        }
        private string getExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
