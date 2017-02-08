using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPts.AppExceptions;
using iClickerQuizPts.Itfs;
using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides a mechanism for verifying that named ranges still exist in workbook. 
    /// </summary>
    public class NamedRangeWrapper : INamedRangeWrapping
    {
        /// <summary>
        /// Tests whether a specified workbook-scoped <see cref="Excel.Name"/> both exists 
        /// and refers to a valid <see cref="Excel.Range"/>.
        /// </summary>
        /// <param name="rngNm">The name of the <see cref="Excel.Range"/>.</param>
        /// <returns><c>true</c> if the <see cref="Excel.Name"/> both exists and refers to a valid 
        /// <see cref="Excel.Range"/>; otherwise <c>false</c>.</returns>
        public virtual bool WorkbookScopedRangeExists(string rngNm)
        {
            bool rngFound = false;
            int nmbrWbkNmz = Globals.ThisWorkbook.Names.Count;
            Excel.Name wbkNm;

            for(int i = 1; i <= nmbrWbkNmz; i++)
            {
                wbkNm = Globals.ThisWorkbook.Names.Item(i);
                if(wbkNm.Name == rngNm)
                {
                    rngFound = true;
                    break;
                }
            }

            if (!rngFound)
                return false;
            else // ...the named range exists
            {
                // Compiler needs to see that we have, in fact, assigned a value to this variable...
                wbkNm = Globals.ThisWorkbook.Names.Item(rngNm);

                // Now see if the named range has a valid reference...
                try
                {
                    Excel.Range r = wbkNm.RefersToRange;
                }
                catch
                {
                    rngFound = false;
                }
            }

            // If the catch clause was not hit the variable remains true...
            return rngFound;
        }

        /// <summary>
        /// Tests whether a specified worksheet-scoped <see cref="Excel.Name"/> both exists 
        /// and refers to a valid <see cref="Excel.Range"/>.
        /// </summary>
        /// <param name="wshNm">The name of the parent <see cref="Excel.Worksheet"/>.</param>
        /// <param name="rngNm">The name of the <see cref="Excel.Range"/>.</param>
        /// <returns><c>true</c> if the <see cref="Excel.Name"/> both exists and refers to a valid 
        /// <see cref="Excel.Range"/>; otherwise <c>false</c>.</returns>
        public virtual bool WorksheetScopedRangeExists(string wshNm, string rngNm)
        {
            bool rngFound = false;
            Excel.Worksheet ws = Globals.ThisWorkbook.Worksheets.Item[wshNm];
            int nmbrWshNms = ws.Names.Count;
            Excel.Name XLnm;

            for (int i = 1; i <= nmbrWshNms; i++)
            {
                XLnm = ws.Names.Item(i);
                if (XLnm.Name == rngNm)
                {
                    rngFound = true;
                    break;
                }
            }

            if (!rngFound)
                return false;
            else // ...the named range exists
            {
                // Compiler needs to see that we have, in fact, assigned a value to this variable...
                XLnm = ws.Names.Item(rngNm);

                // Now see if the named range has a valid reference...
                try
                {
                    Excel.Range r = XLnm.RefersToRange;
                }
                catch
                {
                    rngFound = false;
                }
            }

            // If the catch clause was not hit the variable remains true...
            return rngFound;
        }
    }
}
