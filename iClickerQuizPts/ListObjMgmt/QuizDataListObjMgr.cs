using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts.ListObjMgmt
{
    /// <summary>
    /// Provides a mechanism for interacting with the workbook's <see cref="Excel.ListObject"/> 
    /// of iClicker quiz data.
    /// </summary>
    public class QuizDataListObjMgr : ListObjectManager
    {
        /// <summary>
        /// Initializes a new instance of the 
        /// class <see cref="iClickerQuizPts.ListObjMgmt.QuizDataListObjMgr"/>.
        /// </summary>
        /// <param name="wshTblNmzPair">The properties of this <see langword="struct"/> 
        /// should be populated with the name of the <see cref="Excel.ListObject"/> 
        /// containing the iClicker quiz data and the name 
        /// of the parent <see cref="Excel.Worksheet"/>.</param>
        public QuizDataListObjMgr(WshListobjPair wshTblNmzPair) : base(wshTblNmzPair)
        {
        }


    }
}
