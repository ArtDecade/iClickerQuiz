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
    /// of Double-Dippers (i.e., students who have taken multiple quizzes for a given
    /// course week.
    /// </summary>
    public class DblDippersListObjMgr : ListObjectManager
    {
        /// <summary>
        /// Initializes a new instance of the 
        /// class <see cref="iClickerQuizPts.ListObjMgmt.DblDippersListObjMgr"/>.
        /// </summary>
        /// <param name="wshTblNmzPair">The properties of this <see langword="struct"/> 
        /// should be populated with the name of the <see cref="Excel.ListObject"/> 
        /// containing the double-dipping students and the name 
        /// of the parent <see cref="Excel.Worksheet"/>.</param>
        public DblDippersListObjMgr(WshListobjPair wshTblNmzPair) : base(wshTblNmzPair)
        {
        }
    }
}
