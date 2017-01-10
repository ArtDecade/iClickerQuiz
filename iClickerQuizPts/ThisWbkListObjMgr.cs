using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPts.AppExceptions;
using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides a number of properties of <c>ThisWorkbook</c> 
    /// </summary>
    /// <remarks>The properties are places here, rather than on the <c>ThisWorkbook</c> class
    /// itself in order to facilitate unit testing.  (<c>ThisWorkbook</c> is a sealed class, 
    /// which, for example, obviates extracting and overriding any methods we might 
    /// write within the class.) 
    /// <para>NOTE:  This class is a singleton.</para></remarks>
    public class ThisWbkListObjectManager
    {
        #region Fields
        #region PrivateFlds
        private Excel.ListObject _tblQuizGrades = null;
        private Excel.ListObject _tblDDs = null;
        #endregion
        #region ProtectedFlds
        protected static ThisWbkListObjectManager _twh = null;
        protected WshListobjPair _quizPtsWshAndTbl;
        protected WshListobjPair _dblDpprsWshAndTbl;
        protected bool _listObjsPopulated = false;
        #endregion
        #endregion



        /// <summary>
        /// Creates a (<see langword="protected"/>) instance of the class.
        /// </summary>
        protected ThisWbkListObjectManager()
        {
            SetWshListObjPairs();

            // Trap for missing ListObjects...
            if (DoesTtlQuizPtsListObjectExist()==false)
            {
                MissingListObjectException ex = 
                    new MissingListObjectException { WshListObjPair = _quizPtsWshAndTbl };
                throw ex;
            }

            if (DoesDblDippersListObjectExist()==false)
            {
                MissingListObjectException ex = 
                    new MissingListObjectException { WshListObjPair = _dblDpprsWshAndTbl };
                throw ex;
            }

            SetListObjectFields();
        }

        #region Ppts
        /// <summary>
        /// Gets the <see cref="Excel.ListObject"/> representing the master, parsed
        /// list (table) of students' quiz grades.
        /// </summary>
        public Excel.ListObject TblQuizGrades
        {
            get
            { return _tblQuizGrades; }
        }

        /// <summary>
        /// Gets the <see cref="Excel.ListObject"/> which contains the list (table) of students
        /// who have taken more than one iClicker quiz during a given semester week.
        /// </summary>
        /// <remarks>
        /// The scores this table are those which have been excluded from the master
        /// table used to calculate each student's end-of-semester quiz totals.
        /// </remarks>
        public Excel.ListObject TblDoubleDippers
        {
            get
            { return _tblDDs; }
        }

        /// <summary>
        /// Gets whether the ListObject properties have been populated.
        /// </summary>
        /// <remarks>This property is included for purposes of unit testing.</remarks>
        public bool ListObjectsPopulated
        {
            get
            { return _listObjsPopulated; }
        }
        #endregion

        /// <summary>
        /// The one and only method by which one obtains an instance of this class.  
        /// </summary>
        /// <remarks>The <see cref="ThisWbkListObjectManager"/> class is a singleton.  
        /// As such, its constructor has been defined as <see langword="private"/>.</remarks>
        /// <returns>A (singleton) instance of <see cref="ThisWbkListObjectManager"/>.</returns>
        public static ThisWbkListObjectManager GetInstance()
        {
            if (_twh == null)
                _twh = new ThisWbkListObjectManager();
            return _twh;
        }

        /*The following 2 methods seem like huge DRY-violation code smells.  However, there
        * doesn't seem to be any way to do this more efficiently.  (Trust me - I went pretty far 
        * down some obvious roads towards that end.  I created a struct so that I could pair
        * worksheet names with ListObject names, and then created a generic List<T> of that 
        * type/struct.  The goal was to loop through the members of that generic List<T> in 
        * one, compact method.  Ultimately, however, that seemingly simple approach 
        * became unwieldy.) */
        /// <summary>
        /// Confirms (or not) that the named ListObject of total quiz points 
        /// still exists.
        /// </summary>
        /// <returns>
        /// <c>true</c> if the ListObject still exist; otherwise <c>false</c>.
        /// </returns>
        protected virtual bool DoesTtlQuizPtsListObjectExist()
        {
            bool loExists = false;
            int nmbrWshTbls = Globals.Sheet1.ListObjects.Count;

            if (nmbrWshTbls == 0 )
                return loExists;
            else
            {
                for(int i = 1; i <= nmbrWshTbls; i++)
                {
                    if(Globals.Sheet1.ListObjects[i].Name == _quizPtsWshAndTbl.ListObjName)
                    {
                        loExists = true;
                        i = nmbrWshTbls; // ...break loop
                    }
                }
                return loExists;
            }
        }

        /// <summary>
        /// Confirms (or not) that the named ListObject of students who have 
        /// taken a quiz more than once within a week still exists.
        /// </summary>
        /// <returns>
        /// <c>true</c> if the ListObject still exist; otherwise <c>false</c>.
        /// </returns>
        protected virtual bool DoesDblDippersListObjectExist()
        {
            bool loExists = false;
            int nmbrWshTbls = Globals.Sheet2.ListObjects.Count;

            if (nmbrWshTbls == 0)
                return loExists;
            else
            {
                for (int i = 1; i <= nmbrWshTbls; i++)
                {
                    if(Globals.Sheet2.ListObjects[i].Name == _dblDpprsWshAndTbl.ListObjName)
                    {
                        loExists = true;
                        i = nmbrWshTbls; // ...break loop
                    }
                }
                return loExists;
            }
        }

        /// <summary>
        /// Sets the private fields containing the <see cref="Excel.ListObject"/> names with 
        /// the name of their respective <see cref="Excel.Worksheet"/>.
        /// </summary>
        protected virtual void SetWshListObjPairs()
        {
            _quizPtsWshAndTbl = new WshListobjPair("tblClkrQuizGrades", Globals.Sheet1.Name);
            _dblDpprsWshAndTbl = new WshListobjPair("tblDblDippers", Globals.Sheet1.Name);
        }

        /// <summary>
        /// Sets the <see cref="Excel.ListObject"/> fields.
        /// </summary>
        protected virtual void SetListObjectFields()
        {
            _tblQuizGrades = Globals.Sheet1.ListObjects[_quizPtsWshAndTbl.ListObjName];
            _tblDDs = Globals.Sheet2.ListObjects[_dblDpprsWshAndTbl.ListObjName];
            _listObjsPopulated = true;
        }
    }
}
