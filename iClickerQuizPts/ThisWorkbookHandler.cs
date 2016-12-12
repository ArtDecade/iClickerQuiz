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
    /// Specifies constants defining which session/recitation within the semester week the grades are from.
    /// </summary>
    /// <remarks>
    /// For the purposes of this program it is assumed that the course is taught in weekly segments.  
    /// It is further assumed that within each week there are three recitation in which a student 
    /// can take a quiz.  This enumeration delineates those three recitation sessions.
    /// <para>It should be noted that a student is only supposed to take each week's iClicker quiz 
    /// once.  It has been discovered, however, that some students were attending multiple recitations 
    /// within a week and taking a week's quiz more than once.  Further, there has been no mechanism 
    /// - other than manual review of the data - for identifying students who take a week's iClicker 
    /// quiz more than once.  The entire purpose of this program is to identify and filter out those 
    /// duplicate quiz scores.</para>
    /// </remarks>
    public enum WkSession : byte
    {
        /// <summary>No session has been selected yet.
        /// </summary>
        None = 0,
        /// <summary>
        /// First recitation of a given week.
        /// </summary>
        First,
        /// <summary>Second recitation of a given week.
        /// </summary>
        Second,
        /// <summary>
        /// Third recitation of a given week.
        /// </summary>
        Third
    }

    /// <summary>
    /// Provides a mechanism for pairing the name of each Excel <see cref="Excel.ListObject"/> 
    /// (i.e., Table) with its parent <see cref="Excel.Worksheet"/>. </summary>
    /// <remarks>Each of the three worksheets in this workbook contains a named 
    /// <see cref="Excel.ListObject"/>.  The <see cref="ThisWorkbook.SetListObjects"/> method 
    /// utilizes the information stored in instances of this struct in order to verify that 
    /// the basic structure of this <see cref="Excel.Workbook"/> has not been altered.</remarks>
    public struct WshListobjPairs
    {
        /// <summary>
        /// Gets the name of the Excel <see cref="Excel.ListObject"/> (i.e., Table) within 
        /// one of <c>ThisWorkbook's</c> Sheet.
        /// </summary>
        public string ListObjName { get; }
        /// <summary>
        /// Gets the name of the <c>Sheet</c> holding the identified <see cref="Excel.ListObject"/>
        /// </summary>
        public string WshNm { get; set; }
        /// <summary>
        /// Gets a value indicating whether both <see cref="WshListobjPairs.ListObjName"/> and
        /// <see cref="WshListobjPairs.WshNm"/> properties have been populated.
        /// </summary>
        /// <remarks>This value is set in the <see cref="WshListobjPairs"/> custom constructor.  
        /// It is only set to <c>true</c> if non-empty, non-null values are provided for both 
        /// <see cref="WshListobjPairs.ListObjName"/> and <see cref="WshListobjPairs.WshNm"/>.
        /// <para>If the structure is instantiated via its default constructor 
        /// (which should not be used) then the value 
        /// of this property will of course remain at its default value of <c>false</c>.</para> </remarks>
        public bool PptsSet { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WshListobjPairs"/> struct.
        /// </summary>
        /// <param name="listObjNm">The name of the <see cref="Excel.ListObject"/> which the
        /// paired <see cref="Excel.Worksheet"/> contains.</param>
        /// <param name="wshNm">A worksheet within this workbook.</param>
        /// <remarks>Each worksheet in this workbook contains contains (or should contain) one
        /// and only one named <see cref="Excel.ListObject"/>.</remarks>
        public WshListobjPairs(string listObjNm, string wshNm) : this()
        {
            // Set structure properties...
            ListObjName = listObjNm;
            WshNm = wshNm;
            if (!string.IsNullOrEmpty(listObjNm) && !string.IsNullOrEmpty(wshNm))
                PptsSet = true;
            else
                PptsSet = false; // ...just to be certain
        }
    }

    /// <summary>
    /// Provides a number of properties of <c>ThisWorkbook</c> 
    /// </summary>
    /// <remarks>The properties are places here, rather than on the <c>ThisWorkbook</c> class
    /// itself in order to facilitate unit testing.  (<c>ThisWorkbook</c> is a sealed class, 
    /// which, for example, obviates extracting and overriding any methods we might 
    /// write within the class.) 
    /// <para>NOTE:  This class is a singleton.</para></remarks>
    public class ThisWorkbookHandler
    {
        #region Fields
        private bool _virginWbk = false;
        private QuizUserControl _ctrl = new QuizUserControl();
        private List<DateTime> _qDts = new List<DateTime>();
        private Excel.ListObject _tblQuizGrades = null;
        private Excel.ListObject _tblDDs = null;
        private List<WshListobjPairs> _listObjsByWsh = new List<WshListobjPairs>();
        private ThisWorkbookHandler _twh = null;
        private Excel.Workbook _wbk = null;
        #endregion

        private ThisWorkbookHandler()
        {
            // PopulateListOfWshListObjectPairs();
            if(DoesTtlQuizPtsListObjectExist()==false)
            {
                throw new MissingListObjectException();
            }
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

        #endregion



        /// <summary>
        /// The one and only method by which one obtains an instance of this class.  
        /// </summary>
        /// <remarks>The <see cref="ThisWorkbookHandler"/> class is a singleton.  
        /// As such, its constructor has been defined as <see langword="private"/>.</remarks>
        /// <returns>A (singleton) instance of <see cref="ThisWorkbookHandler"/>.</returns>
        public ThisWorkbookHandler GetInstance()
        {
            if (_twh == null)
                _twh = new ThisWorkbookHandler();
            return _twh;
        }

        // TODO:  Make PopulateListOfWshListObjectPairs private?
        /// <summary>
        /// Populates the the private <see cref="List{T}"/> containing 
        /// <see cref="iClickerQuizPts.WshListobjPairs"/>.  
        /// </summary>
        /// <remarks>There should be one item in the list for each <c>Sheet</c> in 
        /// <c>ThisWorkbook.</c>.  Further, each such <see cref="iClickerQuizPts.WshListobjPairs"/> 
        /// should be populated with the name of a <c>Sheet</c> and the name of 
        /// the <see cref="Excel.ListObject"/> it contains.
        /// </remarks>
        public void PopulateListOfWshListObjectPairs()
        {
            /* NOTE:  We are not trapping here to verify the existence of the 
             * specified Sheets.  May God have mercy on the end-user who 
             * deletes one of those sheets.*/
            _listObjsByWsh.Add(new WshListobjPairs("tblClkrQuizGrades", Globals.Sheet1.Name));
            _listObjsByWsh.Add(new WshListobjPairs("tblDblDippers", Globals.Sheet2.Name));
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
        public virtual bool DoesTtlQuizPtsListObjectExist()
        {
            bool loExists = false;
            int nmbrWshTbls = Globals.Sheet1.ListObjects.Count;

            if (nmbrWshTbls == 0 )
                return loExists;
            else
            {
                for(int i = 1; i <= nmbrWshTbls; i++)
                {
                    if(Globals.Sheet1.ListObjects[i].Name == "tblClkrQuizGrades")
                    {
                        _tblQuizGrades = Globals.Sheet1.ListObjects[i];
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
        public bool DoesDblDippersListObjectExist()
        {
            bool loExists = false;
            int nmbrWshTbls = Globals.Sheet2.ListObjects.Count;

            if (nmbrWshTbls == 0)
                return loExists;
            else
            {
                for (int i = 1; i <= nmbrWshTbls; i++)
                {
                    if(Globals.Sheet2.ListObjects[i].Name == "tblDblDippers")
                    {
                        _tblDDs = Globals.Sheet2.ListObjects[i];
                        loExists = true;
                        i = nmbrWshTbls; // ...break loop
                    }
                }
                return loExists;
            }
        }




        public void SetListObjectsProperties()
        {

        }

    }
}
