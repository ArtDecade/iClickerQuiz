using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using iClickerQuizPts.AppExceptions;

namespace iClickerQuizPts.ListObjMgmt
{
    /// <summary>
    /// Provides a mechanism for interacting with the workbook's <see cref="Excel.ListObjects"/>.
    /// </summary>
    public abstract class ListObjectManager
    {
        #region Fields
        #region PrivateFlds
        private Excel.Worksheet _ws = null;
        private Excel.ListObject _lo = null;
        private WshListobjPair _wshLoPr;
        #endregion
        #region ProtectedFlds
        /// <summary>
        /// Holds a value indicating whether the underlying <see cref="Excel.ListObject"/> 
        /// contains data.
        /// </summary>
        protected bool _listObjHasData = false;
        #endregion
        #endregion

        #region ppts
        /// <summary>
        /// Gets a value indicating whether the underlying 
        /// <see cref="Excel.ListObject"/> has yet been populated 
        /// with any data.
        /// </summary>
        public virtual bool ListObjectHasData
        {
            get
            { return _listObjHasData; }
        }
        #endregion

        #region ctor
        /// <summary>
        /// Initializes a new instance of the 
        /// class <see cref="iClickerQuizPts.ListObjMgmt.ListObjectManager"/>.
        /// </summary>
        /// <param name="wshTblNmzPair">The properties of this <see langword="struct"/> 
        /// should be populated with the name of the <see cref="Excel.ListObject"/> 
        /// and the name of the parent <see cref="Excel.Worksheet"/>.</param>
        protected ListObjectManager(WshListobjPair wshTblNmzPair)
        {
            // Trap to ensure that constructor parameter has been populated with both
            // a wsh name and a ListObject name...
            if (wshTblNmzPair.PptsSet)
                _wshLoPr = wshTblNmzPair;
            else
            {
                InvalidWshListObjPairException ex = new InvalidWshListObjPairException();
                ex.WshListObjPair = wshTblNmzPair;
                throw ex;
            }

            // Trap to ensure that the worksheet in which the ListObject resides
            // has not been deleted or renamed by the user...
            if (!DoesParentWshExist())
            {
                MissingWorksheetException ex = new MissingWorksheetException();
                ex.WshListObjPair = _wshLoPr;
                throw ex;
            }
            else
                _ws = Globals.ThisWorkbook.Worksheets[_wshLoPr.WshNm];

            // Trap to ensure that the ListObject we need has not been deleted
            // or renamed by the user...
            if (!DoesListObjExist())
            {
                MissingListObjectException ex = new MissingListObjectException();
                ex.WshListObjPair = _wshLoPr;
                throw ex;
            }
            else
                _lo = _ws.ListObjects[_wshLoPr.ListObjName];

            // If here then the ListObject (and its parent wsh) exist.  Now see if 
            // the table has yet been populated with any data...
            _listObjHasData = DoesListObjHaveData();
        }
        #endregion

        #region methods
        /// <summary>
        /// Determines whether the parent <see cref="Excel.Worksheet"/> of 
        /// the <see cref="Excel.ListObject"/> exists.
        /// </summary>
        /// <returns><c>true</c> if the <see cref="Excel.Worksheet"/> exists; 
        /// otherwise <c>false</c>.</returns>
        public virtual bool DoesParentWshExist()
        {
            bool exists = false;
            int noWshs = Globals.ThisWorkbook.Worksheets.Count;
            for (int i = 1; i <= noWshs; i++)
            {
                Excel.Worksheet ws = Globals.ThisWorkbook.Worksheets[i];
                if(ws.Name == _wshLoPr.WshNm)
                {
                    exists = true;
                    break;
                }
            }
            return exists;
        }

        /// <summary>
        /// Determines whether the the <see cref="Excel.ListObject"/> exists.
        /// </summary>
        /// <returns><c>true</c> if the <see cref="Excel.ListObject"/> exists; 
        /// otherwise <c>false</c>.</returns>
        public virtual bool DoesListObjExist()
        {
            bool exists = false;
            int tbls = _ws.ListObjects.Count;

            if (tbls == 0)
                return exists;
            else
            {
                for (int i = 1; i <= tbls; i++)
                {
                    Excel.ListObject tbl;
                    tbl = _ws.ListObjects[i];
                    if (tbl.Name == _wshLoPr.ListObjName)
                    {
                        exists = true;
                        break;
                    }
                }
                return exists;
            }
        }

        /// <summary>
        /// Determines whether the the <see cref="Excel.ListObject"/> has yet 
        /// been populated with any data.
        /// </summary>
        /// <returns><c>true</c> if the <see cref="Excel.ListObject"/> 
        /// contains data; otherwise <c>false</c>.</returns>
        protected virtual bool DoesListObjHaveData()
        {
            bool hasData = false; // ...default

            // Now see if there are data...
            if (_lo.ListRows.Count > 1)
                hasData = true;
            else
            {
                Excel.Range c;
                for(int i = 1; i <= _lo.ListColumns.Count;i++)
                {
                    c = _lo.DataBodyRange[1, i];
                    if(c.Value2 != null)
                    {
                        hasData = true;
                        break;
                    }
                }
            }
            return hasData; 
        }
        #endregion
    }
}
