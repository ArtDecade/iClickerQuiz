using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using iClickerQuizPts.AppExceptions;

namespace iClickerQuizPts
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
        protected bool _listObjHasData = false;
        #endregion
        #endregion

        public virtual bool ListObjectHasData
        {
            get
            { return _listObjHasData; }
        }

        protected ListObjectManager(WshListobjPair wshTblNmzPair)
        {
            if(wshTblNmzPair.PptsSet)
                _wshLoPr = wshTblNmzPair;
            else
            {
                InvalidWshListObjPairException ex = new InvalidWshListObjPairException();
                ex.WshListObjPair = wshTblNmzPair;
                throw ex;
            }


            if (!DoesParentWshExist())
            {
                MissingWorksheetException ex = new MissingWorksheetException();
                ex.WshListObjPair = _wshLoPr;
                throw ex;
            }
            else
                _ws = Globals.ThisWorkbook.Worksheets[_wshLoPr.WshNm];

            if (!DoesListObjExist())
            {
                MissingListObjectException ex = new MissingListObjectException();
                ex.WshListObjPair = _wshLoPr;
                throw ex;
            }
            else
                _lo = _ws.ListObjects[_wshLoPr.ListObjName];

            _listObjHasData = DoesListObjHaveData();
        }

        protected virtual bool DoesParentWshExist()
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

        protected virtual bool DoesListObjExist()
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

        protected virtual bool DoesListObjHaveData()
        {
            bool hasData = false;

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


        
    }
}
