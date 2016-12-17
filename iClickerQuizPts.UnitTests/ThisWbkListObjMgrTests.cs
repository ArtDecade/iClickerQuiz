using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPts;
using iClickerQuizPts.AppExceptions;
using NUnit.Framework;
using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts.UnitTests
{
    [TestFixture]
    public class ThisWbkListObjMgrTests
    {
        [Category("ThisWbkListObjMgr")]
        [Test]
        public void ThisWbkLOMgrGetInstance_QuizPtsTblMissing_ThrowsException()
        {
            NeverFindsQuizPtsTblFakeThisWbkListObjMgr stub;

            var ex = Assert.Catch<MissingListObjectException>(() => 
                stub = (NeverFindsQuizPtsTblFakeThisWbkListObjMgr)NeverFindsQuizPtsTblFakeThisWbkListObjMgr.GetStubInstance() );

            StringAssert.Contains("tblClkrQuizGrades", ex.WshListObjPair.ListObjName);
        }

        [Category("ThisWbkListObjMgr")]
        [Test]
        public void ThisWbkLOMgrGetInstance_DblDpprsTblMissing_ThrowsException()
        {
            NeverFindsDblDpprsTblFakeLThisWbkListObjMgr stub;

            var ex = Assert.Catch<MissingListObjectException>(() =>
                stub = (NeverFindsDblDpprsTblFakeLThisWbkListObjMgr)NeverFindsDblDpprsTblFakeLThisWbkListObjMgr.GetStubInstance());

            StringAssert.Contains("tblDblDippers", ex.WshListObjPair.ListObjName);
        }

        [Category("ThisWbkListObjMgr")]
        [Test]
        public void ThisWbkLOMgrGetInstance_AllTblsExist_ListObjsPopulatedPptyIsTrue()
        {
            AlwaysFindsListObjectsThisWbkListObjMgr stub;

            stub = (AlwaysFindsListObjectsThisWbkListObjMgr)AlwaysFindsListObjectsThisWbkListObjMgr.GetStubInstance();

            Assert.True(stub.ListObjectsPopulated);
        }
    }

    public class NeverFindsQuizPtsTblFakeThisWbkListObjMgr : ThisWbkListObjectManager
    {
        public static ThisWbkListObjectManager GetStubInstance()
        {
           _twh = new NeverFindsQuizPtsTblFakeThisWbkListObjMgr();
            return _twh;
        }

        protected override void SetWshListObjPairs()
        {
            _quizPtsWshAndTbl = new WshListobjPairs("tblClkrQuizGrades", "Sheet1FakeName");
            _dblDpprsWshAndTbl = new WshListobjPairs("tblDblDippers", "Sheet2FakeName");
        }
        protected override bool DoesTtlQuizPtsListObjectExist()
        {
            return false;
        }
        protected override bool DoesDblDippersListObjectExist()
        {
            return true;
        }
    }

    public class NeverFindsDblDpprsTblFakeLThisWbkListObjMgr : ThisWbkListObjectManager
    {
        public static ThisWbkListObjectManager GetStubInstance()
        {
            _twh = new NeverFindsDblDpprsTblFakeLThisWbkListObjMgr();
            return _twh;
        }

        protected override void SetWshListObjPairs()
        {
            _quizPtsWshAndTbl = new WshListobjPairs("tblClkrQuizGrades", "Sheet1FakeName");
            _dblDpprsWshAndTbl = new WshListobjPairs("tblDblDippers", "Sheet2FakeName");
        }
        protected override bool DoesDblDippersListObjectExist()
        {
            return false;
        }
        protected override bool DoesTtlQuizPtsListObjectExist()
        {
            return true;
        }
    }

    public class AlwaysFindsListObjectsThisWbkListObjMgr : ThisWbkListObjectManager
    {
        public static ThisWbkListObjectManager GetStubInstance()
        {
            _twh = new AlwaysFindsListObjectsThisWbkListObjMgr();
            return _twh;
        }
        protected override void SetWshListObjPairs()
        {
            _quizPtsWshAndTbl = new WshListobjPairs("tblClkrQuizGrades", "Sheet1FakeName");
            _dblDpprsWshAndTbl = new WshListobjPairs("tblDblDippers", "Sheet2FakeName");
        }
        protected override bool DoesTtlQuizPtsListObjectExist()
        {
            return true;
        }
        protected override bool DoesDblDippersListObjectExist()
        {
            return true;
        }
        protected override void SetListObjectFields()
        {
            _listObjsPopulated = true;
        }
    }
}
