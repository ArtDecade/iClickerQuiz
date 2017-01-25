using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPts;
using iClickerQuizPts.AppExceptions;
using iClickerQuizPts.ListObjMgmt;

using NUnit.Framework;
using NSubstitute;

using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts.UnitTests
{
    public class GenericListObjMgr : ListObjectManager
    {
        public GenericListObjMgr(WshListobjPair pr) : base(pr)
        {
        }
    }

    [TestFixture]
    [Category("ListObjectManagerTests")]
    public class ListObjectManagerDerivedTests
    {
        const string QZ_GRADES_TBL = "tblClkrQuizGrades";
        const string QZ_GRADES_WSH = "iCLICKERQuizPoints";
        const string DBL_DPPRS_TBL = "tblDblDippers";
        const string DBL_DPPRS_WSH = "DoubleDippers";

        [TestCase(QZ_GRADES_WSH, QZ_GRADES_TBL)]
        [TestCase(DBL_DPPRS_WSH, DBL_DPPRS_TBL)]
        [TestCase("foo","bar")]
        public void InstantiateListObjectMgrs_NonEmptyWshTblNmz_Succeeds(string wshNm, string tblNm)
        {
            // Arrange & act...
            WshListobjPair pr = new WshListobjPair(tblNm, wshNm);
            GenericListObjMgr mgr = new GenericListObjMgr(pr);
           
            // Assert...
            Assert.IsInstanceOf<GenericListObjMgr>(mgr);
        }

        [TestCase(QZ_GRADES_WSH,"")]
        [TestCase("",QZ_GRADES_TBL)]
        [TestCase(DBL_DPPRS_WSH, "")]
        [TestCase("", DBL_DPPRS_TBL)]
        [TestCase("","")]
        public void InstantiateListObjMgrs_MissingCtrParams_Throws(string wshNm, string tblNm)
        {
            WshListobjPair pr = new WshListobjPair(tblNm, wshNm);
            GenericListObjMgr mgr;

            var ex = Assert.Catch<InvalidWshListObjPairException>(() =>
                mgr = new GenericListObjMgr(pr));
        }

        [TestCase("foo","bar")]
        public void SetListObjAndParentWshPpts_MissingWsh_Throws(string wshNm, string tblNm)
        {
            WshListobjPair pr = new WshListobjPair(tblNm, wshNm);
            var mgr = Substitute.ForPartsOf<GenericListObjMgr>(pr);

            mgr.When(x => x.DoesParentWshExist()).DoNotCallBase();
            mgr.DoesParentWshExist().Returns(false);

            var ex = Assert.Catch<MissingWorksheetException>(() =>
                mgr.SetListObjAndParentWshPpts());
        }

        [TestCase("foo", "bar")]
        public void SetListObjAndParentWshPpts_MissingListObj_Throws(string wshNm, string tblNm)
        {
            WshListobjPair pr = new WshListobjPair(tblNm, wshNm);
            var mgr = Substitute.ForPartsOf<GenericListObjMgr>(pr);

            mgr.When(x =>
            {
                x.DoesParentWshExist().Returns(true);
                x.DoesListObjExist().Returns(false);
                var ex = Assert.Catch<MissingListObjectException>(() =>
                    x.SetListObjAndParentWshPpts());
            });
        }

        [TestCase("foo", "bar")]
        public void SetListObjAndParentWshPpt_GoodCtrParam_SetsVerifiedFlagTrue(string wshNm, string tblNm)
        {
            WshListobjPair pr = new WshListobjPair(tblNm, wshNm);
            var mgr = Substitute.ForPartsOf<GenericListObjMgr>(pr);

            mgr.When(x =>
            {
                x.DoesParentWshExist().Returns(true);
                x.DoesListObjExist().Returns(false);
                x.SetListObjAndParentWshPpts();
                Assert.True(x.UnderlyingWshAndListObjVerified);
            });
        }
    }
}
