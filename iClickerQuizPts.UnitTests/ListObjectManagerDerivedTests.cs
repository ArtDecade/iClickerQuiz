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
    [TestFixture]
    public class ListObjectManagerDerivedTests
    {
        const string QZ_GRADES_TBL = "tblClkrQuizGrades";
        const string QZ_GRADES_WSH = "iCLICKERQuizPoints";
        const string DBL_DPPRS_TBL = "tblDblDippers";
        const string DBL_DPPRS_WSH = "DoubleDippers";

        private ListObjectManager loMgr;

        [Category("ListObjectManagerTests")]
        [TestCase(QZ_GRADES_WSH, QZ_GRADES_TBL)]
        public void InstantiateQuizDataListObjectMgr_ValidWshTblNmz_Succeeds(string wshNm, string tblNm)
        {
            // Arrange...
            WshListobjPair pr = new WshListobjPair(tblNm, wshNm);
            var qdMgr = Substitute.ForPartsOf<QuizDataListObjMgr>(pr);
            qdMgr.DoesParentWshExist().Returns(true);
            qdMgr.DoesListObjExist().Returns(true);
            //var loMgr = Substitute.ForPartsOf<ListObjectManager>(pr);
            //loMgr.DoesParentWshExist().Returns(true);
            //loMgr.DoesListObjExist().Returns(true);



            // Act...
            //loMgr = Substitute.For<ListObjectManager>(pr);
            qdMgr = Substitute.For<QuizDataListObjMgr>(pr);


            // Assert...
            // Assert.IsInstanceOf<ListObjectManager>(loMgr);
            Assert.IsInstanceOf<QuizDataListObjMgr>(qdMgr);
            
        }


    }
}
