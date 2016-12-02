using NUnit.Framework;
using iClickerQuizPts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iClickerQuizPts.Tests
{
    [TestFixture()]
    public class ExternalWbkHandlerTests
    {
        //[Test()]
        //public void GetDatePortionOfHeader_ValidHdr_ReturnsDateTimeType()
        //{
        //    DateTime dtHdr =
        //        ExternalWbkHandler.GetDatePortionOfHeader(
        //            "Session 40 Total 5/2/16 [2.00]");
        //}

        //[Test()]
        //public void GetDatePortionofHeader_InvalidHdr_ThrowsEx()
        //{
        //    Assert.Fail();
        //}

        [Test()]
        public void GetDatePortionOfHeader_ValidHdr_ReturnsCorrectDate()
        {
            DateTime dtHdr =
                ExternalWbkHandler.GetDatePortionOfHeader(
                    "Session 40 Total 5/2/16 [2.00]");
            Assert.AreEqual(dtHdr, DateTime.Parse("5/2/16"));
        }
    }
}