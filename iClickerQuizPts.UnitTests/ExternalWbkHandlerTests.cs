using System;

using iClickerQuizPts;
using iClickerQuizPts.AppExceptions;
using NUnit.Framework;

namespace iClickerQuizPts.UnitTests
{
    [TestFixture]
    public class ExternalWbkHandlerTests
    {
        [Category("ExternalWbkHandler")]
        [TestCase("Session 40 Total 5/2/16 [2.00]", "5/2/16")]
        [TestCase("Session 40 Total 05/02/16 [2.00]", "5/2/16")]
        [TestCase("Session 40 Total 05/02/2016 [2.00]", "5/2/16")]
        public void GetDatePortionOfHeader_ValidHdrs_ReturnsCorrectDate(string hdr, string dt)
        {
            ExternalWbkHandler ewh = new ExternalWbkHandler();
            DateTime dtHdr = ewh.GetDatePortionOfHeader(hdr);
            Assert.AreEqual(dtHdr, DateTime.Parse(dt));
        }
    }
}
