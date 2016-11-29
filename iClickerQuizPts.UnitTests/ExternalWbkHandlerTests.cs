using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace iClickerQuizPts.UnitTests
{
    [TestClass]
    public class ExternalWbkHandlerTests
    {
        [TestMethod]
        public void GetDatePortionOfHeader_ValidHdr_ReturnsCorrectDate()
        {
            DateTime dtHdr =
                ExternalWbkHandler.GetDatePortionOfHeader(
                    "Session 40 Total 5/2/16 [2.00]");
            Assert.AreEqual(dtHdr, DateTime.Parse("5/2/16"));
        }
    }
}
