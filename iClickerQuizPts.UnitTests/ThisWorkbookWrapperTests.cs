using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NUnit.Framework;
using NSubstitute;
using iClickerQuizPts.AppExceptions;
using iClickerQuizPts.ListObjMgmt;
using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts.UnitTests
{
    [TestFixture]
    [Category("ThisWorkbookWrapperTests")]
    class ThisWorkbookWrapperTests
    {
        [TestCase("foo")]
        public void VerifyWbkScopedNames_InvalidNames_Throws(string nm)
        {
            var wbw = Substitute.For<ThisWorkbookWrapper>();

            
        }
    }
}
