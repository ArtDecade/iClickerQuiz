using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPts;
using iClickerQuizPts.AppExceptions;

using NUnit.Framework;
using NSubstitute;


namespace iClickerQuizPts.UnitTests
{
    [TestFixture]
    [Category("NamedReferenceMgrTests")]
    class NamedReferenceManagerTests
    {
        [TestCase("foo")]
        public void ConfirmWorkbookScopedRangeExists_RngMissing_Thows(string rngNm)
        {
            var mgr = Substitute.ForPartsOf<NamedRangeManager>();
            mgr.When(x => x.WorkbookScopedRangeExists(rngNm)).DoNotCallBase();
            mgr.WorkbookScopedRangeExists(rngNm).Returns(false);

            var ex = Assert.Catch<MissingInvalidNmdRngException>(() =>
                mgr.ConfirmWorkbookScopedRangeExists(rngNm));
        }

        [TestCase("foo","bar")]
        public void ConfirmWorksheetScopedRangeExists_RngMissing_Throws(string wsNm, string rngNm)
        {
            var mgr = Substitute.ForPartsOf<NamedRangeManager>();
            mgr.When(x => x.WorksheetScopedRangeExists(wsNm, rngNm)).DoNotCallBase();
            mgr.WorksheetScopedRangeExists(wsNm, rngNm).Returns(false);

            var ex = Assert.Catch<MissingInvalidNmdRngException>(() =>
                mgr.ConfirmWorksheetScopedRangeExists(wsNm, rngNm));
        }
    }
}
