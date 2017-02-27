using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPts;
using iClickerQuizPts.AppExceptions;
using NUnit.Framework;
using NSubstitute;
using System.Windows.Forms;

namespace iClickerQuizPts.UnitTests
{
    [TestFixture]
    [Category("SessionTests")]
    public class SessionTests
    {
        [TestCase("Session 37 Total 4/25/16 [2.00]")]
        [TestCase("Session 2 Total 1/25/16 [2.00]")]
        [TestCase("Session 17 Total 3/3/16 [2.00]")]
        public void FileHeaderCtor_ValidFileHeader_Succeeds(string fHdr)
        {
            Session s;

            try
            {
                s = new Session(fHdr);

                Assert.IsInstanceOf<Session>(s);
            }
            catch(InvalidQuizDataHeaderException ex)
            {
                MsgBoxGenerator.SetInvalidHdrMsg(fHdr);
                MsgBoxGenerator.ShowMsg(MessageBoxButtons.OK);
            }
        }

        [TestCase("Session 37 Total 4/25/16 [2.00]","37","4/25/16","2")]
        [TestCase("Session 2 Total 1/25/16 [2.00]","02","1/25/16","2")]
        [TestCase("Session 17 Total 3/3/16 [2.00]","17","3/3/16","2")]
        public void FileHeaderCtor_ValidFileHeader_PpptsPopulated(string fHdr, string sNo, string dt, string pts)
        {
            Session s;

            s = new Session(fHdr);

            Assert.AreEqual(s.SessionNo, sNo);
            Assert.AreEqual(s.QuizDate, DateTime.Parse(dt));
            Assert.AreEqual(s.MaxPts, byte.Parse(pts));
        }

        [TestCase("foo")]
        public void FileHeaderCtor_InvalidFileHeader_Throws(string fHdr)
        {
            Session s;

            var ex = Assert.Catch<InvalidQuizDataHeaderException>(() =>
                s = new Session(fHdr));
        }

        [TestCase("7", "2/24/17", 2,"07","2/24/17",2)]
        public void ThreeParamCtor_ValidParams_PptsPopulated(string sNo, 
            DateTime dt, byte maxPts,string sNoPpty, string dtPpty, byte maxPpty)
        {
            Session s;

            s = new Session(sNo, dt, maxPts);

            Assert.AreEqual(s.SessionNo, sNoPpty);
            Assert.AreEqual(s.QuizDate, DateTime.Parse(dtPpty));
            Assert.AreEqual(s.MaxPts, maxPpty);
        }

    }
}
