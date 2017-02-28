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
        private const string SESS_37 = "Session 37 Total 4/25/16 [2.00]";
        private const string SESS_02 = "Session 2 Total 1/25/16 [2.00]";
        private const string SESS_17 = "Session 17 Total 3/3/16 [2.00]";
        private const string DATE_37 = "4/25/16";
        private const string DATE_02 = "1/25/16";
        private const string DATE_17 = "3/3/16";

        [TestCase(SESS_37)]
        [TestCase(SESS_02)]
        [TestCase(SESS_17)]
        public void FileHeaderCtor_ValidFileHeader_Succeeds(string fHdr)
        {
            Session s;

            s = new Session(fHdr);

            Assert.IsInstanceOf<Session>(s);
        }

        [TestCase(SESS_37,"37",DATE_37,"2")]
        [TestCase(SESS_02,"02",DATE_02,"2")]
        [TestCase(SESS_17,"17",DATE_17,"2")]
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

      
        [TestCase(SESS_37, "37", DATE_37, 2)]
        [TestCase(SESS_02, "2", "1/1/17", 2)]
        [TestCase(SESS_17, "17", "1/1/17", 1)]
        public void SessionEquals_EqualSessNos_True(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsTrue(Session.Equals(s1, s2));
        }

        [TestCase(SESS_37, "37", DATE_37, 2)]
        [TestCase(SESS_02, "2", "1/1/17", 2)]
        [TestCase(SESS_17, "17", "1/1/17", 1)]
        public void SessionEqlSign_EqualSessNos_True(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsTrue(s1 == s2);
        }

        [TestCase(SESS_37, "99", DATE_37, 2)]
        [TestCase(SESS_02, "98", DATE_02, 2)]
        [TestCase(SESS_17, "97", DATE_17, 2)]
        public void SessionEquals_UnequalSessNos_False(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsFalse(Session.Equals(s1, s2));
        }

        [TestCase(SESS_37, "99", DATE_37, 2)]
        [TestCase(SESS_02, "98", DATE_02, 2)]
        [TestCase(SESS_17, "97", DATE_17, 2)]
        public void SessionEqlSign_UnequalSessNos_False(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsFalse(s1 == s2);
        }

        [TestCase(SESS_37, "37", DATE_37, 2)]
        [TestCase(SESS_02, "2", "1/1/17", 2)]
        [TestCase(SESS_17, "17", "1/1/17", 1)]
        public void SessionUneqlSign_EqualSessNos_False(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsFalse(s1 != s2);
        }

        [TestCase(SESS_37, "99", DATE_37, 2)]
        [TestCase(SESS_02, "98", DATE_02, 2)]
        [TestCase(SESS_17, "97", DATE_17, 2)]
        public void SessionUnqlSign_UnequalSessNos_True(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsTrue(s1 != s2);
        }

        [TestCase(SESS_37, "30", DATE_37, 2)]
        [TestCase(SESS_02, "1", DATE_02, 2)]
        [TestCase(SESS_17, "16", DATE_17, 2)]
        public void SessionGTSign_GTSessNos_True(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsTrue(s1 > s2);
        }

        [TestCase(SESS_37, "37", DATE_37, 2)]
        [TestCase(SESS_02, "2", DATE_02, 2)]
        [TestCase(SESS_17, "16", DATE_17, 2)]
        public void SessionGTESign_GTESessNos_True(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsTrue(s1 >= s2);
        }

        [TestCase(SESS_37, "40", DATE_37, 2)]
        [TestCase(SESS_02, "2", DATE_02, 2)]
        [TestCase(SESS_17, "17", DATE_17, 2)]
        public void SessionGTSign_NotGTSessNos_False(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsFalse(s1 > s2);
        }

        [TestCase(SESS_37, "40", DATE_37, 2)]
        [TestCase(SESS_02, "3", DATE_02, 2)]
        [TestCase(SESS_17, "99", DATE_17, 2)]
        public void SessionGTESign_NotGTESessNos_False(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsFalse(s1 >= s2);
        }

        [TestCase(SESS_37, "38", DATE_37, 2)]
        [TestCase(SESS_02, "20", DATE_02, 2)]
        [TestCase(SESS_17, "20", DATE_17, 2)]
        public void SessionLTSign_LTSessNos_True(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsTrue(s1 < s2);
        }

        [TestCase(SESS_37, "37", DATE_37, 2)]
        [TestCase(SESS_02, "2", DATE_02, 2)]
        [TestCase(SESS_17, "99", DATE_17, 2)]
        public void SessionLTESign_LTESessNos_True(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsTrue(s1 <= s2);
        }

        [TestCase(SESS_37, "36", DATE_37, 2)]
        [TestCase(SESS_02, "2", DATE_02, 2)]
        [TestCase(SESS_17, "17", DATE_17, 2)]
        public void SessionLTSign_NotLTSessNos_False(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsFalse(s1 < s2);
        }

        [TestCase(SESS_37, "1", DATE_37, 2)]
        [TestCase(SESS_02, "1", DATE_02, 2)]
        [TestCase(SESS_17, "1", DATE_17, 2)]
        public void SessionLTESign_NotLTESessNos_False(string fHdr, string sessNo, DateTime dt, byte maxPts)
        {
            Session s1 = new Session(fHdr);
            Session s2 = new Session(sessNo, dt, maxPts);

            Assert.IsFalse(s1 <= s2);
        }
    }
}
