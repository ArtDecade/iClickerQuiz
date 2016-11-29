using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iClickerQuizPts
{
    public static class UserControlsHandler
    {
        public enum DatesToShow : byte
        {
            NoSelection,
            AllDates,
            NewDatesOnly
        }

        public static DatesToShow RadBtnDates { get; private set; } = DatesToShow.NoSelection;
        public static WkSession WhichSession { get; set; } = WkSession.None;
        public static byte CourseWeek { get; set; }
        public static DateTime QuizDate { get; set; } = DateTime.Parse("1/1/2016");

        public static void SetCourseWeek(string selectedWk)
        {
            CourseWeek = byte.Parse(selectedWk);
        }

        public static void SetSessionEnum(string session)
        {
            switch (session)
            {
                case "First":
                    WhichSession = WkSession.First;
                    break;
                case "Second":
                    WhichSession = WkSession.Second;
                    break;
                case "Third":
                    WhichSession = WkSession.Third;
                    break;
                default:
                    WhichSession = WkSession.None;
                    break;
            }
        }

        public static void SetDatesToShowEnum(string btnNm)
        {
            switch (btnNm)
            {
                case "radAllDates":
                    RadBtnDates = DatesToShow.AllDates;
                    break;
                case "radNewDatesOnly":
                    RadBtnDates = DatesToShow.NewDatesOnly;
                    break;
                default:
                    RadBtnDates = DatesToShow.NoSelection;
                    break;
            }
        }

        public static void PopulateQuizDatesComboBox()
        {
            
        }

        public static void ImportDataMaestro()
        {
            
        }
    }
}
