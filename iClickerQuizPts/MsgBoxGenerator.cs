using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides a centralized location for generating all <see cref="MessageBox"/> pop-ups throughout the application.
    /// </summary>
    public static class MsgBoxGenerator
    {
        private static string _caption = string.Empty;
        private static string _msg = string.Empty;
        private const string CANNOT_CONTINUE = 
            "You will not be able to continue until this workbook has been repaired.";

        private static void ResetClassFields()
        {
            _caption = string.Empty;
            _msg = string.Empty;
        }

        /// <summary>
        /// The method which should always be called after any of the Set...Msg methods are invoked.
        /// </summary>
        /// <param name="btns"></param>
        public static void ShowMsg(MessageBoxButtons btns)
        {
            MessageBox.Show(_msg, _caption, btns);
            ResetClassFields();
        }

        /// <summary>
        /// Sets the caption and builds the message that will be presented to the user whenever 
        /// a <see cref="iClickerQuizPts.AppExceptions.MissingNamedRangeException"/> is thrown.
        /// </summary>
        /// <param name="excptnMsg">The message field from the <see cref="iClickerQuizPts.AppExceptions.MissingNamedRangeException"/>.</param>
        public static void SetMissingNamedRngMsg(string excptnMsg)
        {
            _caption = "This Workbook Has Been Altered";

            // Build msg...
            const string S1 = 
                "This program has encountered the following serious error:";
            _msg = string.Format(S1 + "\n\n\t" + excptnMsg + "\n\n" + CANNOT_CONTINUE);
        }

        /// <summary>
        /// Sets the caption and builds the message that will be presented to the user whenever 
        /// a <see cref="iClickerQuizPts.AppExceptions.MissingListObjectException"/> is thrown.
        /// </summary>
        /// <param name="pr">The <see langword="struc"/> which contains the name of the missisng list object and the name of the parent worksheet.</param>
        public static void SetMissingListObjMsg(WshListobjPair pr)
        {
            _caption = "This Workbook Has Been Altered";

            // Build msg...
            const string S1 =
                "We cannot find at least one of the ListObjects (Tables) required to run this application. ";
            _msg = string.Format(S1 + "\n\n\t" + "Missing ListObject(Table):\n\t\t" + pr.ListObjName +
                "\n\n\tWorksheet:\n\t\t" + pr.WshNm + "\n\n" + CANNOT_CONTINUE);
        }
    }
}
