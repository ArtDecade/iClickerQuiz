using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;

namespace iClickerQuizPts
{
    public static class MsgBoxGenerator
    {
        private static string _caption = string.Empty;
        private static string _msg = string.Empty;

        private static void ResetClassFields()
        {
            _caption = string.Empty;
            _msg = string.Empty;
        }

        public static void ShowMsg(MessageBoxButtons btns)
        {
            MessageBox.Show(_msg, _caption, btns);
            ResetClassFields();
        }

        public static void SetMissingNamedRngMsg(string excptnMsg)
        {
            _caption = "This Workbook Has Been Altered";

            // Build msg...
            const string S1 = 
                "This program has encountered the following serious error:";
            const string S2 =
                "You will not be able to continue until this workbook has been repaired.";
            _msg = string.Format(S1 + "\n\n\t" + excptnMsg + "\n\n" + S2);
        }
    }
}
