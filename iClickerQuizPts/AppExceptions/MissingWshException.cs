using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts.AppExceptions
{
    /// <summary>
    /// Represents the <see cref="System.ApplicationException"/>-derived exception that
    /// is thrown whenever a <see cref="Excel.Worksheet"/> cannot be found.
    /// </summary>
    /// <remarks>
    /// This application will throw this exception if the user has deleted (or renamed) any
    /// of the original worksheets.
    /// <remarks>
    [Serializable]
    public class MissingWorksheetException : ApplicationException
    {
        /// <summary>
        /// A <see langword="struc"/> which contains: 
        /// <list type="bullet">
        /// <item>
        /// <description>The <c>string</c> name of a worksheet in this workbook.</description>
        /// </item>
        /// <item>
        /// <description>The name of the worksheet's <see cref="Excel.ListObject"/>.</description>
        /// </item>
        /// </list>
        /// </summary>
        public WshListobjPair WshListObjPair { get; set; }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        public MissingWorksheetException() { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        public MissingWorksheetException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        /// <param name="inner">The exception which caused this exception.</param>
        public MissingWorksheetException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="info">The data needed to serialize or deserialize this exception.</param>
        /// <param name="context">The source and destination of a the stream used
        /// to serialize this exception.</param>
        protected MissingWorksheetException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
