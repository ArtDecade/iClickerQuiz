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
    /// is thrown whenever an invalid <see cref="iClickerQuizPts.WshListobjPair"/> 
    /// instance is utilized in the code.
    /// </summary>
    /// <remarks>
    /// Every <see cref="iClickerQuizPts.WshListobjPair"/> must have both its 
    /// <see cref="iClickerQuizPts.WshListobjPair.ListObjName"/> and its
    /// <see cref="iClickerQuizPts.WshListobjPair.WshNm"/> properties populated.
    /// </remarks>
    [Serializable]
    public class InvalidWshListObjPairException : ApplicationException
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
        /// /// </summary>
        public InvalidWshListObjPairException() { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        public InvalidWshListObjPairException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        /// <param name="inner">The exception which caused this exception.</param>
        public InvalidWshListObjPairException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="info">The data needed to serialize or deserialize this exception.</param>
        /// <param name="context">The source and destination of a the stream used
        /// to serialize this exception.</param>
        protected InvalidWshListObjPairException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
