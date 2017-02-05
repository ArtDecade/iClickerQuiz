using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPts;

namespace iClickerQuizPts.AppExceptions
{
    /// <summary>
    /// Represents the <see cref="System.ApplicationException"/>-derived exception that
    /// is thrown whenever this application is unable to parse a date from an 
    /// external workbook containing raw iClicker quiz data.
    /// </summary>
    /// <remarks>Raw iClicker quiz data files have column headers in a very specific 
    /// format.  Those headers (regrettably) contain several pieces of information, 
    /// inclucing the date on which an iClicker quiz was administered.  
    /// <para>This application has been built to extract the quiz dates from those
    /// column headers.  If this exception is thrown by the application it essentially 
    /// means that the column headers have been reformatted.  It will also mean that, 
    /// at the very least, the <see cref="ExternalWbkWrapper.GetDatePortionOfHeader(string)"/> 
    /// method will need to be refactored.
    /// </para></remarks>
    [Serializable]
    public class InvalidQuizDataHeaderException : ApplicationException
    {
        /// <summary>
        /// Gets or sets the text of the header cell causing this exception.
        /// </summary>
        public string HeaderText { get; set; }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        public InvalidQuizDataHeaderException() { }
        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="msg">A message about this exception.</param>
        public InvalidQuizDataHeaderException(string msg) : base(msg) { }
        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">>A message about this exception.</param>
        /// <param name="inner">The exception which caused this exception.</param>
        public InvalidQuizDataHeaderException(string message, Exception inner) : base(message, inner) { }
        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="info">The data needed to serialize or deserialize this exception.</param>
        /// <param name="context">The source and destination of a the stream used
        /// to serialize this exception.</param>
        protected InvalidQuizDataHeaderException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }


  
}
