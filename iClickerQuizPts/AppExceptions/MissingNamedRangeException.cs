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
    /// is thrown whenever a named <see cref="Excel.Range"/> cannot be found.
    /// </summary>
    /// <remarks>
    /// At design time this workbook was built to include a number of named ranges.
    /// This application will throw this exception if the user has managed to delete 
    /// (or to rename) any of these ranges.
    /// </remarks>
    [Serializable]
    public class MissingNamedRangeException : ApplicationException
    {



        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        public MissingNamedRangeException() { }
        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        public MissingNamedRangeException(string message) : base(message) { }
        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        /// <param name="inner">The exception which caused this exception.</param>
        public MissingNamedRangeException(string message, Exception inner) : base(message, inner) { }
        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="info">The data needed to serialize or deserialize this exception.</param>
        /// <param name="context">The source and destination of a the stream used
        /// to serialize this exception.</param>
        protected MissingNamedRangeException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
