using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iClickerQuizPts.AppExceptions
{
    /// <summary>
    /// Represents the <see cref="System.ApplicationException"/>-derived exception that
    /// is thrown whenever there are issues reading data from a file using EPPlus.
    /// </summary>

    [Serializable]
    public class ReadingExternalWbkException : ApplicationException
    {
        /// <summary>
        /// Gets  or sets the specified <see cref="iClickerQuizPts.ImportResult"/> 
        /// enumeration which caused this exception to be thrown.
        /// </summary>
        public ImportResult ImportResult { get; set; }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// /// </summary>
        public ReadingExternalWbkException() { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        public ReadingExternalWbkException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        /// <param name="inner">The exception which caused this exception.</param>
        public ReadingExternalWbkException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="info">The data needed to serialize or deserialize this exception.</param>
        /// <param name="context">The source and destination of a the stream used
        /// to serialize this exception.</param>
        protected ReadingExternalWbkException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
   
}
