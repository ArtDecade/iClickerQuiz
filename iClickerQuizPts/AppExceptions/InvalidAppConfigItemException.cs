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
    /// is thrown whenever we cannot find a given key in the <code>appSettings</code> section 
    /// of the <code>App.config</code> file.
    /// </summary>
    [Serializable]
    public class InalidAppConfigItemException : ApplicationException
    {
        /// <summary>
        /// Gets or sets the key we cannot find in the <code>App.config</code> file.
        /// </summary>
        public string MissingKey { get; set; }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// /// </summary>
        public InalidAppConfigItemException() { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        public InalidAppConfigItemException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        /// <param name="inner">The exception which caused this exception.</param>
        public InalidAppConfigItemException(string message, Exception inner) : base(message, inner) { }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="info">The data needed to serialize or deserialize this exception.</param>
        /// <param name="context">The source and destination of a the stream used
        /// to serialize this exception.</param>
        protected InalidAppConfigItemException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}