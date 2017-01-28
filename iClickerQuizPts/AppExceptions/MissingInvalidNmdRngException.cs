using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPts;
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
    public class MissingInvalidNmdRngException : ApplicationException
    {
        private RangeScope _scopeEnum = RangeScope.NotSpecified;
        private string _nmdRng;
        private string _pWsh;

        /// <summary>
        /// Gets an <see langword="enum"/> which indicating the scope of the 
        /// named range.
        /// </summary>
        public RangeScope NamedRangeScope
        {
            get
            { return _scopeEnum; }
        }
        
        /// <summary>
        /// Gets the name of the named range.
        /// </summary>
        public string RangeName
        {
            get
            { return _nmdRng; }
        }

        /// <summary>
        /// Gets the name of the parent worksheet for locally-scoped named ranges.
        /// </summary>
        public string ParentWsh
        {
            get
            { return _pWsh; }
        }


        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        public MissingInvalidNmdRngException() { }
        
        /// <summary>
        /// Initiales a new instance of the exception.
        /// </summary>
        /// <param name="scope">An enum indicating the scope of the missing named range.</param>
        /// <param name="rngNm">The name of the missing named range.</param>
        /// <param name="parentWsh">For a locally-scoped range the name of the parent worksheet.</param>
        public MissingInvalidNmdRngException(RangeScope scope, string rngNm, string parentWsh = "")
        {
            _scopeEnum = scope;
            _nmdRng = rngNm;
            _pWsh = parentWsh;
        }

        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        public MissingInvalidNmdRngException(string message) : base(message) { }
        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="message">A message about this exception.</param>
        /// <param name="inner">The exception which caused this exception.</param>
        public MissingInvalidNmdRngException(string message, Exception inner) : base(message, inner) { }
        /// <summary>
        /// Initializes a new instance of the exception.
        /// </summary>
        /// <param name="info">The data needed to serialize or deserialize this exception.</param>
        /// <param name="context">The source and destination of a the stream used
        /// to serialize this exception.</param>
        protected MissingInvalidNmdRngException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
