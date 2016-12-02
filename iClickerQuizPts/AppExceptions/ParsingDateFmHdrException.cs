using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iClickerQuizPts.AppExceptions
{
    [Serializable]
    public class ParsingDateFmHdrException : ApplicationException
    {
        public ParsingDateFmHdrException() { }
        public ParsingDateFmHdrException(string message) : base(message) { }
        public ParsingDateFmHdrException(string message, Exception inner) : base(message, inner) { }
        protected ParsingDateFmHdrException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
        public string HeaderText { get; set; }
    }
}
