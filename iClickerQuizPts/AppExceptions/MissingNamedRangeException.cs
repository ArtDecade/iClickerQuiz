using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iClickerQuizPts.AppExceptions
{
    [Serializable]
    public class MissingNamedRangeException : ApplicationException
    {
        public MissingNamedRangeException() { }
        public MissingNamedRangeException(string message) : base(message) { }
        public MissingNamedRangeException(string message, Exception inner) : base(message, inner) { }
        protected MissingNamedRangeException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
