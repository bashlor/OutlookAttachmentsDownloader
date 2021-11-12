using System;
using System.Runtime.Serialization;

namespace OutlookAttachmentsDownloader.Exceptions
{
    public class AbortedOperationException: Exception
    {
        public AbortedOperationException()
        {
        }

        public AbortedOperationException(string message)
            : base(message)
        {
        }

        public AbortedOperationException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
    public class OAppException : Exception
    {
        public OAppException()
        {
        }

        public OAppException(string message)
            : base(message)
        {
        }

        public OAppException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
