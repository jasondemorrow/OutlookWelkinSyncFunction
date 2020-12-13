namespace OutlookWelkinSync
{
    using System;
    using System.Runtime.Serialization;

    public class LinkException : Exception
    {
        public LinkException()
        {
        }

        public LinkException(string message) : base(message)
        {
        }

        public LinkException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected LinkException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}