using System;
using System.Runtime.Serialization;

namespace ExcelMapper.Exceptions
{
    [Serializable]
    internal class ExcelMappingException : Exception
    {
        public ExcelMappingException()
        {
        }

        public ExcelMappingException(string message) : base(message)
        {
        }

        public ExcelMappingException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected ExcelMappingException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}