using System;

namespace ExcelCore.Exceptions
{
    public class BiffRecordException : Exception
    {
        public BiffRecordException()
        {
        }

        public BiffRecordException(string message)
            : base(message)
        {
        }

        public BiffRecordException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}