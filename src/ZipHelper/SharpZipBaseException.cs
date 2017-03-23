using System;

namespace ZipHelper
{
    public class SharpZipBaseException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the SharpZipBaseException class.
        /// </summary>
        public SharpZipBaseException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the SharpZipBaseException class with a specified error message.
        /// </summary>
        /// <param name="message">A message describing the exception.</param>
        public SharpZipBaseException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the SharpZipBaseException class with a specified
        /// error message and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">A message describing the exception.</param>
        /// <param name="innerException">The inner exception</param>
        public SharpZipBaseException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}