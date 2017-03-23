using System.Collections;
using System.Collections.Generic;
using System.IO;
using ExcelReader.Model;

namespace ExcelReader
{
    public interface IExcelDataReader
    {
        /// <summary>
        /// Initializes the instance with specified file stream.
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        void Initialize(Stream fileStream);

        /// <summary>
        /// Read all data in to DataSet and return it
        /// </summary>
        /// <returns>The DataSet</returns>
        IEnumerable AsIEnumerable();

        IEnumerable AsIEnumerable(List<ExcelHeaderKeyValues> headerValues);

        List<T> AsIEnumerable<T>(List<ExcelHeaderKeyValues> headerValues);

        string AsJson();

        string AsJson(List<ExcelHeaderKeyValues> headerValues);

        /// <summary>
        /// Gets a value indicating whether file stream is valid.
        /// </summary>
        /// <value><c>true</c> if file stream is valid; otherwise, <c>false</c>.</value>
        bool IsValid { get; }

        /// <summary>
        /// Gets the exception message in case of error.
        /// </summary>
        /// <value>The exception message.</value>
        string ExceptionMessage { get; }

        /// <summary>
        /// Gets the sheet name.
        /// </summary>
        /// <value>The sheet name.</value>
        string Name { get; }

        /// <summary>
        /// Gets the number of results (workbooks).
        /// </summary>
        /// <value>The results count.</value>
        int ResultsCount { get; }
        
    }
}