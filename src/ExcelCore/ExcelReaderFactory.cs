using System;
using System.IO;

namespace ExcelCore
{
    public static class ExcelReaderFactory
    {

        /// <summary>
        /// Readers the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <returns>IExcelData Reader</returns>
        /// <exception cref="System.Exception"></exception>
        public static IExcelDataReader Reader(string filePath)
        {
            IExcelDataReader reader;
            FileInfo fileInfo = new FileInfo(filePath);
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (fileInfo.Extension == ".xls")
            {
                reader = new ExcelBinaryReader();
                reader.Initialize(fileStream);                
            }
            else if(fileInfo.Extension == ".xlsx")
            {
                reader = new ExcelOpenXmlReader();
                reader.Initialize(fileStream);
            }          
            else
            {
                throw new Exception(string.Format("File Extension '{0}' invalid, Please check file", fileInfo.Extension));
            }
            fileStream.Dispose();
            return reader;
        }        
    }
}