using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;
using ExcelReader.Core;
using ExcelReader.Core.OpenXmlFormat;
using ExcelReader.Model;
using Newtonsoft.Json;

namespace ExcelReader
{
    public class ExcelOpenXmlReader : IExcelDataReader
    {
        #region Members

        private XlsxWorkbook _workbook;
        private bool _isValid;
        private bool _isClosed;
        private bool _isFirstRead;
        private string _exceptionMessage;
        private int _depth;
        private int _resultIndex;
        private int _emptyRowCount;
        private ZipWorker _zipWorker;
        private XmlReader _xmlReader;
        private Stream _sheetStream;
        private object[] _cellsValues;
        private object[] _savedCellsValues;

        private bool disposed;
        private bool _isFirstRowAsColumnNames;
        private const string COLUMN = "Column";
        private string instanceId = Guid.NewGuid().ToString();

        private List<int> _defaultDateTimeStyles;

        #endregion Members

        internal ExcelOpenXmlReader()
        {
            _isValid = true;
            _isFirstRead = true;

            _defaultDateTimeStyles = new List<int>(new int[]
            {
                14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
            });
        }

        private void ReadGlobals()
        {
            _workbook = new XlsxWorkbook(
                _zipWorker.GetWorkbookStream(),
                _zipWorker.GetWorkbookRelsStream(),
                _zipWorker.GetSharedStringsStream(),
                _zipWorker.GetStylesStream());

            CheckDateTimeNumFmts(_workbook.Styles.NumFmts);
        }

        private void CheckDateTimeNumFmts(List<XlsxNumFmt> list)
        {
            if (list.Count == 0) return;

            foreach (XlsxNumFmt numFmt in list)
            {
                if (string.IsNullOrEmpty(numFmt.FormatCode)) continue;
                string fc = numFmt.FormatCode.ToLower();

                int pos;
                while ((pos = fc.IndexOf('"')) > 0)
                {
                    int endPos = fc.IndexOf('"', pos + 1);

                    if (endPos > 0) fc = fc.Remove(pos, endPos - pos + 1);
                }

                //it should only detect it as a date if it contains
                //dd mm mmm yy yyyy
                //h hh ss
                //AM PM
                //and only if these appear as "words" so either contained in [ ]
                //or delimted in someway
                //updated to not detect as date if format contains a #
                var formatReader = new FormatReader() { FormatString = fc };
                if (formatReader.IsDateFormatString())
                {
                    _defaultDateTimeStyles.Add(numFmt.Id);
                }
            }
        }

        private void ReadSheetGlobals(XlsxWorksheet sheet)
        {
            if (_xmlReader != null) _xmlReader.Dispose();
            if (_sheetStream != null) _sheetStream.Dispose();

            _sheetStream = _zipWorker.GetWorksheetStream(sheet.Path);

            if (null == _sheetStream) return;

            _xmlReader = XmlReader.Create(_sheetStream);

            while (_xmlReader.Read())
            {
                if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.Name == XlsxWorksheet.N_dimension)
                {
                    string dimValue = _xmlReader.GetAttribute(XlsxWorksheet.A_ref);
                    sheet.Dimension = new XlsxDimension(dimValue);
                    break;
                }
            }

            _xmlReader.ReadToFollowing(XlsxWorksheet.N_sheetData);
            if (_xmlReader.IsEmptyElement)
            {
                sheet.IsEmpty = true;
            }
        }

        private bool ReadSheetRow(XlsxWorksheet sheet)
        {
            if (null == _xmlReader) return false;

            if (_emptyRowCount != 0)
            {
                _cellsValues = new object[sheet.ColumnsCount];
                _emptyRowCount--;
                _depth++;

                return true;
            }

            if (_savedCellsValues != null)
            {
                _cellsValues = _savedCellsValues;
                _savedCellsValues = null;
                _depth++;

                return true;
            }

            if ((_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.Name == XlsxWorksheet.N_row) ||
                _xmlReader.ReadToFollowing(XlsxWorksheet.N_row))
            {
                _cellsValues = new object[sheet.ColumnsCount];

                int rowIndex = int.Parse(_xmlReader.GetAttribute(XlsxWorksheet.A_r));
                if (rowIndex != (_depth + 1))
                {
                    _emptyRowCount = rowIndex - _depth - 1;
                }
                bool hasValue = false;
                string a_s = String.Empty;
                string a_t = String.Empty;
                string a_r = String.Empty;
                int col = 0;
                int row = 0;

                while (_xmlReader.Read())
                {
                    if (_xmlReader.Depth == 2) break;

                    if (_xmlReader.NodeType == XmlNodeType.Element)
                    {
                        hasValue = false;

                        if (_xmlReader.Name == XlsxWorksheet.N_c)
                        {
                            a_s = _xmlReader.GetAttribute(XlsxWorksheet.A_s);
                            a_t = _xmlReader.GetAttribute(XlsxWorksheet.A_t);
                            a_r = _xmlReader.GetAttribute(XlsxWorksheet.A_r);
                            XlsxDimension.XlsxDim(a_r, out col, out row);
                        }
                        else if (_xmlReader.Name == XlsxWorksheet.N_v)
                        {
                            hasValue = true;
                        }
                    }

                    if (_xmlReader.NodeType == XmlNodeType.Text && hasValue)
                    {
                        long number;
                        object o = _xmlReader.Value;

                        if (long.TryParse(o.ToString(), out number))
                            o = number;

                        if (null != a_t && a_t == XlsxWorksheet.A_s) //if string
                        {
                            o = Helpers.ConvertEscapeChars(_workbook.SST[int.Parse(o.ToString())]);
                        }
                        else if (a_t == "b") //boolean
                        {
                            o = _xmlReader.Value == "1";
                        }
                        else if (null != a_s) //if something else
                        {
                            XlsxXf xf = _workbook.Styles.CellXfs[int.Parse(a_s)];
                            if (xf.ApplyNumberFormat && o != null && o.ToString() != string.Empty && IsDateTimeStyle(xf.NumFmtId))
                                o = Helpers.ConvertFromOATime(number);
                            else if (xf.NumFmtId == 49)
                                o = o.ToString();
                        }

                        if (col - 1 < _cellsValues.Length)
                            _cellsValues[col - 1] = o;
                    }
                }

                if (_emptyRowCount > 0)
                {
                    _savedCellsValues = _cellsValues;
                    return ReadSheetRow(sheet);
                }
                _depth++;

                return true;
            }

            _xmlReader.Dispose();
            if (_sheetStream != null) _sheetStream.Dispose();

            return false;
        }

        private bool InitializeSheetRead()
        {
            if (ResultsCount <= 0) return false;

            ReadSheetGlobals(_workbook.Sheets[_resultIndex]);

            if (_workbook.Sheets[_resultIndex].Dimension == null) return false;

            _isFirstRead = false;

            _depth = 0;
            _emptyRowCount = 0;

            return true;
        }

        private bool IsDateTimeStyle(int styleId)
        {
            return _defaultDateTimeStyles.Contains(styleId);
        }

        #region IExcelDataReader Members

        public void Initialize(Stream fileStream)
        {
            _zipWorker = new ZipWorker();
            _zipWorker.Extract(fileStream);

            if (!_zipWorker.IsValid)
            {
                _isValid = false;
                _exceptionMessage = _zipWorker.ExceptionMessage;

                Close();

                return;
            }

            ReadGlobals();
        }

        public IEnumerable AsIEnumerable()
        {
            Dictionary<int, List<KeyValuePair<string, object>>> response = new Dictionary<int, List<KeyValuePair<string, object>>>();
            if (!_isValid) return null;

            // for (int ind = 0; ind < _workbook.Sheets.Count; ind++)
            for (int ind = 0; ind < 1; ind++)
            {
                ReadSheetGlobals(_workbook.Sheets[ind]);

                if (_workbook.Sheets[ind].Dimension == null) continue;
                List<string> ColArr = new List<string>();
                if (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    for (int index = 0; index < _cellsValues.Length; index++)
                    {
                        if (_cellsValues[index] != null)
                        {
                            ColArr.Add(_cellsValues[index].ToString());
                        }
                    }
                }
                else continue;
                int i = 0;
                while (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    response.Add(i, GetItem(ColArr, _cellsValues));
                    i++;
                }
            }

            return response.Select(x => x.Value).ToList();
        }

        public IEnumerable AsIEnumerable(List<ExcelHeaderKeyValues> headerValues)
        {
            Dictionary<int, List<KeyValuePair<string, object>>> response = new Dictionary<int, List<KeyValuePair<string, object>>>();
            if (!_isValid) return null;

            //for (int ind = 0; ind < _workbook.Sheets.Count; ind++)
            for (int ind = 0; ind < 1; ind++)
            {
                ReadSheetGlobals(_workbook.Sheets[ind]);

                if (_workbook.Sheets[ind].Dimension == null) continue;

                List<string> ColArr = new List<string>();
                bool InvalidColumn = false;
                if (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    if (headerValues.Count > 0)
                    {
                        for (int index = 0; index < _cellsValues.Length; index++)
                        {
                            if (_cellsValues[index] != null)
                            {
                                var headerTitle = headerValues.FirstOrDefault(x => x.HeaderTitle.ToLower().Equals(Convert.ToString(_cellsValues[index]).ToLower()));
                                if (headerTitle != null)
                                    ColArr.Add(headerTitle.HeaderKey);
                                else
                                    InvalidColumn = true;
                            }
                        }

                        if (InvalidColumn)
                            throw (new Exception("Column name not match ! Please correct column name.  Invalid column name '" + string.Join(",", ColArr.ToArray() + "'")));
                    }
                }
                else continue;
                int i = 0;
                while (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    response.Add(i, GetItem(ColArr, _cellsValues));
                    i++;
                }
            }

            return response.Select(x => x.Value).ToList();
        }

        public List<T> AsIEnumerable<T>(List<ExcelHeaderKeyValues> headerValues)
        {
            List<T> response = new List<T>();
            if (!_isValid) return null;

            //for (int ind = 0; ind < _workbook.Sheets.Count; ind++)
            for (int ind = 0; ind < 1; ind++)
            {
                ReadSheetGlobals(_workbook.Sheets[ind]);

                if (_workbook.Sheets[ind].Dimension == null) continue;

                List<string> ColArr = new List<string>();
                bool InvalidColumn = false;
                if (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    if (headerValues.Count > 0)
                    {
                        for (int index = 0; index < _cellsValues.Length; index++)
                        {
                            if (_cellsValues[index] != null)
                            {
                                var headerTitle = headerValues.FirstOrDefault(x => x.HeaderTitle.ToLower().Equals(Convert.ToString(_cellsValues[index]).ToLower()));
                                if (headerTitle != null)
                                    ColArr.Add(headerTitle.HeaderKey);
                                else
                                    InvalidColumn = true;
                            }
                        }

                        if (InvalidColumn)
                            throw (new Exception("Column name not match ! Please correct column name.  Invalid column name '" + string.Join(",", ColArr.ToArray() + "'")));
                    }
                }
                else continue;
                int i = 0;
                while (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    response.Add(GetItem<T>(ColArr, _cellsValues));
                    i++;
                }
            }

            return response;
        }

        public string AsJson(List<ExcelHeaderKeyValues> headerValues)
        {
            StringBuilder response = new StringBuilder();
            response.Append("[");
            if (!_isValid) return null;
            //for (int ind = 0; ind < 1; ind++)
            for (int ind = 0; ind < 1; ind++)
            {
                ReadSheetGlobals(_workbook.Sheets[ind]);
                if (_workbook.Sheets[ind].Dimension == null) continue;
                var ss = ReadSheetRow(_workbook.Sheets[ind]);
                List<string> ColArr = new List<string>();
                bool InvalidColumn = false;
                if (headerValues.Count > 0)
                {
                    for (int index = 0; index < _cellsValues.Length; index++)
                    {
                        if (_cellsValues[index] != null)
                        {
                            var headerTitle = headerValues.FirstOrDefault(x => x.HeaderTitle.ToLower().Equals(Convert.ToString(_cellsValues[index]).ToLower()));
                            if (headerTitle != null)
                                ColArr.Add(headerTitle.HeaderKey);
                            else
                                InvalidColumn = true;
                        }
                    }

                    if (InvalidColumn)
                        throw (new Exception("Column name not match ! Please correct column name.  Invalid column name '" + string.Join(",", ColArr.ToArray()) + "'"));
                }

                while (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    response.Append(AddRow(_cellsValues, ColArr));
                    response.Append(",");
                }
            }
            response.Remove(response.Length - 1, 1);
            response.Append("]");
            return response.ToString();
        }

        public string AsJson()
        {
            StringBuilder response = new StringBuilder();
            response.Append("[");
            if (!_isValid) return null;
            for (int ind = 0; ind < 1; ind++)
            {
                ReadSheetGlobals(_workbook.Sheets[ind]);

                if (_workbook.Sheets[ind].Dimension == null) continue;

                var ss = ReadSheetRow(_workbook.Sheets[ind]);
                List<string> ColArr = new List<string>();
                for (int index = 0; index < _cellsValues.Length; index++)
                {
                    if (_cellsValues[index] != null)
                        ColArr.Add((string)_cellsValues[index]);
                }

                while (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    response.Append(AddRow(_cellsValues, ColArr));
                    response.Append(",");
                }
            }
            response.Remove(response.Length - 1, 1);
            response.Append("]");
            return response.ToString();
        }

        private static List<KeyValuePair<string, object>> GetItem(List<string> keyName, object[] objectRow)
        {
            List<KeyValuePair<string, object>> response = new List<KeyValuePair<string, object>>();

            for (int i = 0; i < keyName.Count(); i++)
            {
                response.Add(new KeyValuePair<string, object>(keyName[i], objectRow[i]));
            }

            return response;
        }

        private static T GetItem<T>(List<string> keyName, object[] objectRow)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            for (int i = 0; i < keyName.Count; i++)
            {
                PropertyInfo pro = temp.GetProperties().FirstOrDefault(x => x.Name.ToLower().Contains(keyName[i].ToLower()));
                if (pro != null && pro.Name != null && objectRow[i] != null)
                {
                    pro.SetValue(obj, Convert.ChangeType(objectRow[i], pro.PropertyType, CultureInfo.InvariantCulture));
                }
                else
                {
                    continue;
                }
            }

            return obj;
        }

        private string AddRow(object[] m_cellsValues, List<string> titleValue)
        {
            Dictionary<string, object> responseVal = new Dictionary<string, object>();
            for (int i = 0; i < titleValue.Count; i++)
                responseVal.Add(titleValue[i], m_cellsValues[i]);//
            string json = JsonConvert.SerializeObject(responseVal, Formatting.Indented);
            return json;
        }

        public bool IsFirstRowAsColumnNames
        {
            get
            {
                return _isFirstRowAsColumnNames;
            }
            set
            {
                _isFirstRowAsColumnNames = value;
            }
        }

        public bool IsValid
        {
            get { return _isValid; }
        }

        public string ExceptionMessage
        {
            get { return _exceptionMessage; }
        }

        public string Name
        {
            get
            {
                return (_resultIndex >= 0 && _resultIndex < ResultsCount) ? _workbook.Sheets[_resultIndex].Name : null;
            }
        }

        public void Close()
        {
            _isClosed = true;

            if (_xmlReader != null) _xmlReader.Dispose();

            if (_sheetStream != null) _sheetStream.Dispose();

            if (_zipWorker != null) _zipWorker.Dispose();
        }

        public int Depth
        {
            get { return _depth; }
        }

        public int ResultsCount
        {
            get { return _workbook == null ? -1 : _workbook.Sheets.Count; }
        }

        public bool IsClosed
        {
            get { return _isClosed; }
        }

        public bool NextResult()
        {
            if (_resultIndex >= (this.ResultsCount - 1)) return false;

            _resultIndex++;

            _isFirstRead = true;

            return true;
        }

        public bool Read()
        {
            if (!_isValid) return false;

            if (_isFirstRead && !InitializeSheetRead())
            {
                return false;
            }

            return ReadSheetRow(_workbook.Sheets[_resultIndex]);
        }

        public int FieldCount
        {
            get { return (_resultIndex >= 0 && _resultIndex < ResultsCount) ? _workbook.Sheets[_resultIndex].ColumnsCount : -1; }
        }

        public bool GetBoolean(int i)
        {
            if (IsDBNull(i)) return false;

            return Boolean.Parse(_cellsValues[i].ToString());
        }

        public DateTime GetDateTime(int i)
        {
            if (IsDBNull(i)) return DateTime.MinValue;

            try
            {
                return (DateTime)_cellsValues[i];
            }
            catch (InvalidCastException)
            {
                return DateTime.MinValue;
            }
        }

        public decimal GetDecimal(int i)
        {
            if (IsDBNull(i)) return decimal.MinValue;

            return decimal.Parse(_cellsValues[i].ToString());
        }

        public double GetDouble(int i)
        {
            if (IsDBNull(i)) return double.MinValue;

            return double.Parse(_cellsValues[i].ToString());
        }

        public float GetFloat(int i)
        {
            if (IsDBNull(i)) return float.MinValue;

            return float.Parse(_cellsValues[i].ToString());
        }

        public short GetInt16(int i)
        {
            if (IsDBNull(i)) return short.MinValue;

            return short.Parse(_cellsValues[i].ToString());
        }

        public int GetInt32(int i)
        {
            if (IsDBNull(i)) return int.MinValue;

            return int.Parse(_cellsValues[i].ToString());
        }

        public long GetInt64(int i)
        {
            if (IsDBNull(i)) return long.MinValue;

            return long.Parse(_cellsValues[i].ToString());
        }

        public string GetString(int i)
        {
            if (IsDBNull(i)) return null;

            return _cellsValues[i].ToString();
        }

        public object GetValue(int i)
        {
            return _cellsValues[i];
        }

        public bool IsDBNull(int i)
        {
            return (null == _cellsValues[i]);
        }

        public object this[int i]
        {
            get { return _cellsValues[i]; }
        }

        #endregion IExcelDataReader Members

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.

            if (!this.disposed)
            {
                if (disposing)
                {
                    if (_xmlReader != null) ((IDisposable)_xmlReader).Dispose();
                    if (_sheetStream != null) _sheetStream.Dispose();
                    if (_zipWorker != null) _zipWorker.Dispose();
                }

                _zipWorker = null;
                _xmlReader = null;
                _sheetStream = null;

                _workbook = null;
                _cellsValues = null;
                _savedCellsValues = null;

                disposed = true;
            }
        }

        ~ExcelOpenXmlReader()
        {
            Dispose(false);
        }

        #endregion IDisposable Members

       

        
    }
}