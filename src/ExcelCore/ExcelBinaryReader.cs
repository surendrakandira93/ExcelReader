using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelCore;
using ExcelCore.Core;
using ExcelCore.Core.BinaryFormat;
using ExcelCore.Model;
using Newtonsoft.Json;

namespace ExcelCore
{
    /// <summary>
    /// ExcelDataReader Class
    /// </summary>
    public class ExcelBinaryReader : IExcelDataReader
    {
        #region Members

        private Stream m_file;
        private XlsHeader m_hdr;
        private List<XlsWorksheet> m_sheets;
        private XlsBiffStream m_stream;
        private XlsWorkbookGlobals m_globals;
        private ushort m_version;
        private bool m_ConvertOADate;
        private Encoding m_encoding;
        private bool m_isValid;
        private bool m_isClosed;
        private readonly Encoding m_Default_Encoding = Encoding.UTF8;
        private string m_exceptionMessage;
        private object[] m_cellsValues;
        private uint[] m_dbCellAddrs;
        private int m_dbCellAddrsIndex;
        private bool m_canRead;
        private int m_SheetIndex;
        private int m_depth;
        private int m_cellOffset;
        private int m_maxCol;
        private int m_maxRow;
        private bool m_noIndex;
        private XlsBiffRow m_currentRowRecord;

        private bool m_IsFirstRead;
        private bool _isFirstRowAsColumnNames;

        private const string WORKBOOK = "Workbook";
        private const string BOOK = "Book";
        private const string COLUMN = "Column";

        private bool disposed;

        #endregion Members

        internal ExcelBinaryReader()
        {
            m_encoding = m_Default_Encoding;
            m_version = 0x0600;
            m_isValid = true;
            m_SheetIndex = -1;
            m_IsFirstRead = true;
        }

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
                    if (m_sheets != null) m_sheets.Clear();
                }
                m_sheets = null;
                m_stream = null;
                m_globals = null;
                m_encoding = null;
                m_hdr = null;

                disposed = true;
            }
        }

        ~ExcelBinaryReader()
        {
            Dispose(false);
        }

        #endregion IDisposable Members

        #region Private methods

        private int findFirstDataCellOffset(int startOffset)
        {
            //seek to the first dbcell record
            var record = m_stream.ReadAt(startOffset);
            while (!(record is XlsBiffDbCell))
            {
                if (m_stream.Position >= m_stream.Size)
                    return -1;

                if (record is XlsBiffEOF)
                    return -1;

                record = m_stream.Read();
            }

            XlsBiffDbCell startCell = (XlsBiffDbCell)record;
            XlsBiffRow row = null;

            int offs = startCell.RowAddress;

            do
            {
                row = m_stream.ReadAt(offs) as XlsBiffRow;
                if (row == null) break;

                offs += row.Size;
            } while (null != row);

            return offs;
        }

        private void readWorkBookGlobals()
        {
            //Read Header
            try
            {
                m_hdr = XlsHeader.ReadHeader(m_file);
            }
            catch (ExcelCore.Exceptions.HeaderException ex)
            {
                fail(ex.Message);
                return;
            }
            catch (FormatException ex)
            {
                fail(ex.Message);
                return;
            }

            XlsRootDirectory dir = new XlsRootDirectory(m_hdr);
            XlsDirectoryEntry workbookEntry = dir.FindEntry(WORKBOOK) ?? dir.FindEntry(BOOK);

            if (workbookEntry == null)
            { fail(Errors.ErrorStreamWorkbookNotFound); return; }

            if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
            { fail(Errors.ErrorWorkbookIsNotStream); return; }

            m_stream = new XlsBiffStream(m_hdr, workbookEntry.StreamFirstSector, workbookEntry.IsEntryMiniStream, dir);

            m_globals = new XlsWorkbookGlobals();

            m_stream.Seek(0, SeekOrigin.Begin);

            XlsBiffRecord rec = m_stream.Read();
            XlsBiffBOF bof = rec as XlsBiffBOF;

            if (bof == null || bof.Type != BIFFTYPE.WorkbookGlobals)
            { fail(Errors.ErrorWorkbookGlobalsInvalidData); return; }

            bool sst = false;

            m_version = bof.Version;
            m_sheets = new List<XlsWorksheet>();

            while (null != (rec = m_stream.Read()))
            {
                switch (rec.ID)
                {
                    case BIFFRECORDTYPE.INTERFACEHDR:
                        m_globals.InterfaceHdr = (XlsBiffInterfaceHdr)rec;
                        break;

                    case BIFFRECORDTYPE.BOUNDSHEET:
                        XlsBiffBoundSheet sheet = (XlsBiffBoundSheet)rec;

                        if (sheet.Type != XlsBiffBoundSheet.SheetType.Worksheet) break;

                        sheet.IsV8 = isV8();
                        sheet.UseEncoding = m_encoding;

                        m_sheets.Add(new XlsWorksheet(m_globals.Sheets.Count, sheet));
                        m_globals.Sheets.Add(sheet);

                        break;

                    case BIFFRECORDTYPE.MMS:
                        m_globals.MMS = rec;
                        break;

                    case BIFFRECORDTYPE.COUNTRY:
                        m_globals.Country = rec;
                        break;

                    case BIFFRECORDTYPE.CODEPAGE:

                        m_globals.CodePage = (XlsBiffSimpleValueRecord)rec;

                        try
                        {
                            m_encoding = Encoding.GetEncoding(m_globals.CodePage.Value);
                        }
                        catch (ArgumentException)
                        {
                            // Warning - Password protection
                            // TODO: Attach to ILog
                        }

                        break;

                    case BIFFRECORDTYPE.FONT:
                    case BIFFRECORDTYPE.FONT_V34:
                        m_globals.Fonts.Add(rec);
                        break;

                    case BIFFRECORDTYPE.FORMAT_V23:
                        {
                            var fmt = (XlsBiffFormatString)rec;
                            fmt.UseEncoding = m_encoding;
                            m_globals.Formats.Add((ushort)m_globals.Formats.Count, fmt);
                        }
                        break;

                    case BIFFRECORDTYPE.FORMAT:
                        {
                            var fmt = (XlsBiffFormatString)rec;
                            m_globals.Formats.Add(fmt.Index, fmt);
                        }
                        break;

                    case BIFFRECORDTYPE.XF:
                    case BIFFRECORDTYPE.XF_V4:
                    case BIFFRECORDTYPE.XF_V3:
                    case BIFFRECORDTYPE.XF_V2:
                        m_globals.ExtendedFormats.Add(rec);
                        break;

                    case BIFFRECORDTYPE.SST:
                        m_globals.SST = (XlsBiffSST)rec;
                        sst = true;
                        break;

                    case BIFFRECORDTYPE.CONTINUE:
                        if (!sst) break;
                        XlsBiffContinue contSST = (XlsBiffContinue)rec;
                        m_globals.SST.Append(contSST);
                        break;

                    case BIFFRECORDTYPE.EXTSST:
                        m_globals.ExtSST = rec;
                        sst = false;
                        break;

                    case BIFFRECORDTYPE.PROTECT:
                    case BIFFRECORDTYPE.PASSWORD:
                    case BIFFRECORDTYPE.PROT4REVPASSWORD:
                        //IsProtected
                        break;

                    case BIFFRECORDTYPE.EOF:
                        if (m_globals.SST != null)
                            m_globals.SST.ReadStrings();
                        return;

                    default:
                        continue;
                }
            }
        }

        private bool readWorkSheetGlobals(XlsWorksheet sheet, out XlsBiffIndex idx, out XlsBiffRow row)
        {
            idx = null;
            row = null;

            m_stream.Seek((int)sheet.DataOffset, SeekOrigin.Begin);

            XlsBiffBOF bof = m_stream.Read() as XlsBiffBOF;
            if (bof == null || bof.Type != BIFFTYPE.Worksheet) return false;

            //DumpBiffRecords();

            XlsBiffRecord rec = m_stream.Read();
            if (rec == null) return false;
            if (rec is XlsBiffIndex)
            {
                idx = rec as XlsBiffIndex;
            }
            else if (rec is XlsBiffUncalced)
            {
                // Sometimes this come before the index...
                idx = m_stream.Read() as XlsBiffIndex;
            }

            //if (null == idx)
            //{
            //	// There is a record before the index! Chech his type and see the MS Biff Documentation
            //	return false;
            //}

            if (idx != null)
                idx.IsV8 = isV8();

            XlsBiffRecord trec;
            XlsBiffDimensions dims = null;

            do
            {
                trec = m_stream.Read();
                if (trec.ID == BIFFRECORDTYPE.DIMENSIONS)
                {
                    dims = (XlsBiffDimensions)trec;
                    break;
                }
            } while (trec != null && trec.ID != BIFFRECORDTYPE.ROW);

            //if we are already on row record then set that as the row, otherwise step forward till we get to a row record
            if (trec.ID == BIFFRECORDTYPE.ROW)
                row = (XlsBiffRow)trec;

            XlsBiffRow rowRecord = null;
            while (rowRecord == null)
            {
                if (m_stream.Position >= m_stream.Size)
                    break;
                var thisRec = m_stream.Read();
                if (thisRec is XlsBiffEOF)
                    break;
                rowRecord = thisRec as XlsBiffRow;
            }

            row = rowRecord;

            m_maxCol = 256;

            if (dims != null)
            {
                dims.IsV8 = isV8();
                m_maxCol = dims.LastColumn - 1;
                sheet.Dimensions = dims;
            }

            m_maxRow = idx == null ? (int)dims.LastRow : (int)idx.LastExistingRow;

            if (idx != null && idx.LastExistingRow <= idx.FirstExistingRow)
            {
                return false;
            }
            else if (row == null)
            {
                return false;
            }

            m_depth = 0;

            return true;
        }

        private void DumpBiffRecords()
        {
            XlsBiffRecord rec = null;
            var startPos = m_stream.Position;

            do
            {
                rec = m_stream.Read();
                Console.WriteLine(rec.ID.ToString());
            } while (rec != null && m_stream.Position < m_stream.Size);

            m_stream.Seek(startPos, SeekOrigin.Begin);
        }

        private bool readWorkSheetRow()
        {
            m_cellsValues = new object[m_maxCol];

            while (m_cellOffset < m_stream.Size)
            {
                XlsBiffRecord rec = m_stream.ReadAt(m_cellOffset);
                m_cellOffset += rec.Size;

                if ((rec is XlsBiffDbCell)) { break; };//break;
                if (rec is XlsBiffEOF) { return false; };

                XlsBiffBlankCell cell = rec as XlsBiffBlankCell;

                if ((null == cell) || (cell.ColumnIndex >= m_maxCol)) continue;
                if (cell.RowIndex != m_depth) { m_cellOffset -= rec.Size; break; };

                pushCellValue(cell);
            }

            m_depth++;

            return m_depth < m_maxRow;
        }

        private void pushCellValue(XlsBiffBlankCell cell)
        {
            double _dValue;

            switch (cell.ID)
            {
                case BIFFRECORDTYPE.BOOLERR:
                    if (cell.ReadByte(7) == 0)
                        m_cellsValues[cell.ColumnIndex] = cell.ReadByte(6) != 0;
                    break;

                case BIFFRECORDTYPE.BOOLERR_OLD:
                    if (cell.ReadByte(8) == 0)
                        m_cellsValues[cell.ColumnIndex] = cell.ReadByte(7) != 0;
                    break;

                case BIFFRECORDTYPE.INTEGER:
                case BIFFRECORDTYPE.INTEGER_OLD:
                    m_cellsValues[cell.ColumnIndex] = ((XlsBiffIntegerCell)cell).Value;
                    break;

                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:

                    _dValue = ((XlsBiffNumberCell)cell).Value;

                    m_cellsValues[cell.ColumnIndex] = !m_ConvertOADate ?
                        _dValue : tryConvertOADateTime(_dValue, cell.XFormat);

                    break;

                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.RSTRING:
                    m_cellsValues[cell.ColumnIndex] = ((XlsBiffLabelCell)cell).Value;
                    break;

                case BIFFRECORDTYPE.LABELSST:
                    string tmp = m_globals.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex);
                    m_cellsValues[cell.ColumnIndex] = tmp;
                    break;

                case BIFFRECORDTYPE.RK:

                    _dValue = ((XlsBiffRKCell)cell).Value;

                    m_cellsValues[cell.ColumnIndex] = !m_ConvertOADate ?
                        _dValue : tryConvertOADateTime(_dValue, cell.XFormat);

                    break;

                case BIFFRECORDTYPE.MULRK:

                    XlsBiffMulRKCell _rkCell = (XlsBiffMulRKCell)cell;
                    for (ushort j = cell.ColumnIndex; j <= _rkCell.LastColumnIndex; j++)
                    {
                        _dValue = _rkCell.GetValue(j);
                        m_cellsValues[j] = !m_ConvertOADate ? _dValue : tryConvertOADateTime(_dValue, _rkCell.GetXF(j));
                    }

                    break;

                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                case BIFFRECORDTYPE.MULBLANK:
                    // Skip blank cells

                    break;

                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_OLD:

                    object _oValue = ((XlsBiffFormulaCell)cell).Value;

                    if (null != _oValue && _oValue is FORMULAERROR)
                    {
                        _oValue = null;
                    }
                    else
                    {
                        m_cellsValues[cell.ColumnIndex] = !m_ConvertOADate ?
                            _oValue : tryConvertOADateTime(_oValue, (ushort)(cell.XFormat));//date time offset
                    }

                    break;

                default:
                    break;
            }
        }

        private bool moveToNextRecord()
        {
            //if sheet has no index
            if (m_noIndex)
            {
                return moveToNextRecordNoIndex();
            }

            //if sheet has index
            if (null == m_dbCellAddrs ||
                m_dbCellAddrsIndex == m_dbCellAddrs.Length ||
                m_depth == m_maxRow) return false;

            m_canRead = readWorkSheetRow();

            //read last row
            if (!m_canRead && m_depth > 0) m_canRead = true;

            if (!m_canRead && m_dbCellAddrsIndex < (m_dbCellAddrs.Length - 1))
            {
                m_dbCellAddrsIndex++;
                m_cellOffset = findFirstDataCellOffset((int)m_dbCellAddrs[m_dbCellAddrsIndex]);
                if (m_cellOffset < 0)
                    return false;
                m_canRead = readWorkSheetRow();
            }

            return m_canRead;
        }

        private bool moveToNextRecordNoIndex()
        {
            //seek from current row record to start of cell data where that cell relates to the next row record
            XlsBiffRow rowRecord = m_currentRowRecord;

            if (rowRecord == null)
                return false;

            if (rowRecord.RowIndex < m_depth)
            {
                m_stream.Seek(rowRecord.Offset + rowRecord.Size, SeekOrigin.Begin);
                do
                {
                    if (m_stream.Position >= m_stream.Size)
                        return false;

                    var record = m_stream.Read();
                    if (record is XlsBiffEOF)
                        return false;

                    rowRecord = record as XlsBiffRow;
                } while (rowRecord == null || rowRecord.RowIndex < m_depth);
            }

            m_currentRowRecord = rowRecord;
            //m_depth = m_currentRowRecord.RowIndex;

            //we have now found the row record for the new row, the we need to seek forward to the first cell record
            XlsBiffBlankCell cell = null;
            do
            {
                if (m_stream.Position >= m_stream.Size)
                    return false;

                var record = m_stream.Read();
                if (record is XlsBiffEOF)
                    return false;

                if (record.IsCell)
                {
                    var candidateCell = record as XlsBiffBlankCell;
                    if (candidateCell != null)
                    {
                        if (candidateCell.RowIndex == m_currentRowRecord.RowIndex)
                            cell = candidateCell;
                    }
                }
            } while (cell == null);

            m_cellOffset = cell.Offset;
            m_canRead = readWorkSheetRow();

            //read last row
            //if (!m_canRead && m_depth > 0) m_canRead = true;

            //if (!m_canRead && m_dbCellAddrsIndex < (m_dbCellAddrs.Length - 1))
            //{
            //	m_dbCellAddrsIndex++;
            //	m_cellOffset = findFirstDataCellOffset((int)m_dbCellAddrs[m_dbCellAddrsIndex]);

            //	m_canRead = readWorkSheetRow();
            //}

            return m_canRead;
        }

        private void initializeSheetRead()
        {
            if (m_SheetIndex == ResultsCount) return;

            m_dbCellAddrs = null;

            m_IsFirstRead = false;

            if (m_SheetIndex == -1) m_SheetIndex = 0;

            XlsBiffIndex idx;

            if (!readWorkSheetGlobals(m_sheets[m_SheetIndex], out idx, out m_currentRowRecord))
            {
                //read next sheet
                m_SheetIndex++;
                initializeSheetRead();
                return;
            };

            if (idx == null)
            {
                //no index, but should have the first row record
                m_noIndex = true;
            }
            else
            {
                m_dbCellAddrs = idx.DbCellAddresses;
                m_dbCellAddrsIndex = 0;
                m_cellOffset = findFirstDataCellOffset((int)m_dbCellAddrs[m_dbCellAddrsIndex]);
                if (m_cellOffset < 0)
                {
                    fail("Badly formed binary file. Has INDEX but no DBCELL");
                    return;
                }
            }
        }

        private void fail(string message)
        {
            m_exceptionMessage = message;
            m_isValid = false;

            m_file.Dispose();
            m_isClosed = true;
            m_sheets = null;
            m_stream = null;
            m_globals = null;
            m_encoding = null;
            m_hdr = null;
        }

        private object tryConvertOADateTime(long value, ushort XFormat)
        {
            ushort format = 0;
            if (XFormat >= 0 && XFormat < m_globals.ExtendedFormats.Count)
            {
                var rec = m_globals.ExtendedFormats[XFormat];
                switch (rec.ID)
                {
                    case BIFFRECORDTYPE.XF_V2:
                        format = (ushort)(rec.ReadByte(2) & 0x3F);
                        break;

                    case BIFFRECORDTYPE.XF_V3:
                        if ((rec.ReadByte(3) & 4) == 0)
                            return value;
                        format = rec.ReadByte(1);
                        break;

                    case BIFFRECORDTYPE.XF_V4:
                        if ((rec.ReadByte(5) & 4) == 0)
                            return value;
                        format = rec.ReadByte(1);
                        break;

                    default:
                        if ((rec.ReadByte(m_globals.Sheets[m_globals.Sheets.Count - 1].IsV8 ? 9 : 7) & 4) == 0)
                            return value;

                        format = rec.ReadUInt16(2);
                        break;
                }
            }
            else
            {
                format = XFormat;
            }

            switch (format)
            {
                // numeric built in formats
                case 0: //"General";
                case 1: //"0";
                case 2: //"0.00";
                case 3: //"#,##0";
                case 4: //"#,##0.00";
                case 5: //"\"$\"#,##0_);(\"$\"#,##0)";
                case 6: //"\"$\"#,##0_);[Red](\"$\"#,##0)";
                case 7: //"\"$\"#,##0.00_);(\"$\"#,##0.00)";
                case 8: //"\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
                case 9: //"0%";
                case 10: //"0.00%";
                case 11: //"0.00E+00";
                case 12: //"# ?/?";
                case 13: //"# ??/??";
                case 0x30:// "##0.0E+0";

                case 0x25:// "_(#,##0_);(#,##0)";
                case 0x26:// "_(#,##0_);[Red](#,##0)";
                case 0x27:// "_(#,##0.00_);(#,##0.00)";
                case 40:// "_(#,##0.00_);[Red](#,##0.00)";
                case 0x29:// "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)";
                case 0x2a:// "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)";
                case 0x2b:// "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)";
                case 0x2c:// "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)";
                    return value;

                // date formats
                case 14: //this.GetDefaultDateFormat();
                case 15: //"D-MM-YY";
                case 0x10: // "D-MMM";
                case 0x11: // "MMM-YY";
                case 0x12: // "h:mm AM/PM";
                case 0x13: // "h:mm:ss AM/PM";
                case 20: // "h:mm";
                case 0x15: // "h:mm:ss";
                case 0x16: // string.Format("{0} {1}", this.GetDefaultDateFormat(), this.GetDefaultTimeFormat());

                case 0x2d: // "mm:ss";
                case 0x2e: // "[h]:mm:ss";
                case 0x2f: // "mm:ss.0";
                    return Helpers.ConvertFromOATime(value);

                case 0x31:// "@";
                    return value.ToString();

                default:
                    XlsBiffFormatString fmtString;
                    if (m_globals.Formats.TryGetValue(format, out fmtString))
                    {
                        var fmt = fmtString.Value;
                        var formatReader = new FormatReader() { FormatString = fmt };
                        if (formatReader.IsDateFormatString())
                            return Helpers.ConvertFromOATime(value);
                    }
                    return value;
            }
        }

        private object tryConvertOADateTime(object value, ushort XFormat)
        {
            long _dValue;

            if (long.TryParse(value.ToString(), out _dValue))
                return DateTime.FromBinary(_dValue);

            return value;
        }

        private bool isV8()
        {
            return m_version >= 0x600;
        }

        #endregion Private methods

        #region IExcelDataReader Members

        public void Initialize(Stream fileStream)
        {
            m_file = fileStream;

            readWorkBookGlobals();
        }

        public IEnumerable AsIEnumerable()
        {
            Dictionary<int, List<KeyValuePair<string, object>>> response = new Dictionary<int, List<KeyValuePair<string, object>>>();
            if (!m_isValid) return null;

            //for (int ind = 0; ind < m_sheets.Count; ind++)
            for (int ind = 0; ind < 1; ind++)
            {
                XlsBiffIndex idx;
                if (!readWorkSheetGlobals(m_sheets[ind], out idx, out m_currentRowRecord)) return null;
                List<string> ColArr = new List<string>();
                while (Read())
                {
                    if (m_depth == m_maxRow) break;
                    if (m_depth == 1)
                    {
                        for (int index = 0; index < m_maxCol; index++)
                        {
                            if (m_cellsValues[index] != null)
                            {
                                ColArr.Add(m_cellsValues[index].ToString());
                            }
                        }
                    }

                    if (m_depth > 0 && m_depth != 1)
                        response.Add(m_depth - 1, GetItem(ColArr, m_cellsValues));
                }

                if (m_depth > 0 && m_maxRow != 1)
                    response.Add(m_depth, GetItem(ColArr, m_cellsValues));
            }

            return response.Select(x => x.Value).ToList();
        }

        public IEnumerable AsIEnumerable(List<ExcelHeaderKeyValues> headerValues)
        {
            Dictionary<int, List<KeyValuePair<string, object>>> response = new Dictionary<int, List<KeyValuePair<string, object>>>();
            if (!m_isValid) return null;

            //for (int ind = 0; ind < m_sheets.Count; ind++)
            for (int ind = 0; ind < 1; ind++)
            {
                XlsBiffIndex idx;
                if (!readWorkSheetGlobals(m_sheets[ind], out idx, out m_currentRowRecord)) return null;
                List<string> ColArr = new List<string>();
                bool InvalidColumn = false;
                while (Read())
                {
                    if (m_depth == m_maxRow) break;
                    if (m_depth == 1)
                    {
                        for (int index = 0; index < m_maxCol; index++)
                        {
                            if (m_cellsValues[index] != null)
                            {
                                var headerTitle = headerValues.FirstOrDefault(x => x.HeaderTitle.ToLower().Equals(Convert.ToString(m_cellsValues[index]).ToLower()));
                                if (headerTitle != null)
                                    ColArr.Add(headerTitle.HeaderKey);
                                else
                                    InvalidColumn = true;
                            }
                        }

                        if (InvalidColumn)
                            throw (new Exception("Column name not match ! Please correct column name.  Invalid column name '" + string.Join(",", ColArr.ToArray()) + "'"));
                    }

                    if (m_depth > 0 && m_depth != 1)
                        response.Add(m_depth - 1, GetItem(ColArr, m_cellsValues));
                }

                if (m_depth > 0 && m_maxRow != 1)
                    response.Add(m_depth, GetItem(ColArr, m_cellsValues));
            }

            return response.Select(x => x.Value).ToList();
        }

        public List<T> AsIEnumerable<T>(List<ExcelHeaderKeyValues> headerValues)
        {
            List<T> response = new List<T>();
            if (!m_isValid) return null;

            //for (int ind = 0; ind < m_sheets.Count; ind++)
            for (int ind = 0; ind < 1; ind++)
            {
                XlsBiffIndex idx;
                if (!readWorkSheetGlobals(m_sheets[ind], out idx, out m_currentRowRecord)) return null;
                List<string> ColArr = new List<string>();
                bool InvalidColumn = false;
                while (Read())
                {
                    if (m_depth == m_maxRow) break;
                    if (m_depth == 1)
                    {
                        for (int index = 0; index < m_maxCol; index++)
                        {
                            if (m_cellsValues[index] != null)
                            {
                                var headerTitle = headerValues.FirstOrDefault(x => x.HeaderTitle.ToLower().Equals(Convert.ToString(m_cellsValues[index]).ToLower()));
                                if (headerTitle != null)
                                    ColArr.Add(headerTitle.HeaderKey);
                                else
                                    InvalidColumn = true;
                            }
                        }

                        if (InvalidColumn)
                            throw (new Exception("Column name not match ! Please correct column name.  Invalid column name '" + string.Join(",", ColArr.ToArray()) + "'"));
                    }

                    if (m_depth > 0 && m_depth != 1)
                        response.Add(GetItem<T>(ColArr, m_cellsValues));
                }

                if (m_depth > 0 && m_maxRow != 1)
                    response.Add(GetItem<T>(ColArr, m_cellsValues));
            }

            return response;
        }

        public string AsJson(List<ExcelHeaderKeyValues> headerValues)
        {
            StringBuilder response = new StringBuilder();
            response.Append("[");
            if (!m_isValid) return null;
            //for (int ind = 0; ind < m_sheets.Count; ind++)
            for (int ind = 0; ind < 1; ind++)
            {
                XlsBiffIndex idx;
                if (!readWorkSheetGlobals(m_sheets[ind], out idx, out m_currentRowRecord)) return null;
                List<string> ColArr = new List<string>();
                bool InvalidColumn = false;
                bool IsFirstRow = true;
                while (Read())
                {
                    if (m_depth == m_maxRow) break;
                    if (m_depth == 1 && IsFirstRow)
                    {
                        IsFirstRow = false;
                        for (int index = 0; index < m_maxCol; index++)
                        {
                            if (m_cellsValues[index] != null)
                            {
                                var headerTitle = headerValues.FirstOrDefault(x => x.HeaderTitle.ToLower().Equals(Convert.ToString(m_cellsValues[index]).ToLower()));
                                if (headerTitle != null)
                                    ColArr.Add(headerTitle.HeaderKey);
                                else
                                    InvalidColumn = true;
                            }
                        }

                        if (InvalidColumn)
                            throw (new Exception("Column name not match ! Please correct column name.  Invalid column name '" + string.Join(",", ColArr.ToArray()) + "'"));
                    }

                    if (m_depth > 0 && m_depth != 1)
                    {
                        response.Append(AddRow(m_cellsValues, ColArr));
                        response.Append(",");
                    }
                }

                if (m_depth > 0 && m_maxRow != 1)
                {
                    response.Append(AddRow(m_cellsValues, ColArr));
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
            if (!m_isValid) return null;
            //for (int ind = 0; ind < m_sheets.Count; ind++)
            for (int ind = 0; ind < 1; ind++)
            {
                XlsBiffIndex idx;

                if (!readWorkSheetGlobals(m_sheets[ind], out idx, out m_currentRowRecord)) return null;
                List<string> ColArr = new List<string>();
                while (Read())
                {
                    if (m_depth == m_maxRow) break;

                    bool justAddedColumns = false;

                    if (m_depth == 1)
                    {
                        for (int i = 0; i < m_maxCol; i++)
                        {
                            if (m_cellsValues[i] != null)
                            {
                                ColArr.Add((string)m_cellsValues[i]);
                            }
                        }
                    }

                    if (!justAddedColumns && m_depth > 0 && m_depth != 1)
                    {
                        response.Append(AddRow(m_cellsValues, ColArr));
                        response.Append(",");
                    }
                }

                if (m_depth > 0 && m_maxRow != 1)
                {
                    response.Append(AddRow(m_cellsValues, ColArr));
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
                PropertyInfo pro = temp.GetProperties().FirstOrDefault(x => x.Name.ToLower().Equals(keyName[i].ToLower()));
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
                responseVal.Add(titleValue[i], m_cellsValues[i]);
            string json = JsonConvert.SerializeObject(responseVal, Formatting.Indented);
            return json;
        }

        public string ExceptionMessage
        {
            get { return m_exceptionMessage; }
        }

        public string Name
        {
            get
            {
                if (null != m_sheets && m_sheets.Count > 0)
                    return m_sheets[m_SheetIndex].Name;
                else
                    return null;
            }
        }

        public bool IsValid
        {
            get { return m_isValid; }
        }

        public void Close()
        {
            m_file.Dispose();
            m_isClosed = true;
        }

        public int Depth
        {
            get { return m_depth; }
        }

        public int ResultsCount
        {
            get { return m_globals.Sheets.Count; }
        }

        public bool IsClosed
        {
            get { return m_isClosed; }
        }

        public bool NextResult()
        {
            if (m_SheetIndex >= (this.ResultsCount - 1)) return false;

            m_SheetIndex++;

            m_IsFirstRead = true;

            return true;
        }

        public bool Read()
        {
            if (!m_isValid) return false;

            if (m_IsFirstRead) initializeSheetRead();

            return moveToNextRecord();
        }

        public int FieldCount
        {
            get { return m_maxCol; }
        }

        public bool GetBoolean(int i)
        {
            if (IsDBNull(i)) return false;

            return Boolean.Parse(m_cellsValues[i].ToString());
        }

        public DateTime GetDateTime(int i)
        {
            if (IsDBNull(i)) return DateTime.MinValue;

            string val = m_cellsValues[i].ToString();
            long dVal;

            try
            {
                dVal = long.Parse(val);
            }
            catch (FormatException)
            {
                return DateTime.Parse(val);
            }

            return DateTime.FromBinary(dVal);
        }

        public decimal GetDecimal(int i)
        {
            if (IsDBNull(i)) return decimal.MinValue;

            return decimal.Parse(m_cellsValues[i].ToString());
        }

        public double GetDouble(int i)
        {
            if (IsDBNull(i)) return double.MinValue;

            return double.Parse(m_cellsValues[i].ToString());
        }

        public float GetFloat(int i)
        {
            if (IsDBNull(i)) return float.MinValue;

            return float.Parse(m_cellsValues[i].ToString());
        }

        public short GetInt16(int i)
        {
            if (IsDBNull(i)) return short.MinValue;

            return short.Parse(m_cellsValues[i].ToString());
        }

        public int GetInt32(int i)
        {
            if (IsDBNull(i)) return int.MinValue;

            return int.Parse(m_cellsValues[i].ToString());
        }

        public long GetInt64(int i)
        {
            if (IsDBNull(i)) return long.MinValue;

            return long.Parse(m_cellsValues[i].ToString());
        }

        public string GetString(int i)
        {
            if (IsDBNull(i)) return null;

            return m_cellsValues[i].ToString();
        }

        public object GetValue(int i)
        {
            return m_cellsValues[i];
        }

        public bool IsDBNull(int i)
        {
            return (null == m_cellsValues[i]);
        }

        public object this[int i]
        {
            get { return m_cellsValues[i]; }
        }

        #endregion IExcelDataReader Members

      

       

        
    }
}