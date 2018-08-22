using ClosedXML.Excel;
using ListAnalyzer.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace ListAnalyzer
{
    public static class HelperFunctions
    {
        public static string SerializeObject<T>(T source)
        {
            var serializer = new XmlSerializer(typeof(T));

            using (var sw = new StringWriter())
            using (var writer = new XmlTextWriter(sw))
            {
                serializer.Serialize(writer, source);
                return sw.ToString();
            }
        }

        public static T DeSerializeObject<T>(string xml)
        {
            using (var sr = new StringReader(xml))
            {
                var serializer = new XmlSerializer(typeof(T));
                return (T)serializer.Deserialize(sr);
            }
        }

        public static object ReturnZeroIfNull(this object value)
        {
            if (value == DBNull.Value)
                return 0;
            if (value == null)
                return 0;
            return value;
        }

        public static object ReturnEmptyIfNull(this object value)
        {
            if (value == DBNull.Value)
                return string.Empty;
            if (value == null)
                return string.Empty;
            return value;
        }

        public static object ReturnFalseIfNull(this object value)
        {
            if (value == DBNull.Value)
                return false;
            if (value == null)
                return false;
            return value;
        }

        public static object ReturnDateTimeMinIfNull(this object value)
        {
            if (value == DBNull.Value)
                return DateTime.MinValue;
            if (value == null)
                return DateTime.MinValue;
            return value;
        }

        public static object ReturnNullIfDbNull(this object value)
        {
            if (value == DBNull.Value)
                return '\0';
            if (value == null)
                return '\0';
            return value;
        }

        //This function formats the display-name of a user,
        //and removes unnecessary extra information.
        public static string FormatUserDisplayName(string displayName = null, string defaultValue = "tBill Users",
            bool returnNameIfExists = false, bool returnAddressPartIfExists = false)
        {
            //Get the first part of the Users's Display Name if s/he has a name like this: "firstname lastname (extra text)"
            //removes the "(extra text)" part
            if (!string.IsNullOrEmpty(displayName))
            {
                if (returnNameIfExists)
                    return Regex.Replace(displayName, @"\ \(\w{1,}\)", "");
                return (displayName.Split(' '))[0];
            }
            if (returnAddressPartIfExists)
            {
                var emailParts = defaultValue.Split('@');
                return emailParts[0];
            }
            return defaultValue;
        }

        public static string FormatUserTelephoneNumber(this string telephoneNumber)
        {
            var result = string.Empty;

            if (!string.IsNullOrEmpty(telephoneNumber))
            {
                //result = telephoneNumber.ToLower().Trim().Trim('+').Replace("tel:", "");
                result = telephoneNumber.ToLower().Trim().Replace("tel:", "");

                if (result.Contains(";"))
                {
                    if (!result.ToLower().Contains(";ext="))
                        result = result.Split(';')[0];
                }
            }

            return result;
        }

        public static bool IsValidEmail(this string emailAddress)
        {
            const string pattern = @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z";

            return Regex.IsMatch(emailAddress, pattern);
        }

        /// <summary>
        /// Convert DateTime to string
        /// </summary>
        /// <param name="datetTime"></param>
        /// <param name="excludeHoursAndMinutes">if true it will execlude time from datetime string. Default is false</param>
        /// <returns></returns>
        public static string ConvertDate(this DateTime datetTime, bool excludeHoursAndMinutes = false)
        {
            if (datetTime != DateTime.MinValue)
            {
                if (excludeHoursAndMinutes)
                    return datetTime.ToString("yyyy-MM-dd");
                return datetTime.ToString("yyyy-MM-dd HH:mm:ss.fff");
            }
            return null;
        }

        [SuppressMessage("ReSharper", "PossibleLossOfFraction")]
        public static string ConvertSecondsToReadable(this int secondsParam)
        {
            var hours = Convert.ToInt32(Math.Floor((double)(secondsParam / 3600)));
            var minutes = Convert.ToInt32(Math.Floor((double)(secondsParam - (hours * 3600)) / 60));
            var seconds = secondsParam - (hours * 3600) - (minutes * 60);

            var hoursStr = hours.ToString();
            var minsStr = minutes.ToString();
            var secsStr = seconds.ToString();

            if (hours < 10)
            {
                hoursStr = "0" + hoursStr;
            }

            if (minutes < 10)
            {
                minsStr = "0" + minsStr;
            }
            if (seconds < 10)
            {
                secsStr = "0" + secsStr;
            }

            return hoursStr + ':' + minsStr + ':' + secsStr;
        }

        [SuppressMessage("ReSharper", "PossibleLossOfFraction")]
        public static string ConvertSecondsToReadable(this long secondsParam)
        {
            var hours = Convert.ToInt32(Math.Floor((double)(secondsParam / 3600)));
            var minutes = Convert.ToInt32(Math.Floor((double)(secondsParam - (hours * 3600)) / 60));
            var seconds = Convert.ToInt32(secondsParam - (hours * 3600) - (minutes * 60));

            var hoursStr = hours.ToString();
            var minsStr = minutes.ToString();
            var secsStr = seconds.ToString();

            if (hours < 10)
            {
                hoursStr = "0" + hoursStr;
            }

            if (minutes < 10)
            {
                minsStr = "0" + minsStr;
            }
            if (seconds < 10)
            {
                secsStr = "0" + secsStr;
            }

            return hoursStr + ':' + minsStr + ':' + secsStr;
        }

        public static List<Report> ExcelToList(string importPath)
        {
            string connectionStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + importPath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";
            connectionStr = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + importPath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007

            DataTable dataTable = new DataTable();
            using (OleDbConnection connection = new OleDbConnection(connectionStr))
            {
                try
                {
                    connection.Open();
                    DataTable dt = connection.GetSchema("Tables");
                    if (dt == null || dt.Rows.Count <= 0) return null;
                    string firstSheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + firstSheetName + "]", connection);//here we read data from sheet1
                    oleAdpt.Fill(dataTable);//fill excel data into dataTable
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(ex.Message.ToString());
                }
            }
            return dataTable.ToList<Report>();
        }

        public static List<Report> CountDuplicate(List<Report> reports)
        {
            return reports.OrderBy(x => x.DateTime).GroupBy(x => new { x.CID, x.LAC })
                               .Where(x => x.Count() > 1)
                               .Select(x => new Report
                               {
                                   CID = x.First().CID,
                                   LAC = x.First().LAC,
                                   Count = x.Count(),
                                   FirstAppear = x.First().DateTime,
                                   LastAppear = x.Last().DateTime
                               }).OrderByDescending(x => x.Count).ToList();

        }

        public static List<Report> FindOverlap(List<Report> reports)
        {
            return reports
                .FindAll(
                    x => x.DateTime >= new DateTime(x.DateTime.Year, x.DateTime.Month, x.DateTime.Day, 19, 20, 0) &&
                    x.DateTime <= new DateTime(x.DateTime.Year, x.DateTime.Month, x.DateTime.Day, 19, 30, 0));
        }

        public static void ExportReport(string saveFilePath, List<List<Report>> data)
        {
            if (data == null) return;
            List<DataTable> list = new List<DataTable>();
            List<int> rowNamePos = new List<int>();
            List<string> reportNames = new List<string>();
            #region Duplicate Report
            List<string> columnNames = new List<string> { "Thời gian bắt đầu", "Thời gian kết thúc", "Cell ID", "LAC", "Vị trí","Số lần xuất hiện"};
            var reportname = "Phân tích list";
            reportNames.Add(reportname);
            DataTable table = new DataTable();
            table.TableName = reportname;

            foreach (string columnName in columnNames)
            {
                table.Columns.Add(columnName);
            }

            /*
             * CreateReportHeader has to be called after add columns
             */
            int rowForColumnName = CreateReportHeader(table, reportname);
            rowNamePos.Add(rowForColumnName);
            table.Rows.Add(columnNames.ToArray());
            foreach (Report item in data[0])
            {
                table.Rows.Add(
                    item.FirstAppear.ToString("dd/MM/yy HH:mm:ss"),
                    item.LastAppear.ToString("dd/MM/yy HH:mm:ss"),
                    item.CID,
                    item.LAC,
                    item.Location,
                    item.Count);
            }
            list.Add(table);
            #endregion
            columnNames = new List<string> { "Thời gian", "Cell ID", "LAC", "Vị trí"};
            reportname = "Vùng giao thoa";
            reportNames.Add(reportname);
            table = new DataTable();
            table.TableName = reportname;

            foreach (string columnName in columnNames)
            {
                table.Columns.Add(columnName);
            }

            /*
             * CreateReportHeader has to be called after add columns
             */
            rowForColumnName = CreateReportHeader(table, reportname);
            rowNamePos.Add(rowForColumnName);
            table.Rows.Add(columnNames.ToArray());
            foreach (Report item in data[1])
            {
                table.Rows.Add(
                    item.DateTime.ToString("dd/MM/yy HH:mm:ss"),
                    item.CID,
                    item.LAC,
                    item.Location);
            }
            list.Add(table);
            #region OverlapReport

            #endregion OverlapReport
            // Export to file
            DatatableToExcel(saveFilePath, reportNames, list, rowNamePos);
        }

        private static int CreateReportHeader(DataTable table, string reportName)
        {
            if (table.Columns.Count <= 0)
            {
                return 0;
            }
            table.Rows.Add("");
            var typeText = "Loại báo cáo: ";
            var dateText = "Thời gian Tạo: ";
            var dateDetail = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

            table.Rows.Add(typeText + reportName);
            table.Rows.Add(dateText + dateDetail);
            table.Rows.Add("");
            table.Rows.Add("");
            return table.Rows.Count + 1;
        }

        private static void DatatableToExcel(string filePath, List<string> reportNames, List<DataTable> tables, List<int> rowForColumnNames)
        {
            // Creating a new workbook
            var wb = new XLWorkbook();
            for(int i = 0; i < tables.Count; i++)
            {
                //Adding a worksheet
                var ws = wb.Worksheets.Add(reportNames[i]);
                // Insert data
                ws.FirstRow().FirstCell().InsertData(tables[i].Rows);
                // merge row for full text
                var endRowInfoNumber = rowForColumnNames[i] - 1;
                var rangeInfo = ws.Range(1, 1, endRowInfoNumber, tables[i].Columns.Count);
                for (int j = 1; j <= endRowInfoNumber; j++)
                {
                    rangeInfo.Row(j).Merge();
                }
                // Format column name
                var range = ws.Range(rowForColumnNames[i], 1, rowForColumnNames[i], tables[i].Columns.Count);
                range.Style.Fill.BackgroundColor = XLColor.PaleAqua;
                ws.Columns().AdjustToContents();
            }
            wb.SaveAs(filePath);
        }
    }
}
