using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListAnalyzer.Models
{
    public class Report
    {
        public string Time { get; set; }
        public string CID { get; set; }
        public string LAC { get; set; }
        public string Location { get; set; }
        public int Count { get; set; }
        public DateTime DateTime
        {
            get
            {
                var success = DateTime.TryParseExact(Time, "dd/MM/yy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateTime);
                if (success) return dateTime;
                return new DateTime();
            }
        }
        public DateTime FirstAppear { get; set; }
        public DateTime LastAppear { get; set; }
    }
}
