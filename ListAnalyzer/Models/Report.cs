using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListAnalyzer.Models
{
    public class Report
    {
        public string TG { get; set; }
        public string CID { get; set; }
        public string LAC { get; set; }
        public string Location { get; set; }
        public int Count { get; set; }
        public string Duration { get; set; }
        public DateTime Time
        {
            get
            {
                DateTime dateTime = new DateTime();
                var success = DateTime.TryParseExact(TG, "dd/MM/yy HH:mm:ss", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out dateTime);
                if (!success)
                {
                    DateTime.TryParseExact(TG, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out dateTime);
                };
                return dateTime;
            }
        }
        public DateTime FirstAppear { get; set; }
        public DateTime LastAppear { get; set; }
    }
}
