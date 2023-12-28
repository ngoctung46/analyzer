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
        public string TimeStr { get; set; }
        public string From { get; set; }
        public string IMEI { get; set; }
        public string IMSI { get; set; }
        public string To { get; set; }
        public string CID { get; set; }
        public string LAC { get; set; }
        public string Location { get; set; }
        public int Count { get; set; }
        public string Duration { get; set; }
        public int NetworkCode { get; set; }
        public DateTime Time
        {
            get
            {
                DateTime dateTime = new DateTime();
                var success = false;
                switch(NetworkCode)
                {
                    case 1: case 4: success = DateTime.TryParseExact(TimeStr, "dd/MM/yyyy HH:mm:ss", 
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime);break;
                    case 2:
                    {
                        success = DateTime.TryParseExact(TimeStr, "yyMMdd-HHmmss", CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out dateTime); break;
                        }
                        
                }
                if (!success)
                {
                    DateTime.TryParseExact(TimeStr, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out dateTime);
                };
                return dateTime;
            }
        }
        public DateTime FirstAppear { get; set; }
        public DateTime LastAppear { get; set; }
    }
}
