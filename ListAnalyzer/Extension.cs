using ListAnalyzer.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ListAnalyzer
{
    public static class Extension
    {
        public static bool IsValid(this Report report)
        {
            return !string.IsNullOrWhiteSpace(report.CID) && !string.IsNullOrWhiteSpace(report.LAC);
        }
    }
}
