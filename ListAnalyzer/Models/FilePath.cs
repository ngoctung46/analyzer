using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListAnalyzer.Models
{
    public class FilePath : ReactiveObject
    {
        private string importPath;
        public string ImportPath
        {
            get => importPath;
            set => this.RaiseAndSetIfChanged(ref importPath, value);
        }

        private string reportPath;
        public string ReportPath
        {
            get => reportPath;
            set => this.RaiseAndSetIfChanged(ref reportPath, value);
        }
    }
}
