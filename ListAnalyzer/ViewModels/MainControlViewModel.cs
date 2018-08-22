using ListAnalyzer.Models;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Reactive;
using System.Text;
using System.Threading.Tasks;

namespace ListAnalyzer.ViewModels
{
    public class MainControlViewModel : ReactiveObject
    {
        private FilePath filePath;
        public FilePath FilePath
        {
            get => filePath;
            set => this.RaiseAndSetIfChanged(ref filePath, value);
        }

        private List<List<Report>> reports = new List<List<Report>>();
        private readonly ObservableAsPropertyHelper<string> importPath;
        private readonly ObservableAsPropertyHelper<string> reportPath;
        public string ImportPath => importPath.Value;
        public string ReportPath => reportPath.Value;
        public Interaction<string, string> ImportInteraction { get; }
        public Interaction<string, string> ReportInteraction { get; }
        public ReactiveCommand<Unit, Unit> ImportCommand { get; protected set; }
        public ReactiveCommand<Unit, Unit> ReportCommand { get; protected set; }
        public ReactiveCommand<Unit, Unit> SubmitCommand { get; protected set; }
        public MainControlViewModel()
        {
            ImportInteraction = new Interaction<string, string>();
            ReportInteraction = new Interaction<string, string>();
            ImportCommand = ReactiveCommand.Create(Import);
            ReportCommand = ReactiveCommand.Create(Report);
            FilePath = new FilePath();
            importPath = FilePath.WhenAnyValue(x => x.ImportPath).ToProperty(this, x => x.ImportPath);
            reportPath = FilePath.WhenAnyValue(x => x.ReportPath).ToProperty(this, x => x.ReportPath);
            var canSubmit = FilePath
                .WhenAnyValue(x => x.ImportPath, y => y.ReportPath, (x, y)
                    => !string.IsNullOrWhiteSpace(x) && !string.IsNullOrWhiteSpace(y));
            SubmitCommand = ReactiveCommand.Create(Submit, canSubmit);
        }

        private void Submit()
        {
            Console.WriteLine($"Import path: {ImportPath}");
            Console.WriteLine($"Report path: {ReportPath}");
            var list = HelperFunctions.ExcelToList(ImportPath);
            var duplicateList = HelperFunctions.CountDuplicate(list);
            reports.Add(duplicateList);
            var overlapList = HelperFunctions.FindOverlap(list);
            reports.Add(overlapList);
            HelperFunctions.ExportReport(ReportPath, reports);
        }

        private void Report()
        {
            ReportInteraction.Handle("Chọn folder lưu báo cáo")
                .Subscribe(path =>
                {
                    if (!String.IsNullOrWhiteSpace(path)) FilePath.ReportPath = path;
                });
        }

        private void Import()
        {
            ImportInteraction.Handle("Chọn file để phân tích")
                 .Subscribe(path =>
                 {
                     if (!String.IsNullOrWhiteSpace(path)) FilePath.ImportPath = path;
                 });
        }
    }
}
