using ListAnalyzer.Models;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reactive;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

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

        private Network selectedNetwork = null;
        public Network SelectedNetwork
        {
            get => selectedNetwork;
            set => this.RaiseAndSetIfChanged(ref selectedNetwork, value);
        }

        public ObservableCollection<Network> Networks = new ObservableCollection<Network>()
        {
            new Network() { NetworkCode = 1, NetworkName = "Mobifone"},
            new Network() { NetworkCode = 2, NetworkName = "Vinaphone"},
            new Network(){ NetworkCode = 4, NetworkName = "Viettel"},
            new Network() { NetworkCode = 5, NetworkName = "Vienamobile"}
        };

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
            canSubmit = this.WhenAny(x => x.SelectedNetwork, x => x != null);
            SubmitCommand = ReactiveCommand.Create(Submit, canSubmit);
        }

        private void Submit()
        {
            try
            {
                var columns = HelperFunctions.GetColumnsByNetwork(SelectedNetwork.NetworkCode);
                var list = HelperFunctions.ReadExcel(ImportPath, columns);
                if(list.Count <= 0) { throw new Exception("File excel không có dữ liệu hoặc không đúng định dạng cột"); }
                var duplicateList = list.CountDuplicate();
                reports.Add(duplicateList);
                var overlapList = list.FindOverlap();
                reports.Add(overlapList);
                var mostDurationList = list.FindMostDuration();
                reports.Add(mostDurationList);
                var dayList = list.FindInRange(startHour: 7, endHour: 17);
                reports.Add(dayList);
                var eveningList = list.FindInRange(startHour: 17, endHour: 22);
                reports.Add(eveningList);
                var nightList = list.FindInRange();
                reports.Add(nightList);
                var imeiList = list.GroupBy(x => x.IMEI).Select(x => x.First()).ToList();
                var imsiList = list.GroupBy(x => x.IMSI).Select(x => x.First()).ToList();
                reports.Add(imeiList.Union(imsiList).ToList());
                var contactList = list.CountContact();
                reports.Add(contactList);

                HelperFunctions.ExportReport(ReportPath, reports);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox
                    .Show(ex.Message.ToString(), "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }

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
