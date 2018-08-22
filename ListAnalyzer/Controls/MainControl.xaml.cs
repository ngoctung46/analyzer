using ListAnalyzer.ViewModels;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;
using System.Windows.Forms;
using MahApps.Metro.Controls.Dialogs;


namespace ListAnalyzer.Controls
{
    /// <summary>
    /// Interaction logic for MainControl.xaml
    /// </summary>
    public partial class MainControl : System.Windows.Controls.UserControl, IViewFor<MainControlViewModel>
    {
        private readonly string SAVE_FILE_FILTER = "Excel files(.xlsx)| *.xlsx";
        private readonly string SAVE_FILE_DEFAULT_EXTENSION = ".xlsx";

        public MainControl()
        {
            InitializeComponent();
            ViewModel = new MainControlViewModel();
            this.WhenActivated(Bind);
        }

        private void Bind(Action<IDisposable> d)
        {

            d(this.OneWayBind(ViewModel, vm => vm.ImportPath, v => v.ImportTextBox.Text));
            d(this.OneWayBind(ViewModel, vm => vm.ReportPath, v => v.ReportTextBox.Text));
            d(this.BindCommand(ViewModel, vm => vm.ImportCommand, v => v.ImportButton));
            d(this.BindCommand(ViewModel, vm => vm.ReportCommand, v => v.ReportButton));
            d(this.BindCommand(ViewModel, vm => vm.SubmitCommand, v => v.AnalyzeButton));
            d(ViewModel.ImportInteraction.RegisterHandler(interaction => interaction.SetOutput(GetImportPath())));
            d(ViewModel.ReportInteraction.RegisterHandler(interaction => interaction.SetOutput(GetReportPath())));
            d(this.WhenAnyObservable(x => x.ViewModel.SubmitCommand).Subscribe(_ =>
            {
                MessageTextBlock.Text = $"File báo cáo đã được tạo thành công tại: {ViewModel.ReportPath}";
            }));
        }

        internal string GetImportPath()
        {
            OpenFileDialog file = new OpenFileDialog();//open dialog to choose file
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)//if there is a file choosen by the user
            {
                var filePath = file.FileName;//get the path of the file
                var fileExt = Path.GetExtension(filePath);//get the file extension

                if (fileExt.CompareTo(".xls") != 0 || fileExt.CompareTo(".xlsx") != 0)
                {
                    return file.FileName;
                }
                else
                {
                    System.Windows.MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning");//custom messageBox to show error
                    return String.Empty;
                }
            }
            else
            {
                return String.Empty;
            }
        }

        internal string GetReportPath()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = SAVE_FILE_FILTER;
            dialog.DefaultExt = SAVE_FILE_DEFAULT_EXTENSION;
            dialog.Title = "Chọn đường dẫn lưu báo cáo";
            var result = dialog.ShowDialog();
            if (result == DialogResult.OK && dialog.FileName != "")
            {
                return dialog.FileName;
            }
            return String.Empty;
        }

        public static readonly DependencyProperty ViewModelProperty =
            DependencyProperty.Register("ViewModel", typeof(MainControlViewModel), typeof(MainControl));

        public MainControlViewModel ViewModel
        {
            get => GetValue(ViewModelProperty) as MainControlViewModel;
            set => SetValue(ViewModelProperty, value);
        }
        object IViewFor.ViewModel
        {
            get => ViewModel;
            set => ViewModel = value as MainControlViewModel;
        }
    }
}
