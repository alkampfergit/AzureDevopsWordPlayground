using Serilog;
using Serilog.Exceptions;
using System.ComponentModel;
using System.Windows;
using WordExporter.UI.Support;

namespace WordExporter.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            Log.Logger = new LoggerConfiguration()
                .Enrich.WithExceptionDetails()
                .MinimumLevel.Debug()
                .WriteTo.File(
                    "logs.txt",
                     rollingInterval: RollingInterval.Day
                )
                .WriteTo.File(
                    "errors.txt",
                     rollingInterval: RollingInterval.Day,
                     restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Error
                )
                .CreateLogger();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            StatePersister.Instance.Persist();
            base.OnClosing(e);
        }
    }
}
