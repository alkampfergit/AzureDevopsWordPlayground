using System;
using Serilog;
using Serilog.Exceptions;
using System.ComponentModel;
using System.Security;
using System.Windows;
using System.Windows.Controls;
using WordExporter.UI.Support;
using WordExporter.UI.ViewModel;

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
                    "logs\\logs.txt",
                     rollingInterval: RollingInterval.Day
                )
                .WriteTo.File(
                    "logs\\errors.txt",
                     rollingInterval: RollingInterval.Day,
                     restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Error
                )
                .WriteTo.Sink(new LogInterceptorSink())
                .CreateLogger();

            var lv = new LogWindows();
            lv.Show();

            Log.Information("Word exporter started!");
        }

        protected override void OnInitialized(EventArgs e)
        {
            base.OnInitialized(e);

            var savedPassword = StatePersister.Instance.Load<String>("password");
            if (!String.IsNullOrEmpty(savedPassword))
            {
                var decrypted = EncryptionUtils.Decrypt(savedPassword);
                PasswordBox.Password = decrypted;
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            StatePersister.Instance.Persist();
            base.OnClosing(e);
        }

        private void PasswordBox_OnPasswordChanged(object sender, RoutedEventArgs e)
        {
            var pb = (PasswordBox) sender;
            PasswordBoxPassword.SetEncryptedPassword(pb, pb.SecurePassword);
        }
    }
}
