using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Messaging;
using Serilog.Events;
using System;
using System.Collections.ObjectModel;

namespace WordExporter.UI.ViewModel
{
    public class LogViewModel : ViewModelBase
    {
        private String _level;

        public String Level
        {
            get
            {
                return _level;
            }
            set
            {
                Set<String>(() => this.Level, ref _level, value);
            }
        }

        private String _message;

        public String Message
        {
            get
            {
                return _message;
            }
            set
            {
                Set<String>(() => this.Message, ref _message, value);
            }
        }
    }

    public class LogViewModelCollector : ViewModelBase
    {
        public LogViewModelCollector()
        {
            Messenger.Default.Register<LogEvent>(this, Append);
        }

        private ObservableCollection<LogViewModel> _logs = new ObservableCollection<LogViewModel>();

        public ObservableCollection<LogViewModel> Logs
        {
            get
            {
                return _logs;
            }
            set
            {
                _logs = value;
                RaisePropertyChanged(nameof(Logs));
            }
        }

        internal void Append(LogEvent logEvent)
        {
            Logs.Add(new LogViewModel()
            {
                Level = logEvent.Level.ToString(),
                Message = logEvent.RenderMessage(),
            });
        }
    }
}
