using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Messaging;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using WordExporter.UI.Support;

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
            AiKey = StatePersister.Instance.Load<String>("aiKey");
        }

        private TelemetryClient _telemetryClient;

        private String _aiKey;

        public String AiKey
        {
            get
            {
                return _aiKey;
            }
            set
            {
                _telemetryClient = null;
                Set<String>(() => this.AiKey, ref _aiKey, value);
                if (!String.IsNullOrEmpty(_aiKey))
                {
                    _telemetryClient = new TelemetryClient();
                    TelemetryConfiguration.Active.InstrumentationKey = _aiKey;
                }
                StatePersister.Instance.Save<String>("aiKey", _aiKey);
            }
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
            string renderedMessage = logEvent.RenderMessage();
            if (logEvent.Exception != null)
            {
                renderedMessage += "\n" + logEvent.Exception.ToString();
            }
            Logs.Add(new LogViewModel()
            {
                Level = logEvent.Level.ToString(),
                Message = renderedMessage,
            });

            if (_telemetryClient != null)
            {
                if (logEvent.Exception != null)
                {
                    _telemetryClient.TrackException(logEvent.Exception,
                        new Dictionary<String, String>()
                        {
                            ["message"] = renderedMessage,
                        });
                }

                _telemetryClient.TrackTrace(renderedMessage);
            }
        }
    }
}
