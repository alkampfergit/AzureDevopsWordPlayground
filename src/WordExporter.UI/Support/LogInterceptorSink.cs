using GalaSoft.MvvmLight.Messaging;
using Serilog.Core;
using Serilog.Events;

namespace WordExporter.UI.Support
{
    public class LogInterceptorSink : ILogEventSink
    {
        public void Emit(LogEvent logEvent)
        {
            Messenger.Default.Send(logEvent);
        }
    }
}
