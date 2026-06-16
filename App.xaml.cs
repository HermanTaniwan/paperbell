using System.IO;
using System.Windows;
using PaperbellAppDotNet;

namespace Paperbell_App
{
    public partial class App : Application
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PaperbellAppDotNet", "startup_error.log");

        private static readonly string TraceLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PaperbellAppDotNet", "startup_trace.log");

        public static void Trace(string step)
        {
            try
            {
                var dir = Path.GetDirectoryName(TraceLogPath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.AppendAllText(TraceLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] {step}\n");
            }
            catch { }
        }

        public App()
        {
            this.DispatcherUnhandledException += (s, e) =>
            {
                var msg = e.Exception.ToString();
                WriteLog(msg);

                // Kalau belum ada window yang terbuka = startup fatal, langsung shutdown
                if (Current.Windows.Count == 0)
                {
                    MessageBox.Show(msg, "Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    e.Handled = true;
                    Current.Shutdown(1);
                }
                else
                {
                    MessageBox.Show(msg, "Unhandled UI Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    e.Handled = true;
                }
            };

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                var msg = e.ExceptionObject?.ToString() ?? "(null)";
                WriteLog(msg);
                MessageBox.Show(msg, "Unhandled Domain Error", MessageBoxButton.OK, MessageBoxImage.Error);
            };
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                try { File.Delete(TraceLogPath); } catch { }
                Trace("OnStartup begin");
                base.OnStartup(e);
                Trace("base.OnStartup done");
                var window = new MainWindow();
                Trace("MainWindow ctor done");
                MainWindow = window;
                window.Show();
                Trace("window.Show done");
            }
            catch (Exception ex)
            {
                var msg = ex.ToString();
                WriteLog(msg);
                Trace("OnStartup EXCEPTION: " + msg);
                MessageBox.Show(msg, "Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(1);
            }
        }

        private static void WriteLog(string message)
        {
            try
            {
                var dir = Path.GetDirectoryName(LogPath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.WriteAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}]\n{message}\n");
            }
            catch { }
        }
    }
}
