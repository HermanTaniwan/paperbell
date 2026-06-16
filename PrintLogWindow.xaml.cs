using System.IO;
using System.Windows;

namespace PaperbellAppDotNet
{
    public partial class PrintLogWindow : Window
    {
        private readonly string _logPath;

        public PrintLogWindow(string logPath)
        {
            InitializeComponent();
            _logPath = logPath;
            Reload();
        }

        private void Reload()
        {
            try
            {
                LogBox.Text = File.Exists(_logPath)
                    ? File.ReadAllText(_logPath)
                    : "(No log entries yet.)";
            }
            catch (System.Exception ex)
            {
                LogBox.Text = "Could not read log: " + ex.Message;
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e) => Reload();

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
