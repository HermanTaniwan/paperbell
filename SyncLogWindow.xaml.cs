using System;
using System.Windows;

namespace PaperbellAppDotNet
{
    public partial class SyncLogWindow : Window
    {
        public SyncLogWindow()
        {
            InitializeComponent();
        }

        public void AppendLog(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return;
            LogBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {line}\r\n");
            LogBox.ScrollToEnd();
        }

        public void SetProgress(double percent, string? label = null)
        {
            if (percent < 0) percent = 0;
            if (percent > 100) percent = 100;

            Bar.Value = percent;
            TxtProgress.Text = label ?? $"{percent:0}%";
        }



        public event Action? CancelRequested;
        private bool _running = true;

        public void SetDone(string finalTitle = "Sync completed ✅")
        {
            TxtTitle.Text = finalTitle;
            _running = false;
            BtnClose.Content = "Close";
            BtnClose.IsEnabled = true;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            if (_running)
            {
                CancelRequested?.Invoke();
                BtnClose.IsEnabled = false;
                BtnClose.Content = "Cancelling...";
                return;
            }
            Close();
        }
    }
}
