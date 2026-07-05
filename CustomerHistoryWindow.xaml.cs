using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace PaperbellAppDotNet
{
    public sealed class CustomerHistoryLine
    {
        public string OrderSn { get; init; } = "";
        public string OrderCreatedText { get; init; } = "";
        public string ItemName { get; init; } = "";
        public string ModelName { get; init; } = "";
        public int Qty { get; init; }
    }

    public sealed class CustomerHistoryInfo
    {
        public string CustomerUsername { get; init; } = "";
        public string SummaryText { get; init; } = "";
        public IReadOnlyList<CustomerHistoryLine> Lines { get; init; } = [];
    }

    public partial class CustomerHistoryWindow : Window
    {
        public CustomerHistoryWindow(CustomerHistoryInfo info)
        {
            InitializeComponent();
            TxtCustomer.Text = string.IsNullOrWhiteSpace(info.CustomerUsername)
                ? "Customer tidak diketahui"
                : "Customer: " + info.CustomerUsername;
            TxtSummary.Text = string.IsNullOrWhiteSpace(info.SummaryText)
                ? "Belum ada riwayat order yang bisa ditampilkan."
                : info.SummaryText;
            HistoryGrid.ItemsSource = info.Lines;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape) return;
            e.Handled = true;
            Close();
        }
    }
}
