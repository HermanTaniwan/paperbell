using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace PaperbellAppDotNet
{
    public sealed class OrderPackDetailLine
    {
        public int Index { get; init; }
        public string ItemName { get; init; } = "";
        public string ModelName { get; init; } = "";
        public int Qty { get; init; }
        public bool IsPrinted { get; init; }
        public string PrintStatusLabel => IsPrinted ? "Sudah dicetak" : "Belum dicetak";
        public string PrintSidesLabel { get; init; } = "";
    }

    public sealed class OrderPackDetailInfo
    {
        public string OrderSn { get; init; } = "";
        public string ResiStatusText { get; init; } = "";
        public string PrintSummaryText { get; init; } = "";
        public string PackSummaryText { get; init; } = "";
        public string NotesText { get; init; } = "";
        public IReadOnlyList<OrderPackDetailLine> Lines { get; init; } = [];
    }

    public partial class OrderPackDetailWindow : Window
    {
        public OrderPackDetailWindow(OrderPackDetailInfo info)
        {
            InitializeComponent();
            TxtOrderSn.Text = "Order: " + info.OrderSn;
            TxtResiStatus.Text = info.ResiStatusText;
            TxtPrintSummary.Text = info.PrintSummaryText;
            TxtPackSummary.Text = info.PackSummaryText;
            TxtOrderNotes.Text = string.IsNullOrWhiteSpace(info.NotesText)
                ? "Tidak ada catatan yang tersimpan untuk order ini."
                : info.NotesText;
            LinesGrid.ItemsSource = info.Lines;
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
