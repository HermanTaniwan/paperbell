using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PaperbellAppDotNet
{
    public partial class InventoryWindow : Window
    {
        private readonly MainWindow _owner;
        private readonly ObservableCollection<MainWindow.SearchItem> _suggestions = new();
        private System.Collections.Generic.List<MainWindow.SearchItem> _searchIndex = new();
        private int _suggestIndex = -1;
        private bool _suppressSearch;

        public InventoryWindow(MainWindow owner)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            InitializeComponent();
            SuggestList.ItemsSource = _suggestions;
            Loaded += InventoryWindow_Loaded;
        }

        private void InventoryWindow_Loaded(object sender, RoutedEventArgs e)
        {
            RebuildSearchIndex();
            RefreshGrid();
            TxtSearch.Focus();
        }

        private void RebuildSearchIndex()
        {
            _searchIndex = _owner.InventoryGetSearchIndex().ToList();
        }

        private void RefreshGrid()
        {
            var rows = _owner.DbListInventoryItems();
            InvGrid.ItemsSource = rows;
        }

        private static bool AliasMatches(string alias, string input)
        {
            if (string.IsNullOrWhiteSpace(alias) || string.IsNullOrWhiteSpace(input))
                return false;

            var aliasLower = alias.ToLowerInvariant();
            var tokens = input
                .ToLowerInvariant()
                .Split(' ', StringSplitOptions.RemoveEmptyEntries);

            foreach (var token in tokens)
            {
                if (!aliasLower.Contains(token))
                    return false;
            }

            return true;
        }

        private void UpdateSuggestions(string query)
        {
            _suggestions.Clear();
            _suggestIndex = -1;

            if (string.IsNullOrWhiteSpace(query))
                return;

            foreach (var item in _searchIndex
                         .Where(x => AliasMatches(x.Alias, query))
                         .Take(50))
                _suggestions.Add(item);

            if (_suggestions.Count == 0)
            {
                foreach (var item in _searchIndex
                             .Where(x => x.Alias.Contains(query, StringComparison.OrdinalIgnoreCase))
                             .Take(50))
                    _suggestions.Add(item);
            }
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressSearch) return;
            UpdateSuggestions((TxtSearch.Text ?? "").Trim());
        }

        private void TxtSearch_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down)
            {
                e.Handled = true;
                if (_suggestions.Count == 0) return;
                _suggestIndex = Math.Min(_suggestIndex + 1, _suggestions.Count - 1);
                SuggestList.SelectedIndex = _suggestIndex;
                SuggestList.ScrollIntoView(SuggestList.SelectedItem);
                return;
            }

            if (e.Key == Key.Up)
            {
                e.Handled = true;
                if (_suggestions.Count == 0) return;
                _suggestIndex = Math.Max(_suggestIndex - 1, 0);
                SuggestList.SelectedIndex = _suggestIndex;
                SuggestList.ScrollIntoView(SuggestList.SelectedItem);
                return;
            }

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                var item = (SuggestList.SelectedItem as MainWindow.SearchItem)
                           ?? _suggestions.FirstOrDefault();
                if (item?.Map != null)
                    CommitPick(item);
            }
        }

        private void SuggestList_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (SuggestList.SelectedItem is MainWindow.SearchItem item)
                CommitPick(item);
        }

        private void CommitPick(MainWindow.SearchItem item)
        {
            _suppressSearch = true;
            TxtSearch.Text = item.Display;
            _suppressSearch = false;
            _suggestions.Clear();
        }

        private void BtnAddFromOrder_Click(object sender, RoutedEventArgs e)
        {
            var sn = (TxtOrderSn.Text ?? "").Trim();
            if (string.IsNullOrEmpty(sn))
            {
                MessageBox.Show(this, "Isi nomor order Shopee (Order SN).", "Inventory",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var n = _owner.DbInventoryAddFromOrderSn(sn);
                if (n == 0)
                {
                    MessageBox.Show(this,
                        "Order tidak ditemukan di database sync, atau tidak ada baris dengan qty > 0.\nPastikan sudah Shopee Sync dan nomor order benar.",
                        "Inventory",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                MessageBox.Show(this,
                    $"Berhasil menambahkan ke inventory: {n} baris item dari order:\n{sn}",
                    "Inventory",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                _owner.ReloadInventoryCacheFromDb();
                RebuildSearchIndex();
                RefreshGrid();
                TxtOrderSn.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Gagal menambah dari order:\n" + ex.Message, "Inventory",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (_searchIndex.Count == 0)
            {
                MessageBox.Show(this, "Data Map kosong atau belum punya kolom Search Alias. Load Data Map dulu.", "Inventory",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!int.TryParse((TxtQty.Text ?? "").Trim(), out var addQty) || addQty <= 0)
            {
                MessageBox.Show(this, "Qty harus angka positif.", "Inventory", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            MainWindow.SearchItem? picked =
                (SuggestList.SelectedItem as MainWindow.SearchItem)
                ?? _suggestions.FirstOrDefault();

            if (picked?.Map == null)
            {
                var q = (TxtSearch.Text ?? "").Trim();
                picked = _searchIndex.FirstOrDefault(x =>
                    string.Equals(x.Display, q, StringComparison.OrdinalIgnoreCase));
            }

            if (picked?.Map == null)
            {
                picked = _searchIndex.FirstOrDefault(x => AliasMatches(x.Alias, (TxtSearch.Text ?? "").Trim()));
            }

            if (picked?.Map == null)
            {
                MessageBox.Show(this, "Pilih produk dari daftar saran atau ketik alias yang cocok.", "Inventory",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var map = picked.Map;
            if (string.IsNullOrWhiteSpace(map.NoRef) || string.IsNullOrWhiteSpace(map.SKUInduk))
            {
                MessageBox.Show(this, "Baris Data Map ini tidak punya NoRef + SKUInduk â€” tidak bisa dipakai sebagai item_key inventory.",
                    "Inventory", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _owner.DbInventoryAddFromMap(map, addQty);
            _owner.ReloadInventoryCacheFromDb();
            RebuildSearchIndex();
            RefreshGrid();

            _suppressSearch = true;
            TxtSearch.Text = "";
            _suppressSearch = false;
            _suggestions.Clear();
            TxtQty.Text = "1";
        }

        private void BtnDeleteRow_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is not InventoryListItem row)
                return;

            _owner.DbInventoryDelete(row.ItemKey);
            _owner.ReloadInventoryCacheFromDb();
            RefreshGrid();
        }

        private void InvGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Column is not DataGridTextColumn col || col.Header?.ToString() != "Qty")
                return;

            if (e.Row.Item is not InventoryListItem row)
                return;

            if (e.EditingElement is not TextBox tb)
                return;

            if (!int.TryParse((tb.Text ?? "").Trim(), out var q) || q < 0)
            {
                MessageBox.Show(this, "Qty tidak valid.", "Inventory", MessageBoxButton.OK, MessageBoxImage.Warning);
                e.Cancel = true;
                return;
            }

            _owner.DbInventorySetQty(row.ItemKey, q);
            _owner.ReloadInventoryCacheFromDb();
            Dispatcher.BeginInvoke(new Action(RefreshGrid), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape) return;
            e.Handled = true;
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();
    }
}
