// ✅ ADD: Excel reader
using ExcelDataReader;
using Microsoft.Win32;
using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing; // PdfiumViewer renders to System.Drawing.Bitmap
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Printing;
using System.Runtime.CompilerServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Json;
using System.Net;
using System.Diagnostics;
using Microsoft.Data.Sqlite;
using SQLitePCL;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Data;

namespace PaperbellAppDotNet
{
    public class SyncProgress
    {
        // untuk progress bar (0..100)
        public double Percent { get; set; }

        // teks progress: "10/120 items" atau "Page 2"
        public string? Label { get; set; }

        // log line (optional)
        public string? Log { get; set; }
    }

    public partial class MainWindow : Window
    {
        // Options for per-row settings (used by DataGrid ComboBoxes)
        public Array DuplexOptions => Enum.GetValues(typeof(DuplexMode));
        public Array PaperOptions => Enum.GetValues(typeof(PaperPreset));

        public ObservableCollection<JobRow> Rows { get; } = new();
        public ObservableCollection<string> Printers { get; } = new();
        public ObservableCollection<ResiRow> ResiRows { get; } = new();

        private double _zoom = 1.0;
        private (string file, int page, int total)? _full;
        private readonly Dictionary<Guid, CancellationTokenSource> _jobCts = new();
        private CancellationTokenSource? _previewCts;
        private int _previewReqId = 0;
        private JobRow? _previewRow;
        private JobRow? _lastUserPickedForPreview;
        private bool _suppressPreviewWhileScrolling;
        private bool _queueGridPreviewDeferredWhileScrolling;
        private System.Windows.Threading.DispatcherTimer? _scrollStopTimer;
        private bool _pdfPreviewAvailable = true;
        private bool _pdfPreviewWarningShown = false;

        private bool _mouseCommitRequested = false;
        private string _typedText = "";
        private TextBox? _searchTextBox;
        private bool _isKeyboardNavigating = false;
        private int _suggestIndex = -1;


        public ICollectionView AllRowsView { get; private set; }
        public ICollectionView NotPrintedView { get; private set; }
        public ICollectionView PrintedView { get; private set; }

        [System.Runtime.InteropServices.DllImport("kernel32", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        private static void NormalizePageRangeForPrint(JobRow r)
        {
            if (r == null) return;

            var raw = (r.Pages ?? "").Trim();
            if (string.IsNullOrWhiteSpace(raw))
                raw = (r.PageRange ?? "").Trim();

            if (string.IsNullOrWhiteSpace(raw))
                return;

            var m = Regex.Match(raw, @"^\s*(\d+)?\s*(?:-\s*(\d+)?)?\s*$");
            if (!m.Success) return;

            int? a = null, b = null;
            if (m.Groups[1].Success && int.TryParse(m.Groups[1].Value, out var aa)) a = aa;
            if (m.Groups[2].Success && int.TryParse(m.Groups[2].Value, out var bb)) b = bb;

            int from = Math.Max(1, a ?? 1);
            int to;

            if (a.HasValue && !raw.Contains("-")) to = from;         // "1"
            else if (!a.HasValue && b.HasValue) to = Math.Max(1, b.Value);
            else if (a.HasValue && !b.HasValue) to = 0;              // "1-"
            else to = Math.Max(0, b ?? from);                        // "1-20"

            if (to > 0 && to < from) (from, to) = (to, from);

            r.PageFrom = from;
            r.PageTo = to;
        }

        private void InitPdfPreviewAvailability()
        {
            try
            {
                var baseDir = AppContext.BaseDirectory;

                // PdfiumViewer.Native.* NuGet copies pdfium.dll under x64\; also support flat layout.
                var paths = new[]
                {
                    Path.Combine(baseDir, "pdfium.dll"),
                    Path.Combine(baseDir, "x64", "pdfium.dll"),
                    Path.Combine(baseDir, "pdfium_x64.dll"),
                    Path.Combine(baseDir, "pdfium_x86.dll"),
                };

                bool ok = false;
                foreach (var path in paths)
                {
                    if (!File.Exists(path))
                        continue;
                    try
                    {
                        var h = LoadLibrary(path);
                        if (h != IntPtr.Zero)
                        {
                            ok = true;
                            break;
                        }
                    }
                    catch { }
                }

                if (!ok)
                {
                    foreach (var dll in new[] { "pdfium.dll", "pdfium_x64.dll", "pdfium_x86.dll" })
                    {
                        try
                        {
                            var h = LoadLibrary(dll);
                            if (h != IntPtr.Zero)
                            {
                                ok = true;
                                break;
                            }
                        }
                        catch { }
                    }
                }

                _pdfPreviewAvailable = ok;
                LogPrint(ok ? "PDF preview available." : "PDF preview disabled at startup: native pdfium dll not found.");
            }
            catch (Exception ex)
            {
                _pdfPreviewAvailable = false;
                LogPrint("PDF preview disabled at startup: " + ex);
            }
        }




        private void DisablePdfPreview(JobRow? row, Exception ex)
        {
            _pdfPreviewAvailable = false;

            try
            {
                _previewCts?.Cancel();
            }
            catch { }

            if (row != null)
            {
                row.Status = "Preview unavailable";
                row.Percent = 0;
            }

            Thumbs.ItemsSource = null;
            HidePreviewLoading();

            if (Modal.Visibility == Visibility.Visible)
                Modal.Visibility = Visibility.Collapsed;

            // optional: log ke file / debug, tapi jangan popup
            LogPrint("PDF preview disabled: " + ex);
        }

        private static bool IsPdfiumFatal(Exception ex)
        {
            if (ex is DllNotFoundException ||
                ex is BadImageFormatException ||
                ex is TypeInitializationException)
                return true;

            var msg = ex.ToString();

            return msg.Contains("pdfium", StringComparison.OrdinalIgnoreCase) ||
                   msg.Contains("Unable to load DLL", StringComparison.OrdinalIgnoreCase) ||
                   msg.Contains("The type initializer", StringComparison.OrdinalIgnoreCase);
        }


        // ===== Shopee UI pagination (from DB) =====
        private enum ShopeeTabFilter
        {
            All,
            NotPrinted,
            Printed
        }

        // filter aktif saat ini (default: NotPrinted, supaya tab default = Not Printed)
        private ShopeeTabFilter _currentTabFilter = ShopeeTabFilter.NotPrinted;

        // state page per tab
        private int _pageIndexAll = 0;
        private int _pageIndexNotPrinted = 0;
        private int _pageIndexPrinted = 0;

        // state page & total untuk tab yang lagi aktif
        private int _shopeePageIndex = 0;          // 0-based
        private const int ShopeePageSize = 10;     // sesuai request kamu
        private int _shopeeTotalItems = 0;

        private enum ResiTabFilter
        {
            All,
            NotPrintedResi,
            PrintedResi
        }

        private ResiTabFilter _resiTabFilter = ResiTabFilter.NotPrintedResi;
        private int _resiPageIndexAll;
        private int _resiPageIndexNotPrinted;
        private int _resiPageIndexPrinted;
        private int _resiPageIndex;
        private int _resiTotalItems;
        private const int ResiPageSize = 25;
        private const int ResiListDaysBack = 30;
        private const double ResiLabelPdfPrintScale = 0.75;
        /// <summary>Substring untuk default combo printer label (mis. "EPSON L3210 Series").</summary>
        private const string ResiDefaultPrinterNameContains = "L3210";
        private readonly JobRow _resiPreviewStub = new() { OrderNo = "Resi", Status = "Ready" };

        private async Task<bool> CheckShopeeSessionAsync()
        {
            try
            {
                // ambil window kecil agar cepat
                var now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
                var timeFrom = now - 60; // 1 menit terakhir

                var extra = new Dictionary<string, string>
                {
                    ["time_range_field"] = "create_time",
                    ["time_from"] = timeFrom.ToString(),
                    ["time_to"] = now.ToString(),
                    ["page_size"] = "1",
                    ["page_no"] = "1",
                    ["order_status"] = "PROCESSED"
                };

                // progress dummy (boleh null kalau function kamu mengizinkan)
                var progress = new Progress<SyncProgress>(_ => { });

                using var doc = await GetShopApiWithLogAsync(
                    "/api/v2/order/get_order_list",
                    extra,
                    progress,
                    CancellationToken.None);

                // kalau sukses parse JSON, berarti token OK
                return true;
            }
            catch (Exception ex)
            {
                var msg = ex.Message ?? "";

                // kalau token invalid / refresh expired -> anggap tidak valid
                if (msg.Contains("invalid_access_token", StringComparison.OrdinalIgnoreCase) ||
                    msg.Contains("invalid_acceess_token", StringComparison.OrdinalIgnoreCase) ||
                    msg.Contains("refresh_token_expired", StringComparison.OrdinalIgnoreCase))
                    return false;

                // error lain (misal internet down) -> kamu bisa pilih:
                // return true (biar tidak force disconnect) atau false (biar user reconnect)
                // aku saran: jangan force disconnect kalau cuma network error
                if (msg.Contains("timed out", StringComparison.OrdinalIgnoreCase) ||
                    msg.Contains("No such host", StringComparison.OrdinalIgnoreCase))
                    return true;

                return false;
            }
        }

        private void DisconnectShopee(string? reason = null)
        {
            // Clear state
            _accessToken = null;
            _refreshToken = null;
            _shopId = 0;

            // kalau ada field ini di project kamu
            _accessTokenExpiredAt = DateTimeOffset.MinValue;

            // simpan supaya pas restart tidak kebaca connected lagi
            SaveAppStateToDb();

            // Update UI di thread UI
            Dispatcher.Invoke(() =>
            {
                UpdateShopeeUi();

                if (!string.IsNullOrWhiteSpace(reason))
                {
                    MessageBox.Show(
                        reason,
                        "Shopee",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            });
        }

        private void PrintedToggle_Click(object sender, RoutedEventArgs e)
{
            if (sender is not Button btn) return;
            if (btn.DataContext is not JobRow row) return;

            // Only Shopee rows (stored in DB) can be toggled.
            if (row.OrderProcessId <= 0) return;

            var newValue = !row.IsPrinted;

            // Persist by PK id so it always hits the correct row.
            DbSetPrintedById(row.OrderProcessId, newValue);

            row.IsPrinted = newValue;

            // 🔽 reload page sekarang di tab aktif
            LoadShopeePageFromDb(_shopeePageIndex);
        }




        // =====================
        // ✅ ADD: Search + AutoSuggest dari DataMap
        // =====================
        public ObservableCollection<SearchItem> SearchSuggestions { get; } = new();
        private List<SearchItem> _searchIndex = new();
        private bool _suppressSearchUpdate;


        private static string BuildSearchDisplay(DataMapRow m)
        {
            // Biar enak dicari: NoRef + Variasi + SKUInduk + Printer + Page
            var parts = new List<string>();

            if (!string.IsNullOrWhiteSpace(m.NoRef)) parts.Add(m.NoRef.Trim());
            if (!string.IsNullOrWhiteSpace(m.Variasi)) parts.Add(m.Variasi.Trim());
            if (!string.IsNullOrWhiteSpace(m.SKUInduk)) parts.Add(m.SKUInduk.Trim());

            // info tambahan (optional)
            if (!string.IsNullOrWhiteSpace(m.Printer)) parts.Add($"[{m.Printer.Trim()}]");
            if (!string.IsNullOrWhiteSpace(m.FilePath)) parts.Add(Path.GetFileName(m.FilePath.Trim()));

            return string.Join(" | ", parts);
        }

        private void SearchBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // menandai bahwa perubahan selection berikutnya datang dari mouse
            _mouseCommitRequested = true;
        }

        private void RebuildSearchIndex()
        {
            _searchIndex = _dataMap.Values
                .Where(m => !string.IsNullOrWhiteSpace(m.SearchAlias))
                .Select(m => new SearchItem
                {
                    Map = m,
                    Alias = m.SearchAlias!.Trim(),
                    Display = m.SearchAlias!.Trim()
                })
                .DistinctBy(x => x.Display, StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x.Display, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void UpdateShopeeUi()
        {
            // contoh: kalau kamu punya Button bernama BtnConnectShopee dan TextBlock bernama ShopeeStatusText
            if (IsConnected())
            {
                BtnConnectShopee.Content = "Shopee Connected";
                BtnConnectShopee.Background = System.Windows.Media.Brushes.SeaGreen;
                BtnConnectShopee.Foreground = System.Windows.Media.Brushes.White;

                // jangan disable, biar style gak dioverride WPF
                BtnConnectShopee.IsEnabled = true;
                BtnConnectShopee.IsHitTestVisible = false;   // ✅ tidak bisa ditekan
                BtnConnectShopee.Focusable = false;
                BtnConnectShopee.Opacity = 1;
            }
            else
            {
                BtnConnectShopee.Content = "Connect Shopee";
                BtnConnectShopee.ClearValue(Button.BackgroundProperty);
                BtnConnectShopee.ClearValue(Button.ForegroundProperty);
                BtnConnectShopee.IsEnabled = true;
                BtnConnectShopee.IsHitTestVisible = true;
                BtnConnectShopee.Focusable = true;
                BtnConnectShopee.Opacity = 1;

            }
        }



        private void CommitSearchToJob()
        {
            if (_dataMap.Count == 0)
            {
                MessageBox.Show("DataMap belum diload. Klik 'Load Data Map (XLSX)…' dulu.");
                return;
            }

            var picked =
                (SuggestList.SelectedItem as SearchItem) ??
                SearchSuggestions.FirstOrDefault();

            if (picked?.Map == null)
                return;

            AddJobFromMap(picked.Map);

            // reset search UI
            _suppressSearchUpdate = true;
            SearchBox.Text = "";
            SuggestList.SelectedItem = null;
            SearchSuggestions.Clear();
            SuggestPopup.IsOpen = false;
            _suppressSearchUpdate = false;
        }

        private void SearchBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (SearchBox.Template.FindName("PART_EditableTextBox", SearchBox) is TextBox tb)
            {
                _searchTextBox = tb;

                tb.TextChanged -= SearchBox_InnerTextChanged;
                tb.TextChanged += SearchBox_InnerTextChanged;
            }
        }

        private void SearchBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!SuggestPopup.IsOpen && (e.Key == Key.Down || e.Key == Key.Up))
            {
                // kalau popup belum kebuka tapi ada suggestion, buka
                if (SearchSuggestions.Count > 0)
                    SuggestPopup.IsOpen = true;
            }

            if (e.Key == Key.Down)
            {
                e.Handled = true;
                if (SearchSuggestions.Count == 0) return;

                _suggestIndex++;
                if (_suggestIndex >= SearchSuggestions.Count) _suggestIndex = SearchSuggestions.Count - 1;

                SuggestList.SelectedIndex = _suggestIndex;
                SuggestList.ScrollIntoView(SuggestList.SelectedItem);
                return;
            }

            if (e.Key == Key.Up)
            {
                e.Handled = true;
                if (SearchSuggestions.Count == 0) return;

                _suggestIndex--;
                if (_suggestIndex < 0) _suggestIndex = 0;

                SuggestList.SelectedIndex = _suggestIndex;
                SuggestList.ScrollIntoView(SuggestList.SelectedItem);
                return;
            }

            if (e.Key == Key.Enter)
            {
                e.Handled = true;

                var item = (SuggestList.SelectedItem as SearchItem)
                           ?? (SearchSuggestions.Count > 0 ? SearchSuggestions[0] : null);

                if (item != null)
                    CommitSearchItem(item);

                return;
            }

            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                ClearSearchBox();
                return;
            }
        }


        private void ClearSearchBox()
        {
            _suppressSearchUpdate = true;
            SearchBox.Text = "";
            _suppressSearchUpdate = false;

            SearchSuggestions.Clear();
            _suggestIndex = -1;
            SuggestPopup.IsOpen = false;

            SearchBox.Focus();
        }

        private void MoveSuggestionSelection(int delta)
        {
            if (_searchTextBox == null) return;
            if (SearchSuggestions.Count == 0) return;

            // simpan text yang user ketik
            _typedText = _searchTextBox.Text ?? "";

            SuggestPopup.IsOpen = true;

            int idx = SuggestList.SelectedIndex;

            if (idx < 0)
                idx = (delta > 0) ? 0 : SearchSuggestions.Count - 1;
            else
            {
                idx += delta;
                if (idx < 0) idx = 0;
                if (idx >= SearchSuggestions.Count) idx = SearchSuggestions.Count - 1;
            }

            // set highlight
            _suppressSearchUpdate = true;
            SuggestList.SelectedIndex = idx;
            _suppressSearchUpdate = false;

            // ✅ penting: jangan biarkan ComboBox mengganti text
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_searchTextBox == null) return;

                _suppressSearchUpdate = true;
                _searchTextBox.Text = _typedText;
                _searchTextBox.SelectionStart = _typedText.Length;
                _searchTextBox.SelectionLength = 0;
                _searchTextBox.CaretIndex = _typedText.Length;
                _suppressSearchUpdate = false;
            }), System.Windows.Threading.DispatcherPriority.Input);
        }

        private bool AliasMatches(string alias, string input)
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

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressSearchUpdate) return;

            var q = (SearchBox.Text ?? "").Trim();
            SearchSuggestions.Clear();
            _suggestIndex = -1;

            if (string.IsNullOrWhiteSpace(q))
            {
                SuggestPopup.IsOpen = false;
                return;
            }

            foreach (var item in _searchIndex
                .Where(x => AliasMatches(x.Alias, q))
                .Take(50))
            {
                SearchSuggestions.Add(item);
            }

            SuggestPopup.IsOpen = SearchSuggestions.Count > 0;
            if (SuggestPopup.IsOpen)
            {
                SuggestList.SelectedIndex = -1;
                SuggestList.ScrollIntoView(SuggestList.SelectedItem);
            }
        }

        private void SuggestList_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (SuggestList.SelectedItem is SearchItem item)
            {
                CommitSearchItem(item);
            }
        }



        private void SearchBox_InnerTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressSearchUpdate || _searchTextBox == null)
                return;

            _typedText = _searchTextBox.Text;

            SearchSuggestions.Clear();

            if (string.IsNullOrWhiteSpace(_typedText))
            {
                SuggestPopup.IsOpen = false;
                return;
            }

            foreach (var item in _searchIndex
                .Where(x => x.Alias.Contains(_typedText, StringComparison.OrdinalIgnoreCase))
                .Take(50))
            {
                SearchSuggestions.Add(item);
            }

            if (SearchSuggestions.Count > 0)
            {
                SuggestPopup.IsOpen = true;

                // cegah auto-select item pertama
                _suppressSearchUpdate = true;
                if (!_isKeyboardNavigating)
                {
                    _suppressSearchUpdate = true;
                    SuggestList.SelectedIndex = -1;
                    _suppressSearchUpdate = false;
                }
                _suppressSearchUpdate = false;

                // ✅ penting: setelah UI update, paksa caret di akhir & unselect text
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    FixCaretNoSelection();
                }), System.Windows.Threading.DispatcherPriority.Input);
            }
        }

        private void FixCaretNoSelection()
        {
            if (_searchTextBox == null) return;

            var t = _searchTextBox.Text ?? "";
            _searchTextBox.SelectionStart = t.Length;
            _searchTextBox.SelectionLength = 0;
            _searchTextBox.CaretIndex = t.Length;
        }

        private void SearchBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressSearchUpdate) return;

            // commit hanya kalau mouse (kalau kamu masih pakai fitur klik langsung add)
            if (_mouseCommitRequested && SuggestList.SelectedItem is SearchItem item)
            {
                _mouseCommitRequested = false;
                CommitSearchItem(item);
            }
        }

        private void RestoreTypedText()
        {
            if (_searchTextBox == null) return;

            _suppressSearchUpdate = true;
            _searchTextBox.Text = _typedText;
            _suppressSearchUpdate = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                FixCaretNoSelection();
            }), System.Windows.Threading.DispatcherPriority.Input);
        }




        private void CommitSearchItem(SearchItem item)
        {
            AddJobFromMap(item.Map);

            ClearSearchBox();
        }

        private void AddJobFromMap(DataMapRow map)
        {
            string pdf = (map.FilePath ?? "").Trim();

            var pageFrom = Math.Max(1, map.PageFrom);
            var pageTo = map.PageTo <= 0 ? 0 : Math.Max(1, map.PageTo);

            var row = new JobRow
            {
                Index = Rows.Count + 1,

                // info dari mapper (biar kebaca di grid)
                OrderNo = "(manual)",
                ProductName = map.SearchAlias ?? "",
                VariationName = (map.Variasi ?? "").Trim(),
                VariationCode = (map.NoRef ?? "").Trim(),
                OrderCreatedAt = DateTime.Now,

                File = pdf,
                Printer = !string.IsNullOrWhiteSpace(map.Printer) ? map.Printer.Trim() : (Printers.FirstOrDefault() ?? ""),
                PageFrom = pageFrom,
                PageTo = pageTo,
                Copies = Math.Max(1, map.Copies),
                Duplex = ParseDuplex(map.Duplex),
                Paper = ParsePaper(map.Paper),

                // legacy field
                Pages = $"{pageFrom}{(pageTo <= 0 ? "-" : (pageTo == pageFrom ? "" : "-" + pageTo))}",

                Status = "Ready",
                Percent = 0,
                TotalPages = 0
            };

            Rows.Add(row);

            QueueGrid.SelectedItem = row;
            QueueGrid.ScrollIntoView(row);
        }

        // Model item untuk autosuggest
        public class SearchItem
        {
            public string Display { get; set; } = "";
            public string Alias { get; set; } = "";
            public DataMapRow Map { get; set; } = default!;
        }

        private static readonly object _logLock = new object();
        private static string LogPath => Path.Combine(AppContext.BaseDirectory, "print_log.txt");

        private void LogPrint(string text)
        {
            try
            {
                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {text}{Environment.NewLine}";
                lock (_logLock)
                {
                    var existing = File.Exists(LogPath) ? File.ReadAllText(LogPath) : "";
                    File.WriteAllText(LogPath, line + existing);
                }
            }
            catch { }
        }

        private void OpenPrintLog_Click(object sender, RoutedEventArgs e)
        {
            var w = new PrintLogWindow(LogPath) { Owner = this };
            w.Show();
        }

        private void MarkScrolling()
        {
            _suppressPreviewWhileScrolling = true;

            _scrollStopTimer ??= new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(250)
            };

            _scrollStopTimer.Stop();
            _scrollStopTimer.Tick -= ScrollStopTimer_Tick;
            _scrollStopTimer.Tick += ScrollStopTimer_Tick;
            _scrollStopTimer.Start();
        }

        private void ScrollStopTimer_Tick(object? sender, EventArgs e)
        {
            _scrollStopTimer!.Stop();
            _suppressPreviewWhileScrolling = false;

            // Jangan ScrollIntoView / paksa SelectedItem — itu menarik scroll balik ke baris terpilih
            // setelah user sengaja menggulir (mis. habis Print di baris lain).
            if (_queueGridPreviewDeferredWhileScrolling && QueueGrid.SelectedItem is JobRow row && Rows.Contains(row))
            {
                _queueGridPreviewDeferredWhileScrolling = false;
                StartPreviewLatest(row);
            }
        }

        private void QueueGrid_PreviewMouseWheel(object sender, MouseWheelEventArgs e) => MarkScrolling();

        private void QueueGrid_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (e.VerticalChange != 0 || e.HorizontalChange != 0)
                MarkScrolling();
        }

        private void QueueGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // setelah klik, baris ter-select → SelectionChanged memicu preview
            _suppressPreviewWhileScrolling = false;
        }

        // =====================
        // ✅ ADD: DataMap (SKU+Variasi -> setting)
        // =====================
        private readonly Dictionary<string, DataMapRow> _dataMap = new(StringComparer.OrdinalIgnoreCase);

        private void LoadMap_Click(object sender, RoutedEventArgs e)
        {
            LoadDataMapXlsx_Click(sender, e); // delegasi ke method yang sudah ada
        }

        // === WRAPPER supaya cocok dengan XAML lama ===
        private void ImportOrders_Click(object sender, RoutedEventArgs e)
        {
            ImportOrdersXlsx_Click(sender, e);
        }

        
        

// =====================
// ✅ Shopee Open Platform V2 (Connect + SQLite state + save order process)
// =====================
private const long PartnerId = 2014528;
private const string PartnerKey = "shpk775655535a714c7243495256765a66737058417a59506c4a48525a765468";
private const string ApiHost = "https://partner.shopeemobile.com";
private const string RedirectUrl = "http://localhost:5123/callback/";

private static readonly string DbFilePath =
    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "PaperbellAppDotNet", "app.db");

        private  HttpClient _shopeeHttp;
private long _shopId;
private string? _accessToken;
private string? _refreshToken;
private DateTimeOffset _accessTokenExpiredAt = DateTimeOffset.MinValue;

private void EnsureDbDirectory()
{
    var dir = Path.GetDirectoryName(DbFilePath);
    if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
        Directory.CreateDirectory(dir);
}

        private int GetOrderProcessCount()
        {
            using var con = OpenDb();
            con.Open();

            using var cmd = con.CreateCommand();

            // Hitung jumlah row sesuai tab aktif
            switch (_currentTabFilter)
            {
                case ShopeeTabFilter.All:
                    cmd.CommandText = "SELECT COUNT(*) FROM order_process;";
                    break;

                case ShopeeTabFilter.NotPrinted:
                    cmd.CommandText = "SELECT COUNT(*) FROM order_process WHERE printed = 0;";
                    break;

                case ShopeeTabFilter.Printed:
                    cmd.CommandText = "SELECT COUNT(*) FROM order_process WHERE printed = 1;";
                    break;
            }

            return Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
        }




        private void LoadShopeePageFromDb(int pageIndex)
        {
            // hitung total item sesuai tab aktif (All / NotPrinted / Printed)
            _shopeeTotalItems = GetOrderProcessCount();

            // clamp page index berdasarkan total di tab ini
            var maxPageIndex = Math.Max(0, (int)Math.Ceiling(_shopeeTotalItems / (double)ShopeePageSize) - 1);
            _shopeePageIndex = Math.Max(0, Math.Min(pageIndex, maxPageIndex));

            var skip = _shopeePageIndex * ShopeePageSize;

            // simpan page index ke state tab masing-masing
            switch (_currentTabFilter)
            {
                case ShopeeTabFilter.All:
                    _pageIndexAll = _shopeePageIndex;
                    break;
                case ShopeeTabFilter.NotPrinted:
                    _pageIndexNotPrinted = _shopeePageIndex;
                    break;
                case ShopeeTabFilter.Printed:
                    _pageIndexPrinted = _shopeePageIndex;
                    break;
            }

            Rows.Clear();

            using var con = OpenDb();
            con.Open();

            using var cmd = con.CreateCommand();

            // bikin WHERE sesuai tab aktif
            string whereClause = "";
            if (_currentTabFilter == ShopeeTabFilter.NotPrinted)
                whereClause = "WHERE printed = 0\n";
            else if (_currentTabFilter == ShopeeTabFilter.Printed)
                whereClause = "WHERE printed = 1\n";

            cmd.CommandText = $@"
SELECT id, order_sn, item_key, model_sku, item_sku, item_name, model_name, qty, status, create_time, printed
FROM order_process
{whereClause}ORDER BY create_time DESC, id DESC
LIMIT $take OFFSET $skip;
";
            cmd.Parameters.AddWithValue("$take", ShopeePageSize);
            cmd.Parameters.AddWithValue("$skip", skip);

            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                var id = rd.IsDBNull(0) ? 0 : rd.GetInt64(0);
                var orderSn = rd.IsDBNull(1) ? "" : rd.GetString(1);
                var itemKey = rd.IsDBNull(2) ? "" : rd.GetString(2); // model_sku+item_sku
                var modelSku = rd.IsDBNull(3) ? "" : rd.GetString(3);
                var itemSku = rd.IsDBNull(4) ? "" : rd.GetString(4);
                var itemName = rd.IsDBNull(5) ? "" : rd.GetString(5);
                var modelName = rd.IsDBNull(6) ? "" : rd.GetString(6);
                var qty = rd.IsDBNull(7) ? 1 : rd.GetInt32(7);
                var createT = rd.IsDBNull(9) ? 0 : rd.GetInt64(9);
                var ordPrinted = rd.GetOrdinal("printed");

                _dataMap.TryGetValue(itemKey, out var map);

                JobRow row;
                if (map != null)
                {
                    var pf = Math.Max(1, map.PageFrom);
                    var pt = map.PageTo <= 0 ? 0 : Math.Max(1, map.PageTo);
                    var baseIndex = (_shopeePageIndex * ShopeePageSize);
                    row = new JobRow
                    {
                        OrderProcessId = id,
                        Index = baseIndex + Rows.Count + 1,
                        OrderNo = orderSn,
                        ProductName = itemName,
                        VariationName = modelName,

                        // PENTING: pakai itemKey (model_sku+item_sku) biar konsisten sama lookup DataMap
                        VariationCode = itemKey,

                        OrderCreatedAt = UnixToLocalDateTime(createT),

                        File = (map.FilePath ?? "").Trim(),
                        Printer = !string.IsNullOrWhiteSpace(map.Printer) ? map.Printer.Trim() : (Printers.FirstOrDefault() ?? ""),
                        PageFrom = pf,
                        PageTo = pt,
                        Copies = Math.Max(1, qty) * Math.Max(1, map.Copies),
                        Duplex = ParseDuplex(map.Duplex),
                        Paper = ParsePaper(map.Paper),
                        Pages = $"{pf}{(pt <= 0 ? "-" : (pt == pf ? "" : "-" + pt))}",
                        Status = "Ready",
                        Percent = 0,
                        TotalPages = 0
                    };
                    row.IsPrinted = !rd.IsDBNull(ordPrinted) && rd.GetInt32(ordPrinted) == 1;

                }
                else
                {
                    row = new JobRow
                    {
                        OrderProcessId = id,
                        Index = Rows.Count + 1,
                        OrderNo = orderSn,
                        ProductName = itemName,
                        VariationName = modelName,
                        VariationCode = itemKey,
                        OrderCreatedAt = UnixToLocalDateTime(createT),
                        Copies = Math.Max(1, qty),
                        Status = "UNMAPPED (DB)",
                        Percent = 0,
                        TotalPages = 0
                    };
                }

                                row.IsPrinted = !rd.IsDBNull(ordPrinted) && rd.GetInt32(ordPrinted) == 1;
                row.CanTogglePrinted = row.OrderProcessId > 0;

                Rows.Add(row);
            }

            UpdateShopeePagingUi();
            RefreshViews();
        }

        private void FilterTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return; // biar nggak kepanggil waktu InitializeComponent

            // Pindahkan state page sekarang ke field tiap tab (safety, walau LoadShopee juga sudah update)
            switch (_currentTabFilter)
            {
                case ShopeeTabFilter.All:
                    _pageIndexAll = _shopeePageIndex;
                    break;
                case ShopeeTabFilter.NotPrinted:
                    _pageIndexNotPrinted = _shopeePageIndex;
                    break;
                case ShopeeTabFilter.Printed:
                    _pageIndexPrinted = _shopeePageIndex;
                    break;
            }

            // Tentukan filter baru & restore page index-nya
            switch (FilterTabs.SelectedIndex)
            {
                case 0: // All
                    _currentTabFilter = ShopeeTabFilter.All;
                    _shopeePageIndex = _pageIndexAll;
                    break;

                case 1: // Not Printed
                default:
                    _currentTabFilter = ShopeeTabFilter.NotPrinted;
                    _shopeePageIndex = _pageIndexNotPrinted;
                    break;

                case 2: // Printed
                    _currentTabFilter = ShopeeTabFilter.Printed;
                    _shopeePageIndex = _pageIndexPrinted;
                    break;
            }

            // reload data berdasarkan tab + page yang baru dipilih
            LoadShopeePageFromDb(_shopeePageIndex);
        }

        private void UpdateShopeePagingUi()
        {
            var totalPages = Math.Max(1, (int)Math.Ceiling(_shopeeTotalItems / (double)ShopeePageSize));

            // Ini butuh 3 control di XAML (lihat step 3)
            TxtShopeePage.Text = $"Page {_shopeePageIndex + 1} / {totalPages}  (Total: {_shopeeTotalItems})";

            BtnShopeePrev.IsEnabled = _shopeePageIndex > 0;
            BtnShopeeNext.IsEnabled = (_shopeePageIndex + 1) < totalPages;
        }

        private void BtnShopeePrev_Click(object sender, RoutedEventArgs e)
        {
            LoadShopeePageFromDb(_shopeePageIndex - 1);
        }

        private void BtnShopeeNext_Click(object sender, RoutedEventArgs e)
        {
            LoadShopeePageFromDb(_shopeePageIndex + 1);
        }

        private SqliteConnection OpenDb() => new SqliteConnection($"Data Source={DbFilePath}");

private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
{
    // 1) load state dari DB
    LoadAppStateFromDb();   // pakai yang kamu punya

    // 2) update UI dulu (biar tidak blank)
    UpdateShopeeUi();

    // 3) kalau kelihatan "connected", coba validasi cepat
    if (IsConnected())
    {
        var ok = await CheckShopeeSessionAsync();
        if (!ok)
        {
            DisconnectShopee("Sesi Shopee sudah tidak valid. Silakan Connect ulang.");
        }
        else
        {
            // tetap connected, pastikan UI sesuai
            Dispatcher.Invoke(UpdateShopeeUi);
        }
    }

    try
    {
        if (CmbResiPrinter != null && CmbResiPrinter.SelectedItem == null && Printers.Count > 0)
            SelectResiPrinterDefaultOrRestore(null);
    }
    catch { }

    UpdatePreviewChromeForActiveTab();
    EnsureResiGridInSelectedTab();
}

private void InitDatabase()
{
    using var con = OpenDb();
    con.Open();

    using var cmd = con.CreateCommand();
    cmd.CommandText =
        """
        CREATE TABLE IF NOT EXISTS app_state (
            key TEXT PRIMARY KEY,
            value TEXT
        );

        CREATE TABLE IF NOT EXISTS orders (
            order_sn TEXT PRIMARY KEY,
            status TEXT,
            create_time INTEGER,
            update_time INTEGER,
            raw_json TEXT
        );

        CREATE TABLE IF NOT EXISTS order_process (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_sn TEXT,
            item_key TEXT,
            model_sku TEXT,
            item_sku TEXT,
            item_name TEXT,
            model_name TEXT,
            qty INTEGER,
            status TEXT,
            create_time INTEGER,
            saved_at INTEGER,
            printed INTEGER DEFAULT 0,
            printed_at INTEGER,
            UNIQUE(order_sn, item_key)
        );

        CREATE TABLE IF NOT EXISTS order_resi (
            order_sn TEXT PRIMARY KEY,
            pdf_path TEXT,
            fetched_at INTEGER,
            resi_printed INTEGER DEFAULT 0,
            resi_printed_at INTEGER
        );
        """;
    cmd.ExecuteNonQuery();
}

        private void DbSetPrinted(string orderSn, string itemKey, bool printed)
        {
            using var con = OpenDb();
            con.Open();

            using var cmd = con.CreateCommand();
            cmd.CommandText = @"
UPDATE order_process
SET printed = $p,
    printed_at = CASE WHEN $p=1 THEN $t ELSE NULL END
WHERE order_sn = $sn AND item_key = $k;
";
            cmd.Parameters.AddWithValue("$p", printed ? 1 : 0);
            cmd.Parameters.AddWithValue("$t", DateTimeOffset.UtcNow.ToUnixTimeSeconds());
            cmd.Parameters.AddWithValue("$sn", orderSn ?? "");
            cmd.Parameters.AddWithValue("$k", itemKey ?? "");
            cmd.ExecuteNonQuery();
        }

        private int DbSetPrintedById(long id, bool printed)
        {
            if (id <= 0) return 0;
            using var con = OpenDb();
            con.Open();

            using var cmd = con.CreateCommand();
            cmd.CommandText = @"
UPDATE order_process
SET printed = $p,
    printed_at = CASE WHEN $p=1 THEN $t ELSE NULL END
WHERE id = $id;
";
            cmd.Parameters.AddWithValue("$p", printed ? 1 : 0);
            cmd.Parameters.AddWithValue("$t", DateTimeOffset.UtcNow.ToUnixTimeSeconds());
            cmd.Parameters.AddWithValue("$id", id);
            return cmd.ExecuteNonQuery();
        }

        private string? GetState(string key)
{
    using var con = OpenDb();
    con.Open();
    using var cmd = con.CreateCommand();
    cmd.CommandText = "SELECT value FROM app_state WHERE key=$k LIMIT 1;";
    cmd.Parameters.AddWithValue("$k", key);
    return cmd.ExecuteScalar() as string;
}

private long GetStateLong(string key, long defaultValue)
{
    var s = GetState(key);
    return long.TryParse(s, out var v) ? v : defaultValue;
}

private void SetState(string key, string value)
{
    using var con = OpenDb();
    con.Open();
    using var cmd = con.CreateCommand();
    cmd.CommandText =
        """
        INSERT INTO app_state(key,value) VALUES($k,$v)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value;
        """;
    cmd.Parameters.AddWithValue("$k", key);
    cmd.Parameters.AddWithValue("$v", value ?? "");
    cmd.ExecuteNonQuery();
}

private void SetStateLong(string key, long value) => SetState(key, value.ToString());

private static string ProtectToBase64(string plain)
{
    var bytes = Encoding.UTF8.GetBytes(plain);
    var enc = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
    return Convert.ToBase64String(enc);
}

private static string UnprotectFromBase64(string b64)
{
    var enc = Convert.FromBase64String(b64);
    var bytes = ProtectedData.Unprotect(enc, null, DataProtectionScope.CurrentUser);
    return Encoding.UTF8.GetString(bytes);
}

private void LoadAppStateFromDb()
{
    _shopId = GetStateLong("shop_id", 0);

    var atEnc = GetState("access_token_enc");
    var rtEnc = GetState("refresh_token_enc");
    var expUnix = GetStateLong("access_token_exp_unix", 0);

    _accessToken = string.IsNullOrWhiteSpace(atEnc) ? null : UnprotectFromBase64(atEnc);
    _refreshToken = string.IsNullOrWhiteSpace(rtEnc) ? null : UnprotectFromBase64(rtEnc);
    _accessTokenExpiredAt = expUnix > 0 ? DateTimeOffset.FromUnixTimeSeconds(expUnix) : DateTimeOffset.MinValue;
}

private void SaveAppStateToDb()
{
    if (_shopId > 0) SetStateLong("shop_id", _shopId);

    if (!string.IsNullOrWhiteSpace(_accessToken))
        SetState("access_token_enc", ProtectToBase64(_accessToken!));

    if (!string.IsNullOrWhiteSpace(_refreshToken))
        SetState("refresh_token_enc", ProtectToBase64(_refreshToken!));

    if (_accessTokenExpiredAt > DateTimeOffset.MinValue)
        SetStateLong("access_token_exp_unix", _accessTokenExpiredAt.ToUnixTimeSeconds());
}

private bool IsConnected() => _shopId > 0 && !string.IsNullOrWhiteSpace(_accessToken);

private static long UnixNow() => DateTimeOffset.UtcNow.ToUnixTimeSeconds();

private static string HmacSha256Hex(string baseString, string key)
{
    using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(key));
    var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(baseString));
    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
}

private string BuildAuthPartnerUrl(string redirectUrl)
{
    var path = "/api/v2/shop/auth_partner";
    var ts = UnixNow();
    var baseString = $"{PartnerId}{path}{ts}";
    var sign = HmacSha256Hex(baseString, PartnerKey);
    return $"{ApiHost}{path}?partner_id={PartnerId}&timestamp={ts}&sign={sign}&redirect={Uri.EscapeDataString(redirectUrl)}";
}

private static void OpenBrowser(string url)
{
    Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
}

private static async Task WriteHttpResponseAsync(HttpListenerResponse resp, string html)
{
    var bytes = Encoding.UTF8.GetBytes(html);
    resp.ContentType = "text/html; charset=utf-8";
    resp.ContentLength64 = bytes.Length;
    await resp.OutputStream.WriteAsync(bytes, 0, bytes.Length);
    resp.OutputStream.Close();
}

private async Task<JsonDocument> PostJsonAsync(string url, object body, CancellationToken ct)
{
    var json = JsonSerializer.Serialize(body);
    using var req = new HttpRequestMessage(HttpMethod.Post, url);
    req.Content = new StringContent(json, Encoding.UTF8, "application/json");
    using var resp = await _shopeeHttp.SendAsync(req, ct);
    var text = await resp.Content.ReadAsStringAsync(ct);
    if (!resp.IsSuccessStatusCode)
        throw new HttpRequestException($"HTTP {(int)resp.StatusCode} {resp.ReasonPhrase}: {text}");
    return JsonDocument.Parse(text);
}

private async Task ExchangeTokenAsync(string code, CancellationToken ct)
{
    var path = "/api/v2/auth/token/get";
    var ts = UnixNow();
    var baseString = $"{PartnerId}{path}{ts}";
    var sign = HmacSha256Hex(baseString, PartnerKey);
    var url = $"{ApiHost}{path}?partner_id={PartnerId}&timestamp={ts}&sign={sign}";

    var doc = await PostJsonAsync(url, new { code }, ct);

    _accessToken = doc.RootElement.TryGetProperty("access_token", out var at) ? at.GetString() : null;
    _refreshToken = doc.RootElement.TryGetProperty("refresh_token", out var rt) ? rt.GetString() : null;

    if (string.IsNullOrWhiteSpace(_accessToken))
        throw new Exception("access_token kosong — token exchange gagal");

    var expireIn = doc.RootElement.TryGetProperty("expire_in", out var ex) ? ex.GetInt64() : 3600;
    _accessTokenExpiredAt = DateTimeOffset.UtcNow.AddSeconds(expireIn - 60);

    if (doc.RootElement.TryGetProperty("shop_id_list", out var shopArr) &&
        shopArr.ValueKind == JsonValueKind.Array && shopArr.GetArrayLength() > 0)
    {
        _shopId = shopArr[0].GetInt64();
    }
}

private async void Connect_Click(object sender, RoutedEventArgs e)
{
    try
    {
        IsEnabled = false;

        using var listener = new HttpListener();
        listener.Prefixes.Add(RedirectUrl);
        listener.Start();

        var authUrl = BuildAuthPartnerUrl(RedirectUrl);
        OpenBrowser(authUrl);

        var ctx = await listener.GetContextAsync();
        var q = ctx.Request.QueryString;

        var code = q["code"];
        var shopIdStr = q["shop_id"] ?? q["main_account_id"];

        await WriteHttpResponseAsync(ctx.Response,
            "<html><body><h3>Authorized.</h3><p>Silakan tutup tab ini dan kembali ke aplikasi.</p></body></html>");

        if (string.IsNullOrWhiteSpace(code))
            throw new InvalidOperationException("Callback missing 'code'.");

        if (!string.IsNullOrWhiteSpace(shopIdStr) && long.TryParse(shopIdStr, out var sid))
            _shopId = sid;

        await ExchangeTokenAsync(code, CancellationToken.None);

        SaveAppStateToDb();
        UpdateShopeeUi();
    }
    catch (Exception ex)
    {
        MessageBox.Show("Connect error:\n" + ex, "Shopee", MessageBoxButton.OK, MessageBoxImage.Error);
    }
    finally
    {
        IsEnabled = true;
    }
}

private string MakeSignedApiSign(string apiPath, long ts, string accessToken, long shopId)
{
    var baseStr = $"{PartnerId}{apiPath}{ts}{accessToken}{shopId}";
    return HmacSha256Hex(baseStr, PartnerKey);
}

private async Task<JsonDocument> GetShopApiAsync(string path, IDictionary<string, string> extraQuery, CancellationToken ct)
{
    if (string.IsNullOrWhiteSpace(_accessToken) || _shopId <= 0)
        throw new InvalidOperationException("Not connected (missing access_token/shop_id).");

    var ts = UnixNow();
    var sign = MakeSignedApiSign(path, ts, _accessToken!, _shopId);

    var q = new Dictionary<string, string>
    {
        ["partner_id"] = PartnerId.ToString(CultureInfo.InvariantCulture),
        ["timestamp"] = ts.ToString(CultureInfo.InvariantCulture),
        ["sign"] = sign,
        ["access_token"] = _accessToken!,
        ["shop_id"] = _shopId.ToString(CultureInfo.InvariantCulture),
    };

    foreach (var kv in extraQuery) q[kv.Key] = kv.Value;

    var qs = string.Join("&", q.Select(kv => $"{Uri.EscapeDataString(kv.Key)}={Uri.EscapeDataString(kv.Value)}"));
    var url = $"{ApiHost}{path}?{qs}";

    using var resp = await _shopeeHttp.GetAsync(url, ct);
    var text = await resp.Content.ReadAsStringAsync(ct);

    if (!resp.IsSuccessStatusCode)
        throw new HttpRequestException($"HTTP {(int)resp.StatusCode} {resp.ReasonPhrase}: {text}");

    return JsonDocument.Parse(text);
}

private static string? TryFormatShopeeResultListFailures(JsonElement root)
        {
            if (!root.TryGetProperty("response", out var resp))
                return null;
            if (!resp.TryGetProperty("result_list", out var rl) || rl.ValueKind != JsonValueKind.Array)
                return null;

            var lines = new List<string>();
            foreach (var item in rl.EnumerateArray())
            {
                var bits = new List<string>();
                if (item.TryGetProperty("order_sn", out var sn) && sn.ValueKind == JsonValueKind.String)
                {
                    var s = sn.GetString();
                    if (!string.IsNullOrWhiteSpace(s))
                        bits.Add($"order {s}");
                }

                if (item.TryGetProperty("fail_message", out var fm))
                {
                    if (fm.ValueKind == JsonValueKind.String)
                    {
                        var t = fm.GetString();
                        if (!string.IsNullOrWhiteSpace(t))
                            bits.Add($"fail_message={t}");
                    }
                }

                if (item.TryGetProperty("fail_error", out var fe))
                {
                    if (fe.ValueKind == JsonValueKind.String)
                    {
                        var t = fe.GetString();
                        if (!string.IsNullOrWhiteSpace(t))
                            bits.Add($"fail_error={t}");
                    }
                    else if (fe.ValueKind == JsonValueKind.Number)
                        bits.Add($"fail_error={fe.GetRawText()}");
                }

                if (item.TryGetProperty("message", out var msgEl) && msgEl.ValueKind == JsonValueKind.String)
                {
                    var t = msgEl.GetString();
                    if (!string.IsNullOrWhiteSpace(t))
                        bits.Add($"message={t}");
                }

                if (bits.Count > 0)
                    lines.Add(string.Join(" | ", bits));
            }

            return lines.Count > 0 ? string.Join(Environment.NewLine, lines) : null;
        }

        private static void EnsureShopeeOkOrThrow(JsonElement root)
        {
            if (root.TryGetProperty("error", out var err) && err.ValueKind == JsonValueKind.String)
            {
                var e = err.GetString();
                if (!string.IsNullOrWhiteSpace(e) && e != "0")
                {
                    var msg = root.TryGetProperty("message", out var m) ? m.GetString() : "";
                    var listDetail = TryFormatShopeeResultListFailures(root);
                    if (!string.IsNullOrWhiteSpace(listDetail))
                        throw new Exception($"Shopee error: {e} {msg}\nDetail:\n{listDetail}");
                    throw new Exception($"Shopee error: {e} {msg}");
                }
            }
        }

        /// <summary>
        /// Download API kadang menolak jika create_shipping_document belum pernah dipanggil untuk order ini,
        /// meskipun get_shipping_document_result sempat menunjukkan READY (mis. beda channel / Seller Centre).
        /// </summary>
        private static bool ShopeeExceptionIsShippingDocPrintFirst(Exception ex)
        {
            var m = ex.Message ?? "";
            return m.IndexOf("shipping_document_should_print_first", StringComparison.OrdinalIgnoreCase) >= 0;
        }

private async Task<List<string>> GetOrderSnsByStatusAsync(
    string status,
    long timeFrom,
    long timeTo,
    IProgress<SyncProgress> progress,
    CancellationToken ct)
{
    var path = "/api/v2/order/get_order_list";
    var query = new Dictionary<string, string>
    {
        ["time_range_field"] = "update_time",
        ["time_from"] = timeFrom.ToString(CultureInfo.InvariantCulture),
        ["time_to"] = timeTo.ToString(CultureInfo.InvariantCulture),
        ["page_size"] = "50",
        ["order_status"] = status
    };

    var all = new List<string>();
    string? cursor = null;

    do
    {
        if (!string.IsNullOrWhiteSpace(cursor))
            query["cursor"] = cursor;

        var doc = await GetShopApiWithLogAsync(path, query, progress, ct);
        EnsureShopeeOkOrThrow(doc.RootElement);

        cursor = null;

        if (doc.RootElement.TryGetProperty("response", out var resp))
        {
            if (resp.TryGetProperty("order_list", out var list) && list.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in list.EnumerateArray())
                {
                    if (item.TryGetProperty("order_sn", out var sn))
                    {
                        var s = sn.GetString();
                        if (!string.IsNullOrWhiteSpace(s)) all.Add(s);
                    }
                }
            }

            var more = resp.TryGetProperty("more", out var mo) && mo.ValueKind == JsonValueKind.True;
            var nextCursor = resp.TryGetProperty("next_cursor", out var nc) ? nc.GetString() : null;
            cursor = (more && !string.IsNullOrWhiteSpace(nextCursor)) ? nextCursor : null;
        }

        await Task.Delay(200, ct);
    }
    while (!string.IsNullOrWhiteSpace(cursor));

    return all;
}

private async Task<List<JsonElement>> GetOrderDetailBatchRawAsync(
    List<string> orderSnList,
    IProgress<SyncProgress> progress,
    CancellationToken ct)
        {
    var path = "/api/v2/order/get_order_detail";

    var query = new Dictionary<string, string>
    {
        ["order_sn_list"] = string.Join(",", orderSnList),
        ["response_optional_fields"] = "order_status,create_time,update_time,item_list"
    };

    var doc = await GetShopApiWithLogAsync(path, query, progress, ct); ;
    EnsureShopeeOkOrThrow(doc.RootElement);

    if (doc.RootElement.TryGetProperty("response", out var resp) &&
        resp.TryGetProperty("order_list", out var ol) &&
        ol.ValueKind == JsonValueKind.Array)
    {
        return ol.EnumerateArray().ToList();
    }

    return new List<JsonElement>();
}

private void UpsertOrderRaw(string orderSn, string status, long createTime, long updateTime, string rawJson)
{
    using var con = OpenDb();
    con.Open();

    using var cmd = con.CreateCommand();
    cmd.CommandText = @"
INSERT INTO orders(order_sn,status,create_time,update_time,raw_json)
VALUES($sn,$st,$ct,$ut,$rj)
ON CONFLICT(order_sn) DO UPDATE SET
  status=excluded.status,
  create_time=excluded.create_time,
  update_time=excluded.update_time,
  raw_json=excluded.raw_json;";
    cmd.Parameters.AddWithValue("$sn", orderSn);
    cmd.Parameters.AddWithValue("$st", status ?? "");
    cmd.Parameters.AddWithValue("$ct", createTime);
    cmd.Parameters.AddWithValue("$ut", updateTime);
    cmd.Parameters.AddWithValue("$rj", rawJson ?? "");
    cmd.ExecuteNonQuery();
}

private void InsertOrderProcess(string orderSn, string itemKey, string modelSku, string itemSku,
    string itemName, string modelName, int qty, string status, long createTime)
{
    using var con = OpenDb();
    con.Open();

    using var cmd = con.CreateCommand();
    cmd.CommandText = @"
INSERT INTO order_process(order_sn,item_key,model_sku,item_sku,item_name,model_name,qty,status,create_time,saved_at)
VALUES($sn,$k,$ms,$is,$in,$mn,$q,$st,$ct,$sa)
ON CONFLICT(order_sn, item_key) DO UPDATE SET
    status=excluded.status,
    qty=excluded.qty,
    saved_at=excluded.saved_at,
    item_name=excluded.item_name,
    model_name=excluded.model_name;
";
    cmd.Parameters.AddWithValue("$sn", orderSn);
    cmd.Parameters.AddWithValue("$k", itemKey);
    cmd.Parameters.AddWithValue("$ms", modelSku ?? "");
    cmd.Parameters.AddWithValue("$is", itemSku ?? "");
    cmd.Parameters.AddWithValue("$in", itemName ?? "");
    cmd.Parameters.AddWithValue("$mn", modelName ?? "");
    cmd.Parameters.AddWithValue("$q", qty);
    cmd.Parameters.AddWithValue("$st", status ?? "");
    cmd.Parameters.AddWithValue("$ct", createTime);
    cmd.Parameters.AddWithValue("$sa", UnixNow());
    cmd.ExecuteNonQuery();
}

        private void LoadShopeeFromDbToUi(int lastDays = 8)
        {
            // ambil data terbaru N hari dari DB
            var fromUnix = DateTimeOffset.UtcNow.AddDays(-lastDays).ToUnixTimeSeconds();

            using var con = OpenDb();
            con.Open();

            using var cmd = con.CreateCommand();
            cmd.CommandText = @"
SELECT id, order_sn, item_key, model_sku, item_sku, item_name, model_name, qty, status, create_time, printed
FROM order_process
WHERE create_time >= $from
ORDER BY create_time DESC;
";
            cmd.Parameters.AddWithValue("$from", fromUnix);

            using var rd = cmd.ExecuteReader();

            // optional: biar gak dobel kalau Rows sudah ada isi
            var existing = new HashSet<string>(
                Rows.Select(r => $"{r.OrderNo}|{r.VariationCode}"),
                StringComparer.OrdinalIgnoreCase);

            int added = 0;

            while (rd.Read())
            {
                var id = rd.IsDBNull(0) ? 0 : rd.GetInt64(0);
                var orderSn = rd.GetString(1);
                var itemKey = rd.GetString(2);
                var modelSku = rd.GetString(3);
                var itemSku = rd.GetString(4);
                var itemName = rd.GetString(5);
                var modelName = rd.GetString(6);
                var qty = rd.GetInt32(7);
                var status = rd.GetString(8);
                var createT = rd.GetInt64(9);
                var printed = !rd.IsDBNull(10) && rd.GetInt32(10) == 1;

                // dedup UI
                var sig = $"{orderSn}|{itemKey}";
                if (existing.Contains(sig)) continue;

                _dataMap.TryGetValue(itemKey, out var map);

                JobRow row;

                if (map != null)
                {
                    row = new JobRow
                    {
                        OrderProcessId = id,
                        Index = Rows.Count + 1,
                        OrderNo = orderSn,
                        ProductName = itemName,
                        VariationName = modelName,
                        VariationCode = itemKey, // <- penting untuk dedup, bisa juga modelSku tapi itemKey lebih unik
                        OrderCreatedAt = UnixToLocalDateTime(createT),

                        File = (map.FilePath ?? "").Trim(),
                        Printer = !string.IsNullOrWhiteSpace(map.Printer) ? map.Printer.Trim() : (Printers.FirstOrDefault() ?? ""),
                        PageFrom = Math.Max(1, map.PageFrom),
                        PageTo = map.PageTo <= 0 ? 0 : Math.Max(1, map.PageTo),
                        Copies = Math.Max(1, qty) * Math.Max(1, map.Copies),
                        Duplex = ParseDuplex(map.Duplex),
                        Paper = ParsePaper(map.Paper),
                        Pages = $"{Math.Max(1, map.PageFrom)}{(map.PageTo <= 0 ? "-" : (map.PageTo == map.PageFrom ? "" : "-" + map.PageTo))}",
                        Status = "Ready",
                        Percent = 0,
                        TotalPages = 0
                    };
                }
                else
                {
                    // kalau DataMap belum ada / belum match
                    row = new JobRow
                    {
                        OrderProcessId = id,
                        Index = Rows.Count + 1,
                        OrderNo = orderSn,
                        ProductName = itemName,
                        VariationName = modelName,
                        VariationCode = itemKey,
                        OrderCreatedAt = UnixToLocalDateTime(createT),
                        Copies = Math.Max(1, qty),
                        Status = "UNMAPPED (DB)",
                        Percent = 0,
                        TotalPages = 0
                    };
                }

                                    row.IsPrinted = printed;
                    row.CanTogglePrinted = row.OrderProcessId > 0;
                    Rows.Add(row);
                existing.Add(sig);
                added++;
            }

            // rapikan index
            int idx = 1;
            foreach (var r in Rows.OrderByDescending(x => x.OrderCreatedAt ?? DateTime.MinValue).ToList())
                r.Index = idx++;

            // (optional) kamu bisa MessageBox kalau mau
            // MessageBox.Show($"Loaded {added} rows from DB");
        }

        private static DateTime? UnixToLocalDateTime(long unix)
    => unix > 0 ? DateTimeOffset.FromUnixTimeSeconds(unix).LocalDateTime : null;


// =====================
        // ✅ ADD: Shopee OpenAPI V2 (basic order fetch) 
        // =====================

        private const string DefaultShopeeBaseUrl = "https://partner.shopeemobile.com";
        private string ShopeeConfigPath => Path.Combine(ConfigDir, "shopee.json");

        private async Task SyncShopeeToDbAsync(IProgress<SyncProgress> progress)
        {
            int page = 0;
            int processedItems = 0;

            string? cursor = null;
            bool more = true;

            progress.Report(new SyncProgress { Percent = 0, Label = "0%", Log = "Fetching order list..." });

            while (more)
            {
                page++;

                progress.Report(new SyncProgress
                {
                    Percent = Math.Min(95, page * 10), // estimasi naik per page
                    Label = $"Page {page}",
                    Log = $"Request order list page {page} (cursor={cursor ?? "-"})"
                });

                // TODO: panggil GetOrderList API kamu (yang sudah ada)
                // var resp = await GetOrderListAsync(cursor);
                // more = resp.more;
                // cursor = resp.next_cursor;
                // var orders = resp.order_list;

                // Contoh loop items:
                foreach (var order in /*orders*/ Array.Empty<object>())
                {
                    // TODO: fetch detail jika kamu lakukan
                    // progress.Report(new SyncProgress { Log = $"Order {orderSn}: fetching detail..." });

                    // TODO: upsert order + insert/upsert order_process
                    // processedItems++;

                    if (processedItems % 20 == 0)
                    {
                        progress.Report(new SyncProgress
                        {
                            Percent = Math.Min(99, 20 + (page * 10)), // atau per item kalau ada total
                            Label = $"{processedItems} items",
                            Log = $"Processed {processedItems} items..."
                        });
                    }
                }

                // kalau nggak ada items / resp.more false
                // more = resp.more;
            }

            progress.Report(new SyncProgress { Percent = 99, Label = "Finalizing...", Log = "Finalizing DB..." });
        }

        private async Task<(string accessToken, string refreshToken)> RefreshShopeeAccessTokenAsync(
    string refreshToken,
    long shopId,
    IProgress<SyncProgress> progress,
    CancellationToken ct)
        {
            var path = "/api/v2/auth/access_token/get";
            var ts = UnixNow();
            var sign = MakeSignedApiSignAuthOnly(path, ts); // base string partner_id + path + timestamp

            var url = $"{ApiHost}{path}?partner_id={PartnerId}&timestamp={ts}&sign={sign}";

            var payload = new
            {
                partner_id = PartnerId,
                shop_id = shopId,
                refresh_token = refreshToken
            };

            progress.Report(new SyncProgress { Log = $"[API][REQ] POST {path} (refresh)" });

            using var resp = await _shopeeHttp.PostAsync(
                url,
                new StringContent(System.Text.Json.JsonSerializer.Serialize(payload), System.Text.Encoding.UTF8, "application/json"),
                ct);

            var body = await resp.Content.ReadAsStringAsync(ct);
            var shortBody = body.Length > 400 ? body[..400] + "..." : body;
            progress.Report(new SyncProgress { Log = $"[API][RESP] {(int)resp.StatusCode} {resp.ReasonPhrase} | {shortBody}" });

            if (!resp.IsSuccessStatusCode)
                throw new HttpRequestException(body);

            using var doc = JsonDocument.Parse(body);
            var newAccess = doc.RootElement.GetProperty("access_token").GetString() ?? "";
            var newRefresh = doc.RootElement.GetProperty("refresh_token").GetString() ?? "";

            if (string.IsNullOrWhiteSpace(newAccess) || string.IsNullOrWhiteSpace(newRefresh))
                throw new InvalidOperationException("Refresh token response missing access_token/refresh_token");

            return (newAccess, newRefresh);
        }



        private ShopeeConfig LoadShopeeConfig()
        {
            try
            {
                if (!File.Exists(ShopeeConfigPath)) return new ShopeeConfig { BaseUrl = DefaultShopeeBaseUrl };
                var json = File.ReadAllText(ShopeeConfigPath);
                var cfg = JsonSerializer.Deserialize<ShopeeConfig>(json) ?? new ShopeeConfig();
                if (string.IsNullOrWhiteSpace(cfg.BaseUrl)) cfg.BaseUrl = DefaultShopeeBaseUrl;
                return cfg;
            }
            catch
            {
                return new ShopeeConfig { BaseUrl = DefaultShopeeBaseUrl };
            }
        }

        private void SaveShopeeConfig(ShopeeConfig cfg)
        {
            Directory.CreateDirectory(ConfigDir);
            var json = JsonSerializer.Serialize(cfg, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ShopeeConfigPath, json);
        }


        private bool JobRowExists(string orderSn, string itemKey)
        {
            return Rows.Any(r =>
                r.OrderNo == orderSn &&
                (r.VariationCode ?? "") == itemKey
            );
        }

        private async void ShopeeSync_Click(object sender, RoutedEventArgs e)
        {
            if (!IsConnected())
            {
                MessageBox.Show("Belum connect. Klik 'Connect Shopee' dulu.", "Shopee",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var win = new SyncLogWindow { Owner = this };
            win.Show();
            await this.Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Background);

            IProgress<SyncProgress> progress =
                new System.Progress<SyncProgress>(p =>
                {
                    if (!string.IsNullOrWhiteSpace(p.Log))
                        win.AppendLog(p.Log);
                    win.SetProgress(p.Percent, p.Label);
                });

            try
            {
                IsEnabled = false;

                var cts = new CancellationTokenSource();
                win.CancelRequested += () => cts.Cancel();

                await Task.Run(async () =>
                {
                    progress.Report(new SyncProgress { Percent = 1, Label = "Starting...", Log = "Starting sync..." });

                    var timeTo = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
                    var timeFrom = DateTimeOffset.UtcNow.AddDays(-1).ToUnixTimeSeconds();
                    var statuses = new[] { "PROCESSED", "READY_TO_SHIP" };

                    var allSns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var st in statuses)
                    {
                        progress.Report(new SyncProgress { Log = $"Fetching order list status={st}..." });
                        var sns = await GetOrderSnsByStatusAsync(
    st, timeFrom, timeTo, progress, cts.Token);
                        foreach (var sn in sns) if (!string.IsNullOrWhiteSpace(sn)) allSns.Add(sn);
                    }

                    var orderSns = allSns.ToList();
                    const int batchSize = 50;

                    for (int i = 0; i < orderSns.Count; i += batchSize)
                    {
                        var batch = orderSns.Skip(i).Take(batchSize).ToList();
                        progress.Report(new SyncProgress
                        {
                            Percent = 5 + (90.0 * i / Math.Max(1, orderSns.Count)),
                            Label = $"{i}/{orderSns.Count}",
                            Log = $"Fetching details {i + 1}-{Math.Min(i + batchSize, orderSns.Count)}..."
                        });

                        var orders = await GetOrderDetailBatchRawAsync(
    batch, progress, cts.Token);

                        foreach (var o in orders)
                        {
                            var orderSn = o.TryGetProperty("order_sn", out var sn) ? sn.GetString() : "";
                            if (string.IsNullOrWhiteSpace(orderSn)) continue;

                            var status = o.TryGetProperty("order_status", out var st) ? st.GetString() ?? "" : "";
                            var createTime = o.TryGetProperty("create_time", out var ct) ? ct.GetInt64() : 0;
                            var updateTime = o.TryGetProperty("update_time", out var ut) ? ut.GetInt64() : 0;

                            UpsertOrderRaw(orderSn, status, createTime, updateTime, o.GetRawText());

                            if (o.TryGetProperty("item_list", out var il) && il.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var it in il.EnumerateArray())
                                {
                                    var itemSku = it.TryGetProperty("item_sku", out var isku) ? isku.GetString() ?? "" : "";
                                    var modelSku = it.TryGetProperty("model_sku", out var msku) ? msku.GetString() ?? "" : "";
                                    var itemName = it.TryGetProperty("item_name", out var iname) ? iname.GetString() ?? "" : "";
                                    var modelName = it.TryGetProperty("model_name", out var mname) ? mname.GetString() ?? "" : "";
                                    var qty = it.TryGetProperty("model_quantity_purchased", out var q) ? q.GetInt32() : 1;

                                    var itemKey = KeyModelItem(modelSku, itemSku);
                                    InsertOrderProcess(orderSn, itemKey, modelSku, itemSku, itemName, modelName, qty, status, createTime);
                                }
                            }
                        }

                        await Task.Delay(50, cts.Token);
                    }

                    progress.Report(new SyncProgress { Percent = 100, Label = "100%", Log = "Done." });
                }, cts.Token);

                // UI refresh sekali saja
                LoadShopeePageFromDb(0);

                win.SetDone("Sync completed ✅");
            }
            catch (OperationCanceledException)
            {
                win.AppendLog("Cancelled by user.");
                win.SetDone("Sync cancelled ⚠️");
            }
            catch (Exception ex)
            {
                win.AppendLog("ERROR: " + ex);
                win.SetDone("Sync failed ❌");
            }
            finally
            {
                IsEnabled = true;
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            // Jangan ganggu kalau user lagi ngetik di TextBox / ComboBox
            if (Keyboard.FocusedElement is TextBox ||
                Keyboard.FocusedElement is ComboBox)
                return;

            if (e.Key == Key.Left && _shopeePageIndex == 0)
                return;

            if (e.Key == Key.Right)
            {
                // Next page
                LoadShopeePageFromDb(_shopeePageIndex + 1);
                e.Handled = true;
            }
            else if (e.Key == Key.Left)
            {
                // Prev page
                LoadShopeePageFromDb(_shopeePageIndex - 1);
                e.Handled = true;
            }
        }

        private string MakeSignedApiSignAuthOnly(string apiPath, long ts)
        {
            var baseStr = $"{PartnerId}{apiPath}{ts}";
            return HmacSha256Hex(baseStr, PartnerKey);
        }

        /// <summary>
        /// Untuk log diagnostik: sembunyikan token/tanda tangan di query string URL Shopee.
        /// </summary>
        private static string RedactSensitiveShopeeUrlForLog(string url)
        {
            if (string.IsNullOrEmpty(url))
                return url;
            try
            {
                return Regex.Replace(
                    url,
                    @"([?&])(access_token|sign|refresh_token)=([^&]*)",
                    m => $"{m.Groups[1].Value}{m.Groups[2].Value}=***",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }
            catch
            {
                return url;
            }
        }

        private async Task<JsonDocument> GetShopApiWithLogAsync(
    string path,
    IDictionary<string, string> extraQuery,
    IProgress<SyncProgress> progress,
    CancellationToken ct,
    bool logFullRequestResponse = false)
        {
            if (string.IsNullOrWhiteSpace(_accessToken) || _shopId <= 0)
                throw new InvalidOperationException("Not connected");

            var ts = UnixNow();
            var sign = MakeSignedApiSign(path, ts, _accessToken!, _shopId);

            var q = new Dictionary<string, string>
            {
                ["partner_id"] = PartnerId.ToString(),
                ["timestamp"] = ts.ToString(),
                ["sign"] = sign,
                ["access_token"] = _accessToken!,
                ["shop_id"] = _shopId.ToString(),
            };

            foreach (var kv in extraQuery)
                q[kv.Key] = kv.Value;

            var qs = string.Join("&", q.Select(kv =>
                $"{Uri.EscapeDataString(kv.Key)}={Uri.EscapeDataString(kv.Value)}"));

            var url = $"{ApiHost}{path}?{qs}";

            if (logFullRequestResponse)
            {
                progress.Report(new SyncProgress
                {
                    Log = $"[API][REQ] GET\n{RedactSensitiveShopeeUrlForLog(url)}"
                });
            }
            else
            {
                progress.Report(new SyncProgress
                {
                    Log = $"[API][REQ] GET {path} ({extraQuery.Count} params)"
                });
            }

            using var resp = await _shopeeHttp.GetAsync(url, ct);
            var body = await resp.Content.ReadAsStringAsync(ct);

            if (logFullRequestResponse)
            {
                progress.Report(new SyncProgress
                {
                    Log = $"[API][RESP] {(int)resp.StatusCode} {resp.ReasonPhrase}\n{body}"
                });
            }
            else
            {
                var shortBody = body.Length > 400 ? body[..400] + "..." : body;
                progress.Report(new SyncProgress
                {
                    Log = $"[API][RESP] {(int)resp.StatusCode} {resp.ReasonPhrase} | {shortBody}"
                });
            }

            // ===============================
            // 🔴 AUTO REFRESH TOKEN DI SINI
            // ===============================
            if (resp.StatusCode == HttpStatusCode.Forbidden &&
                (body.Contains("invalid_access_token", StringComparison.OrdinalIgnoreCase) ||
                 body.Contains("invalid_acceess_token", StringComparison.OrdinalIgnoreCase)))
            {
                progress.Report(new SyncProgress
                {
                    Log = "[AUTH] access_token invalid → trying refresh_token..."
                });

                // pastikan ada refresh_token
                if (string.IsNullOrWhiteSpace(_refreshToken))
                    throw new InvalidOperationException("No refresh_token stored. Reconnect required.");

                try
                {
                    // refresh token
                    var (newAccess, newRefresh) =
                        await RefreshShopeeAccessTokenAsync(_refreshToken, _shopId, progress, ct);

                    // update memory
                    _accessToken = newAccess;
                    _refreshToken = newRefresh;

                    // simpan ke DB
                    SaveAppStateToDb();

                    progress.Report(new SyncProgress
                    {
                        Log = "[AUTH] token refreshed successfully, retrying request..."
                    });

                    // 🔁 retry API ONCE
                    return await GetShopApiWithLogAsync(path, extraQuery, progress, ct, logFullRequestResponse);
                }
                catch (HttpRequestException ex)
                {
                    // Idealnya RefreshShopeeAccessTokenAsync melempar HttpRequestException(bodyJson)
                    // sehingga ex.Message mengandung JSON error.
                    var msg = ex.Message ?? "";

                    if (msg.Contains("refresh_token_expired", StringComparison.OrdinalIgnoreCase))
                    {
                        progress.Report(new SyncProgress
                        {
                            Log = "[AUTH] refresh_token expired → reconnect required (clearing saved tokens)"
                        });

                        // clear auth state supaya nggak loop gagal terus
                        _accessToken = null;
                        _refreshToken = null;

                        // kalau kamu punya field expired-at, reset juga
                        // (hapus baris ini kalau field-nya tidak ada)
                        _accessTokenExpiredAt = DateTimeOffset.MinValue;

                        SaveAppStateToDb();

                        throw new InvalidOperationException(
                            "Shopee authorization expired. Silakan Connect ulang untuk generate token baru.");
                    }

                    // error lain: lempar lagi biar root cause tetap kelihatan
                    throw;
                }
            }

            // ===============================
            // NORMAL ERROR
            // ===============================
            if (!resp.IsSuccessStatusCode)
                throw new HttpRequestException(body);

            // sukses
            return JsonDocument.Parse(body);
        }

        private async Task<JsonDocument> PostShopApiWithLogAsync(
            string path,
            object jsonBody,
            IProgress<SyncProgress> progress,
            CancellationToken ct,
            bool logFullRequestResponse = false)
        {
            if (string.IsNullOrWhiteSpace(_accessToken) || _shopId <= 0)
                throw new InvalidOperationException("Not connected");

            var ts = UnixNow();
            var sign = MakeSignedApiSign(path, ts, _accessToken!, _shopId);

            var q = new Dictionary<string, string>
            {
                ["partner_id"] = PartnerId.ToString(),
                ["timestamp"] = ts.ToString(),
                ["sign"] = sign,
                ["access_token"] = _accessToken!,
                ["shop_id"] = _shopId.ToString(),
            };

            var qs = string.Join("&", q.Select(kv =>
                $"{Uri.EscapeDataString(kv.Key)}={Uri.EscapeDataString(kv.Value)}"));

            var url = $"{ApiHost}{path}?{qs}";
            var json = JsonSerializer.Serialize(jsonBody);
            using var content = new StringContent(json, Encoding.UTF8, "application/json");

            if (logFullRequestResponse)
            {
                progress.Report(new SyncProgress
                {
                    Log = $"[API][REQ] POST {path}\nURL: {RedactSensitiveShopeeUrlForLog(url)}\nBody: {json}"
                });
            }
            else
            {
                progress.Report(new SyncProgress { Log = $"[API][REQ] POST {path}" });
            }

            using var resp = await _shopeeHttp.PostAsync(url, content, ct);
            var body = await resp.Content.ReadAsStringAsync(ct);

            if (logFullRequestResponse)
            {
                progress.Report(new SyncProgress
                {
                    Log = $"[API][RESP] {(int)resp.StatusCode} {resp.ReasonPhrase}\n{body}"
                });
            }
            else
            {
                var shortBody = body.Length > 400 ? body[..400] + "..." : body;
                progress.Report(new SyncProgress
                {
                    Log = $"[API][RESP] {(int)resp.StatusCode} {resp.ReasonPhrase} | {shortBody}"
                });
            }

            if (resp.StatusCode == HttpStatusCode.Forbidden &&
                (body.Contains("invalid_access_token", StringComparison.OrdinalIgnoreCase) ||
                 body.Contains("invalid_acceess_token", StringComparison.OrdinalIgnoreCase)))
            {
                progress.Report(new SyncProgress { Log = "[AUTH] access_token invalid → trying refresh_token..." });

                if (string.IsNullOrWhiteSpace(_refreshToken))
                    throw new InvalidOperationException("No refresh_token stored. Reconnect required.");

                var (newAccess, newRefresh) =
                    await RefreshShopeeAccessTokenAsync(_refreshToken, _shopId, progress, ct);
                _accessToken = newAccess;
                _refreshToken = newRefresh;
                SaveAppStateToDb();

                progress.Report(new SyncProgress { Log = "[AUTH] token refreshed, retrying POST..." });
                return await PostShopApiWithLogAsync(path, jsonBody, progress, ct, logFullRequestResponse);
            }

            if (!resp.IsSuccessStatusCode)
                throw new HttpRequestException(body);

            return JsonDocument.Parse(body);
        }

        private static bool ShopeeResponseBytesLookLikePdf(ReadOnlySpan<byte> bytes) =>
            bytes.Length >= 5 &&
            bytes[0] == (byte)'%' &&
            bytes[1] == (byte)'P' &&
            bytes[2] == (byte)'D' &&
            bytes[3] == (byte)'F' &&
            bytes[4] == (byte)'-';

        /// <summary>
        /// POST download_shipping_document — Shopee kadang mengembalikan PDF mentah di body (bukan JSON).
        /// </summary>
        private async Task<(bool savedRawPdf, JsonDocument? json)> PostDownloadShippingDocumentWithLogAsync(
            object jsonBody,
            string pdfOutPath,
            IProgress<SyncProgress> progress,
            CancellationToken ct,
            bool logFullRequestResponse = false)
        {
            const string path = "/api/v2/logistics/download_shipping_document";

            if (string.IsNullOrWhiteSpace(_accessToken) || _shopId <= 0)
                throw new InvalidOperationException("Not connected");

            var ts = UnixNow();
            var sign = MakeSignedApiSign(path, ts, _accessToken!, _shopId);

            var q = new Dictionary<string, string>
            {
                ["partner_id"] = PartnerId.ToString(),
                ["timestamp"] = ts.ToString(),
                ["sign"] = sign,
                ["access_token"] = _accessToken!,
                ["shop_id"] = _shopId.ToString(),
            };

            var qs = string.Join("&", q.Select(kv =>
                $"{Uri.EscapeDataString(kv.Key)}={Uri.EscapeDataString(kv.Value)}"));

            var url = $"{ApiHost}{path}?{qs}";
            var json = JsonSerializer.Serialize(jsonBody);
            using var content = new StringContent(json, Encoding.UTF8, "application/json");

            if (logFullRequestResponse)
            {
                progress.Report(new SyncProgress
                {
                    Log = $"[API][REQ] POST {path}\nURL: {RedactSensitiveShopeeUrlForLog(url)}\nBody: {json}"
                });
            }
            else
            {
                progress.Report(new SyncProgress { Log = $"[API][REQ] POST {path}" });
            }

            using var resp = await _shopeeHttp.PostAsync(url, content, ct);
            var bytes = await resp.Content.ReadAsByteArrayAsync(ct);
            var bodyText = Encoding.UTF8.GetString(bytes);

            if (logFullRequestResponse)
            {
                var respLog = ShopeeResponseBytesLookLikePdf(bytes)
                    ? $"[API][RESP] {(int)resp.StatusCode} {resp.ReasonPhrase}\n[body: PDF binary, {bytes.Length} bytes — tidak di-dump ke log]"
                    : $"[API][RESP] {(int)resp.StatusCode} {resp.ReasonPhrase}\n{bodyText}";
                progress.Report(new SyncProgress { Log = respLog });
            }
            else
            {
                var logLine = ShopeeResponseBytesLookLikePdf(bytes)
                    ? $"[PDF binary, {bytes.Length} bytes]"
                    : (bodyText.Length > 400 ? bodyText[..400] + "..." : bodyText);
                progress.Report(new SyncProgress
                {
                    Log = $"[API][RESP] {(int)resp.StatusCode} {resp.ReasonPhrase} | {logLine}"
                });
            }

            if (resp.StatusCode == HttpStatusCode.Forbidden &&
                (bodyText.Contains("invalid_access_token", StringComparison.OrdinalIgnoreCase) ||
                 bodyText.Contains("invalid_acceess_token", StringComparison.OrdinalIgnoreCase)))
            {
                progress.Report(new SyncProgress { Log = "[AUTH] access_token invalid → trying refresh_token..." });

                if (string.IsNullOrWhiteSpace(_refreshToken))
                    throw new InvalidOperationException("No refresh_token stored. Reconnect required.");

                var (newAccess, newRefresh) =
                    await RefreshShopeeAccessTokenAsync(_refreshToken, _shopId, progress, ct);
                _accessToken = newAccess;
                _refreshToken = newRefresh;
                SaveAppStateToDb();

                progress.Report(new SyncProgress { Log = "[AUTH] token refreshed, retrying POST..." });
                return await PostDownloadShippingDocumentWithLogAsync(jsonBody, pdfOutPath, progress, ct,
                    logFullRequestResponse);
            }

            if (!resp.IsSuccessStatusCode)
                throw new HttpRequestException(bodyText);

            var media = resp.Content.Headers.ContentType?.MediaType ?? "";
            var treatAsPdf = media.Contains("pdf", StringComparison.OrdinalIgnoreCase) ||
                             ShopeeResponseBytesLookLikePdf(bytes);

            if (treatAsPdf)
            {
                await File.WriteAllBytesAsync(pdfOutPath, bytes, ct);
                LogPrint($"[label] download_shipping_document: PDF mentah disimpan ({bytes.Length} bytes).");
                return (true, null);
            }

            return (false, JsonDocument.Parse(bodyText));
        }

        private static string ResiStorageDir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PaperbellAppDotNet", "resi");

        /// <summary>
        /// Body create_shipping_document; sertakan tracking_number bila sudah diketahui (syarat beberapa channel).
        /// </summary>
        private static object ShopeeLogisticsOrderListCreate(
            string orderSn,
            string shippingDocumentType,
            string? packageNumber,
            string? trackingNumber)
        {
            var hasPkg = !string.IsNullOrWhiteSpace(packageNumber);
            var hasTn = !string.IsNullOrWhiteSpace(trackingNumber);
            if (!hasPkg && !hasTn)
            {
                return new
                {
                    order_list = new[]
                    {
                        new { order_sn = orderSn, shipping_document_type = shippingDocumentType }
                    }
                };
            }

            if (!hasPkg && hasTn)
            {
                return new
                {
                    order_list = new[]
                    {
                        new
                        {
                            order_sn = orderSn,
                            shipping_document_type = shippingDocumentType,
                            tracking_number = trackingNumber!
                        }
                    }
                };
            }

            if (hasPkg && !hasTn)
            {
                return new
                {
                    order_list = new[]
                    {
                        new
                        {
                            order_sn = orderSn,
                            package_number = packageNumber!,
                            shipping_document_type = shippingDocumentType
                        }
                    }
                };
            }

            return new
            {
                order_list = new[]
                {
                    new
                    {
                        order_sn = orderSn,
                        package_number = packageNumber!,
                        shipping_document_type = shippingDocumentType,
                        tracking_number = trackingNumber!
                    }
                }
            };
        }

        private static string? TryExtractTrackingNumberForOrder(JsonElement root, string orderSn, string? packageNumber)
        {
            if (!root.TryGetProperty("response", out var resp))
                return null;

            if (resp.TryGetProperty("tracking_number", out var tnTop) && tnTop.ValueKind == JsonValueKind.String)
            {
                var sTop = tnTop.GetString();
                if (!string.IsNullOrWhiteSpace(sTop))
                {
                    var orderMatches = !resp.TryGetProperty("order_sn", out var osTop) ||
                                       osTop.ValueKind != JsonValueKind.String ||
                                       string.Equals(osTop.GetString(), orderSn, StringComparison.Ordinal);
                    if (orderMatches)
                        return sTop;
                }
            }

            // get_mass_tracking_number mengembalikan response.success_list[] (bukan result_list).
            foreach (var listName in new[] { "success_list", "result_list", "package_list" })
            {
                if (!resp.TryGetProperty(listName, out var arr) || arr.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (var item in arr.EnumerateArray())
                {
                    if (item.TryGetProperty("order_sn", out var os) && os.ValueKind == JsonValueKind.String)
                    {
                        if (!string.Equals(os.GetString(), orderSn, StringComparison.Ordinal))
                            continue;
                    }

                    if (!string.IsNullOrWhiteSpace(packageNumber) &&
                        item.TryGetProperty("package_number", out var pn) && pn.ValueKind == JsonValueKind.String)
                    {
                        if (!string.Equals(pn.GetString(), packageNumber, StringComparison.Ordinal))
                            continue;
                    }

                    if (TryReadTrackingNumberFromLogisticsItem(item, out var trk))
                        return trk;
                }
            }

            if (resp.TryGetProperty("order_list", out var ol) && ol.ValueKind == JsonValueKind.Array)
            {
                foreach (var ord in ol.EnumerateArray())
                {
                    if (ord.TryGetProperty("order_sn", out var os2) && os2.ValueKind == JsonValueKind.String &&
                        !string.Equals(os2.GetString(), orderSn, StringComparison.Ordinal))
                        continue;

                    if (ord.TryGetProperty("package_list", out var pkgs) && pkgs.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var pkg in pkgs.EnumerateArray())
                        {
                            if (!string.IsNullOrWhiteSpace(packageNumber) &&
                                pkg.TryGetProperty("package_number", out var pn2) &&
                                pn2.ValueKind == JsonValueKind.String &&
                                !string.Equals(pn2.GetString(), packageNumber, StringComparison.Ordinal))
                                continue;

                            if (TryReadTrackingNumberFromLogisticsItem(pkg, out var trk))
                                return trk;
                        }
                    }

                    if (TryReadTrackingNumberFromLogisticsItem(ord, out var trk2))
                        return trk2;
                }
            }

            return null;
        }

        private static string? TryExtractPackageNumberForOrder(JsonElement root, string orderSn)
        {
            if (!root.TryGetProperty("response", out var resp))
                return null;

            foreach (var listName in new[] { "success_list", "result_list", "package_list" })
            {
                if (!resp.TryGetProperty(listName, out var arr) || arr.ValueKind != JsonValueKind.Array)
                    continue;
                foreach (var item in arr.EnumerateArray())
                {
                    if (item.TryGetProperty("order_sn", out var os) && os.ValueKind == JsonValueKind.String &&
                        !string.Equals(os.GetString(), orderSn, StringComparison.Ordinal))
                        continue;
                    if (item.TryGetProperty("package_number", out var pn) && pn.ValueKind == JsonValueKind.String)
                    {
                        var p = pn.GetString();
                        if (!string.IsNullOrWhiteSpace(p))
                            return p;
                    }
                }
            }

            if (resp.TryGetProperty("order_list", out var ol) && ol.ValueKind == JsonValueKind.Array)
            {
                foreach (var ord in ol.EnumerateArray())
                {
                    if (ord.TryGetProperty("order_sn", out var os2) && os2.ValueKind == JsonValueKind.String &&
                        !string.Equals(os2.GetString(), orderSn, StringComparison.Ordinal))
                        continue;
                    if (ord.TryGetProperty("package_list", out var pkgs) && pkgs.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var pkg in pkgs.EnumerateArray())
                        {
                            if (pkg.TryGetProperty("package_number", out var pn2) &&
                                pn2.ValueKind == JsonValueKind.String)
                            {
                                var p = pn2.GetString();
                                if (!string.IsNullOrWhiteSpace(p))
                                    return p;
                            }
                        }
                    }
                }
            }

            return null;
        }

        private static bool TryReadTrackingNumberFromLogisticsItem(JsonElement item, out string? tracking)
        {
            tracking = null;
            foreach (var key in new[] { "tracking_number", "first_mile_tracking_number" })
            {
                if (item.TryGetProperty(key, out var t) && t.ValueKind == JsonValueKind.String)
                {
                    var s = t.GetString();
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        tracking = s;
                        return true;
                    }
                }
            }

            if (item.TryGetProperty("package_list", out var pkgs) && pkgs.ValueKind == JsonValueKind.Array)
            {
                foreach (var pkg in pkgs.EnumerateArray())
                {
                    if (TryReadTrackingNumberFromLogisticsItem(pkg, out tracking))
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Body get_shipping_document_result / download_shipping_document — sertakan shipping_document_type
        /// (mis. THERMAL_AIR_WAYBILL) agar sama dengan create_shipping_document.
        /// </summary>
        private static object ShopeeLogisticsOrderListResult(
            string orderSn,
            string? packageNumber,
            string shippingDocumentType)
        {
            if (string.IsNullOrWhiteSpace(packageNumber))
            {
                return new
                {
                    order_list = new[]
                    {
                        new { order_sn = orderSn, shipping_document_type = shippingDocumentType }
                    }
                };
            }

            return new
            {
                order_list = new[]
                {
                    new
                    {
                        order_sn = orderSn,
                        package_number = packageNumber,
                        shipping_document_type = shippingDocumentType
                    }
                }
            };
        }

        /// <summary>
        /// Body get_mass_tracking_number — Shopee mengharapkan package_number (bukan order_sn di package_list).
        /// </summary>
        private static object ShopeeLogisticsMassTrackingNumberBody(string packageNumber) =>
            new
            {
                package_list = new[] { new { package_number = packageNumber } },
                response_optional_fields = "first_mile_tracking_number"
            };

        private static string? TryGetDownloadUrl(JsonElement root)
        {
            if (!root.TryGetProperty("response", out var resp))
                return null;
            foreach (var prop in new[] { "url", "file_url", "shipping_document_url", "document_url", "pdf_url" })
            {
                if (resp.TryGetProperty(prop, out var u) && u.ValueKind == JsonValueKind.String)
                {
                    var s = u.GetString();
                    if (!string.IsNullOrWhiteSpace(s))
                        return s;
                }
            }

            return null;
        }

        private enum ShippingDocPoll
        {
            Processing,
            Ready,
            Failed
        }

        private static ShippingDocPoll ClassifyShippingDocResult(JsonElement root, out string? failMessage)
        {
            failMessage = null;
            if (!root.TryGetProperty("response", out var resp))
                return ShippingDocPoll.Processing;

            if (resp.TryGetProperty("result_list", out var rl) && rl.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in rl.EnumerateArray())
                {
                    if (!item.TryGetProperty("status", out var st))
                        continue;
                    var s = st.GetString() ?? "";
                    if (string.Equals(s, "FAILED", StringComparison.OrdinalIgnoreCase))
                    {
                        failMessage = item.TryGetProperty("fail_message", out var fm) ? fm.GetString() : "FAILED";
                        return ShippingDocPoll.Failed;
                    }

                    if (string.Equals(s, "READY", StringComparison.OrdinalIgnoreCase))
                        return ShippingDocPoll.Ready;
                }
            }

            if (resp.TryGetProperty("status", out var st2))
            {
                var s2 = st2.GetString() ?? "";
                if (string.Equals(s2, "READY", StringComparison.OrdinalIgnoreCase))
                    return ShippingDocPoll.Ready;
                if (string.Equals(s2, "FAILED", StringComparison.OrdinalIgnoreCase))
                {
                    failMessage = "FAILED";
                    return ShippingDocPoll.Failed;
                }
            }

            return ShippingDocPoll.Processing;
        }

        private static bool LogisticsFailIsTrackingNumberInvalid(string? failErr, string? failMsg)
        {
            var e = failErr ?? "";
            var m = failMsg ?? "";
            return e.IndexOf("tracking_number_invalid", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   m.IndexOf("tracking number is invalid", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Shopee sering mengembalikan HTTP 200 dengan error top-level (mis. common.batch_api_all_failed) +
        /// detail di response.result_list. Kalau untuk order ini satu-satunya kegagalan adalah
        /// logistics.tracking_number_invalid, alur label boleh lanjut (sinkron resi / create ulang).
        /// </summary>
        private static bool ShopeeBatchResponseOnlyTrackingNumberInvalidForOrder(JsonElement root, string orderSn)
        {
            if (!root.TryGetProperty("error", out var errTop) || errTop.ValueKind != JsonValueKind.String)
                return false;
            var eTop = errTop.GetString() ?? "";
            if (string.IsNullOrWhiteSpace(eTop) || eTop == "0")
                return false;

            if (!root.TryGetProperty("response", out var resp))
                return false;
            if (!resp.TryGetProperty("result_list", out var arr) || arr.ValueKind != JsonValueKind.Array)
                return false;
            if (arr.GetArrayLength() == 0)
                return false;

            var matchedOurOrder = false;
            foreach (var item in arr.EnumerateArray())
            {
                var sn = item.TryGetProperty("order_sn", out var os) && os.ValueKind == JsonValueKind.String
                    ? os.GetString()
                    : null;

                string? failMsg = null;
                string? failErr = null;
                if (item.TryGetProperty("fail_message", out var fm) && fm.ValueKind == JsonValueKind.String)
                    failMsg = fm.GetString();
                if (item.TryGetProperty("fail_error", out var fe))
                {
                    if (fe.ValueKind == JsonValueKind.String)
                        failErr = fe.GetString();
                    else if (fe.ValueKind == JsonValueKind.Number)
                        failErr = fe.GetRawText();
                }

                var hasFail = !string.IsNullOrWhiteSpace(failMsg) || !string.IsNullOrWhiteSpace(failErr);
                var isOur = string.Equals(sn, orderSn, StringComparison.Ordinal);

                if (!isOur)
                {
                    if (hasFail)
                        return false;
                    continue;
                }

                matchedOurOrder = true;
                if (!hasFail)
                    continue;
                if (!LogisticsFailIsTrackingNumberInvalid(failErr, failMsg))
                    return false;
            }

            return matchedOurOrder;
        }

        private static void EnsureShopeeOkOrThrowUnlessRecoverableTrackingBatch(JsonElement root, string orderSn)
        {
            if (ShopeeBatchResponseOnlyTrackingNumberInvalidForOrder(root, orderSn))
                return;
            EnsureShopeeOkOrThrow(root);
        }

        /// <param name="ignoreTrackingNumberInvalid">
        /// Untuk langkah pra-create (mis. get_shipping_document_parameter): Shopee bisa mengembalikan
        /// logistics.tracking_number_invalid padahal langkah berikutnya memang memanggil get_mass_tracking_number
        /// untuk mengisi resi — jangan hentikan alur di sini.
        /// </param>
        private static void ThrowIfLogisticsListItemFailedForOrder(
            JsonElement root,
            string orderSn,
            string apiStep,
            bool ignoreTrackingNumberInvalid = false)
        {
            if (!root.TryGetProperty("response", out var resp))
                return;

            foreach (var listName in new[] { "result_list", "document_list" })
            {
                if (!resp.TryGetProperty(listName, out var arr) || arr.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (var item in arr.EnumerateArray())
                {
                    if (!item.TryGetProperty("order_sn", out var os) || os.GetString() != orderSn)
                        continue;

                    string? failMsg = null;
                    if (item.TryGetProperty("fail_message", out var fm) && fm.ValueKind == JsonValueKind.String)
                        failMsg = fm.GetString();

                    string? failErr = null;
                    if (item.TryGetProperty("fail_error", out var fe))
                    {
                        if (fe.ValueKind == JsonValueKind.String)
                            failErr = fe.GetString();
                        else if (fe.ValueKind == JsonValueKind.Number)
                            failErr = fe.GetRawText();
                    }

                    if (!string.IsNullOrWhiteSpace(failMsg) || !string.IsNullOrWhiteSpace(failErr))
                    {
                        if (ignoreTrackingNumberInvalid && LogisticsFailIsTrackingNumberInvalid(failErr, failMsg))
                            continue;

                        var detail = string.Join(" | ",
                            new[] { failMsg, failErr }.Where(x => !string.IsNullOrWhiteSpace(x)));
                        throw new InvalidOperationException($"{apiStep}: {detail}");
                    }
                }
            }
        }

        private static (string docType, string? packageNumber) ApplyLogisticsParameterForOrder(
            JsonElement root,
            string orderSn,
            string docType,
            string? packageNumber)
        {
            if (!root.TryGetProperty("response", out var resp))
                return (docType, packageNumber);

            foreach (var listName in new[] { "result_list", "document_list" })
            {
                if (!resp.TryGetProperty(listName, out var arr) || arr.ValueKind != JsonValueKind.Array)
                    continue;
                foreach (var item in arr.EnumerateArray())
                {
                    if (!item.TryGetProperty("order_sn", out var os) || os.GetString() != orderSn)
                        continue;

                    if (item.TryGetProperty("package_number", out var pn) && pn.ValueKind == JsonValueKind.String)
                    {
                        var p = pn.GetString();
                        if (!string.IsNullOrWhiteSpace(p))
                            packageNumber = p;
                    }

                    if (item.TryGetProperty("suggest_shipping_document_type", out var sdt))
                    {
                        if (sdt.ValueKind == JsonValueKind.String)
                        {
                            var t = sdt.GetString();
                            if (!string.IsNullOrWhiteSpace(t))
                                docType = t!;
                        }
                        else if (sdt.ValueKind == JsonValueKind.Array && sdt.GetArrayLength() > 0)
                        {
                            var z = sdt[0];
                            if (z.ValueKind == JsonValueKind.String)
                            {
                                var t = z.GetString();
                                if (!string.IsNullOrWhiteSpace(t))
                                    docType = t!;
                            }
                        }
                    }

                    if (item.TryGetProperty("shipping_document_type", out var sdoc) &&
                        sdoc.ValueKind == JsonValueKind.String)
                    {
                        var t = sdoc.GetString();
                        if (!string.IsNullOrWhiteSpace(t))
                            docType = t!;
                    }

                    return (docType, packageNumber);
                }
            }

            if (resp.TryGetProperty("shipping_document_type", out var t0) && t0.ValueKind == JsonValueKind.String)
            {
                var t = t0.GetString();
                if (!string.IsNullOrWhiteSpace(t))
                    docType = t!;
            }
            else if (resp.TryGetProperty("doc_type", out var t1) && t1.ValueKind == JsonValueKind.String)
            {
                var t = t1.GetString();
                if (!string.IsNullOrWhiteSpace(t))
                    docType = t!;
            }

            return (docType, packageNumber);
        }

        /// <summary>
        /// Alur label: package_number dari get_order_detail (package_list), lalu get_mass_tracking_number.
        /// Jika mass API sukses (tanpa gagal di result_list), lanjut — tracking_number dari mass dipakai bila ada;
        /// jika kosong, create_shipping_document tetap dipanggil dengan package + tipe dokumen saja.
        /// </summary>
        private async Task<(string? TrackingNumber, string PackageNumber, string DocType)> RequireLabelTrackingSyncedAsync(
            string orderSn,
            IProgress<SyncProgress> progress,
            CancellationToken ct,
            bool logFullRequestResponse)
        {
            const string docType = "THERMAL_AIR_WAYBILL";

            progress.Report(new SyncProgress
            {
                Log = "[label] (0) get_order_detail (package_list → package_number)…"
            });
            using var detailDoc = await GetShopApiWithLogAsync(
                "/api/v2/order/get_order_detail",
                new Dictionary<string, string>
                {
                    ["order_sn_list"] = orderSn,
                    ["response_optional_fields"] = "package_list"
                },
                progress,
                ct,
                logFullRequestResponse);
            EnsureShopeeOkOrThrow(detailDoc.RootElement);

            var packageNumber = TryExtractPackageNumberForOrder(detailDoc.RootElement, orderSn);
            if (string.IsNullOrWhiteSpace(packageNumber))
                throw new InvalidOperationException(
                    "get_order_detail tidak memuat package_number di package_list untuk order ini. " +
                    "Pastikan order sudah diproses logistik di Shopee dan coba Shopee Sync.");

            progress.Report(new SyncProgress { Log = "[label] (1) get_mass_tracking_number…" });
            var massDoc = await PostShopApiWithLogAsync(
                "/api/v2/logistics/get_mass_tracking_number",
                ShopeeLogisticsMassTrackingNumberBody(packageNumber),
                progress,
                ct,
                logFullRequestResponse);
            EnsureShopeeOkOrThrow(massDoc.RootElement);
            ThrowIfLogisticsListItemFailedForOrder(massDoc.RootElement, orderSn, "get_mass_tracking_number");

            var fromMass = TryExtractTrackingNumberForOrder(massDoc.RootElement, orderSn, packageNumber);
            await Task.Delay(500, ct);
            return (fromMass, packageNumber, docType);
        }

        private async Task<string> DownloadShopeeResiPdfAsync(
            string orderSn,
            IProgress<SyncProgress> progress,
            CancellationToken ct,
            bool logFullRequestResponse = false)
        {
            Directory.CreateDirectory(ResiStorageDir);
            var outPath = Path.Combine(ResiStorageDir, $"{orderSn}_resi.pdf");

            var (trackingForCreateDoc, packageNumber, docType) =
                await RequireLabelTrackingSyncedAsync(orderSn, progress, ct, logFullRequestResponse);
            LogPrint("[label] sinkron selesai — tracking_number: " + (trackingForCreateDoc ?? "(kosong, create pakai package saja)") +
                     $", package_number: {packageNumber}, shipping_document_type: {docType}");

            async Task CreateShippingDocAndPollLoopAsync()
            {
                JsonDocument? createDoc = null;
                try
                {
                    progress.Report(new SyncProgress { Log = "[label] (2) create_shipping_document…" });
                    createDoc = await PostShopApiWithLogAsync(
                        "/api/v2/logistics/create_shipping_document",
                        ShopeeLogisticsOrderListCreate(orderSn, docType, packageNumber, trackingForCreateDoc),
                        progress,
                        ct,
                        logFullRequestResponse);
                    if (!logFullRequestResponse)
                    {
                        var createJson = createDoc.RootElement.GetRawText();
                        LogPrint("[label] create_shipping_document RESPONSE:\n" + createJson);
                        progress.Report(new SyncProgress
                        {
                            Log = "[label] create_shipping_document RESPONSE:\n" + createJson
                        });
                    }

                    EnsureShopeeOkOrThrow(createDoc.RootElement);
                    ThrowIfLogisticsListItemFailedForOrder(createDoc.RootElement, orderSn, "create_shipping_document");

                    progress.Report(new SyncProgress
                    {
                        Log = "[label] tunggu dokumen siap (get_shipping_document_result)…"
                    });
                    for (var i = 0; i < 45; i++)
                    {
                        await Task.Delay(600, ct);
                        var resDoc = await PostShopApiWithLogAsync(
                            "/api/v2/logistics/get_shipping_document_result",
                            ShopeeLogisticsOrderListResult(orderSn, packageNumber, docType),
                            progress,
                            ct,
                            logFullRequestResponse);
                        EnsureShopeeOkOrThrow(resDoc.RootElement);
                        var poll = ClassifyShippingDocResult(resDoc.RootElement, out var fail);
                        if (poll == ShippingDocPoll.Failed)
                            throw new InvalidOperationException(fail ?? "Gagal membuat dokumen resi.");
                        if (poll == ShippingDocPoll.Ready)
                            break;

                        if (i == 44)
                            throw new TimeoutException("Resi belum siap (get_shipping_document_result). Coba lagi nanti.");
                    }
                }
                finally
                {
                    createDoc?.Dispose();
                }
            }

            async Task<string> DownloadPdfToOutPathAsync()
            {
                JsonDocument? dlDoc = null;
                try
                {
                    progress.Report(new SyncProgress { Log = "[label] (3) download_shipping_document…" });
                    var (savedRawPdf, doc) = await PostDownloadShippingDocumentWithLogAsync(
                        ShopeeLogisticsOrderListResult(orderSn, packageNumber, docType),
                        outPath,
                        progress,
                        ct,
                        logFullRequestResponse);
                    if (savedRawPdf)
                        return outPath;

                    dlDoc = doc ?? throw new InvalidOperationException("Respons unduh label kosong.");
                    EnsureShopeeOkOrThrow(dlDoc.RootElement);

                    var url = TryGetDownloadUrl(dlDoc.RootElement);
                    if (string.IsNullOrWhiteSpace(url))
                    {
                        if (dlDoc.RootElement.TryGetProperty("response", out var resp) &&
                            resp.TryGetProperty("file_base64", out var b64) &&
                            b64.ValueKind == JsonValueKind.String)
                        {
                            var raw = b64.GetString();
                            if (!string.IsNullOrWhiteSpace(raw))
                            {
                                if (logFullRequestResponse)
                                {
                                    progress.Report(new SyncProgress
                                    {
                                        Log =
                                            $"[API][RESP] file_base64: PDF didekode dari JSON response ({raw.Length} chars base64) → disimpan ke disk"
                                    });
                                }

                                var bytes = Convert.FromBase64String(raw);
                                await File.WriteAllBytesAsync(outPath, bytes, ct);
                                return outPath;
                            }
                        }

                        throw new InvalidOperationException("Respons download tidak berisi URL/PDF.");
                    }

                    if (logFullRequestResponse)
                    {
                        progress.Report(new SyncProgress
                        {
                            Log = $"[API][REQ] GET (unduh file label)\n{RedactSensitiveShopeeUrlForLog(url)}"
                        });
                    }

                    using var pdfResp = await _shopeeHttp.GetAsync(url, ct);
                    if (logFullRequestResponse)
                    {
                        var len = pdfResp.Content.Headers.ContentLength;
                        progress.Report(new SyncProgress
                        {
                            Log =
                                $"[API][RESP] {(int)pdfResp.StatusCode} {pdfResp.ReasonPhrase}\n[body: stream PDF → file; Content-Length: {(len.HasValue ? len.Value.ToString(CultureInfo.InvariantCulture) : "?")} bytes]"
                        });
                    }

                    pdfResp.EnsureSuccessStatusCode();
                    await using var fs = File.Create(outPath);
                    await pdfResp.Content.CopyToAsync(fs, ct);

                    return outPath;
                }
                finally
                {
                    dlDoc?.Dispose();
                }
            }

            await CreateShippingDocAndPollLoopAsync();

            try
            {
                return await DownloadPdfToOutPathAsync();
            }
            catch (Exception ex) when (ShopeeExceptionIsShippingDocPrintFirst(ex))
            {
                LogPrint(
                    "[label] download: shipping_document_should_print_first — create ulang lalu unduh lagi.");
                await CreateShippingDocAndPollLoopAsync();
                return await DownloadPdfToOutPathAsync();
            }
        }

        private void DbUpsertOrderResiPdf(string orderSn, string pdfPath)
        {
            using var con = OpenDb();
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText = @"
INSERT INTO order_resi(order_sn, pdf_path, fetched_at)
VALUES($sn,$p,$t)
ON CONFLICT(order_sn) DO UPDATE SET
  pdf_path=excluded.pdf_path,
  fetched_at=excluded.fetched_at;";
            cmd.Parameters.AddWithValue("$sn", orderSn);
            cmd.Parameters.AddWithValue("$p", pdfPath);
            cmd.Parameters.AddWithValue("$t", UnixNow());
            cmd.ExecuteNonQuery();
        }

        private void DbSetResiPrinted(string orderSn, bool printed)
        {
            using var con = OpenDb();
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText = @"
INSERT INTO order_resi(order_sn, resi_printed, resi_printed_at)
VALUES($sn,$p,$t)
ON CONFLICT(order_sn) DO UPDATE SET
  resi_printed=excluded.resi_printed,
  resi_printed_at=excluded.resi_printed_at;";
            cmd.Parameters.AddWithValue("$sn", orderSn);
            cmd.Parameters.AddWithValue("$p", printed ? 1 : 0);
            cmd.Parameters.AddWithValue("$t", printed ? UnixNow() : (object?)null);
            cmd.ExecuteNonQuery();
        }

        private int GetResiOrderCount()
        {
            var fromUnix = DateTimeOffset.UtcNow.AddDays(-ResiListDaysBack).ToUnixTimeSeconds();
            using var con = OpenDb();
            con.Open();
            using var cmd = con.CreateCommand();
            var where = "o.create_time >= $from ";
            switch (_resiTabFilter)
            {
                case ResiTabFilter.NotPrintedResi:
                    where += "AND (COALESCE(r.resi_printed,0) = 0) ";
                    break;
                case ResiTabFilter.PrintedResi:
                    where += "AND COALESCE(r.resi_printed,0) = 1 ";
                    break;
            }

            cmd.CommandText = $@"SELECT COUNT(*) FROM orders o
LEFT JOIN order_resi r ON r.order_sn = o.order_sn
WHERE {where};";
            cmd.Parameters.AddWithValue("$from", fromUnix);
            var scalar = cmd.ExecuteScalar();
            return scalar is long l ? (int)l : Convert.ToInt32(scalar ?? 0);
        }

        private void LoadResiPageFromDb(int pageIndex)
        {
            _resiTotalItems = GetResiOrderCount();
            var totalPages = Math.Max(1, (int)Math.Ceiling(_resiTotalItems / (double)ResiPageSize));
            if (pageIndex < 0) pageIndex = 0;
            if (pageIndex >= totalPages) pageIndex = totalPages - 1;
            _resiPageIndex = pageIndex;

            switch (_resiTabFilter)
            {
                case ResiTabFilter.All:
                    _resiPageIndexAll = _resiPageIndex;
                    break;
                case ResiTabFilter.NotPrintedResi:
                    _resiPageIndexNotPrinted = _resiPageIndex;
                    break;
                case ResiTabFilter.PrintedResi:
                    _resiPageIndexPrinted = _resiPageIndex;
                    break;
            }

            var fromUnix = DateTimeOffset.UtcNow.AddDays(-ResiListDaysBack).ToUnixTimeSeconds();
            var skip = _resiPageIndex * ResiPageSize;

            var where = "o.create_time >= $from ";
            switch (_resiTabFilter)
            {
                case ResiTabFilter.NotPrintedResi:
                    where += "AND (COALESCE(r.resi_printed,0) = 0) ";
                    break;
                case ResiTabFilter.PrintedResi:
                    where += "AND COALESCE(r.resi_printed,0) = 1 ";
                    break;
            }

            ResiRows.Clear();
            using var con = OpenDb();
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText = $@"SELECT o.order_sn, o.create_time, r.pdf_path, COALESCE(r.resi_printed,0)
FROM orders o
LEFT JOIN order_resi r ON r.order_sn = o.order_sn
WHERE {where}
ORDER BY o.create_time DESC
LIMIT $take OFFSET $skip;";
            cmd.Parameters.AddWithValue("$from", fromUnix);
            cmd.Parameters.AddWithValue("$take", ResiPageSize);
            cmd.Parameters.AddWithValue("$skip", skip);

            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                var sn = rd.GetString(0);
                var ct = rd.IsDBNull(1) ? 0 : rd.GetInt64(1);
                var pdf = rd.IsDBNull(2) ? null : rd.GetString(2);
                var rp = rd.IsDBNull(3) ? 0 : rd.GetInt32(3);
                ResiRows.Add(new ResiRow
                {
                    OrderSn = sn,
                    CreateTimeUnix = ct,
                    PdfPath = string.IsNullOrWhiteSpace(pdf) ? null : pdf,
                    ResiPrinted = rp == 1
                });
            }

            UpdateResiPagingUi();
        }

        private void UpdateResiPagingUi()
        {
            var totalPages = Math.Max(1, (int)Math.Ceiling(_resiTotalItems / (double)ResiPageSize));
            if (TxtResiPage != null)
                TxtResiPage.Text = $"Page {_resiPageIndex + 1} / {totalPages}  (Total: {_resiTotalItems})";
            if (BtnResiPrev != null)
                BtnResiPrev.IsEnabled = _resiPageIndex > 0;
            if (BtnResiNext != null)
                BtnResiNext.IsEnabled = (_resiPageIndex + 1) < totalPages;
        }

        private void MainWorkspaceTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || !ReferenceEquals(e.OriginalSource, MainWorkspaceTabs))
                return;

            UpdatePreviewChromeForActiveTab();

            if (MainWorkspaceTabs.SelectedIndex == 1)
            {
                LoadResiPageFromDb(_resiPageIndex);
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    ResiSynchronizePreviewForRow(ResiGrid.SelectedItem as ResiRow);
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
            else
            {
                TxtPreviewHint.Visibility = Visibility.Collapsed;
                if (MainWorkspaceTabs.SelectedIndex == 0 && QueueGrid.SelectedItem is JobRow jr)
                {
                    _lastUserPickedForPreview = jr;
                    if (!_suppressPreviewWhileScrolling)
                        StartPreviewLatest(jr);
                }
            }
        }

        private void UpdatePreviewChromeForActiveTab()
        {
            if (TxtPdfPreviewTitle == null)
                return;
            if (MainWorkspaceTabs?.SelectedIndex == 1)
                TxtPdfPreviewTitle.Text = "Pratinjau label pengiriman";
            else
                TxtPdfPreviewTitle.Text = "PDF Preview";
        }

        /// <summary>
        /// Tampilkan PDF label di panel kanan; kosongkan + petunjuk jika belum ada file.
        /// </summary>
        private void ResiSynchronizePreviewForRow(ResiRow? rr)
        {
            if (TxtPreviewHint == null)
                return;

            try
            {
                _previewCts?.Cancel();
            }
            catch { }

            if (rr == null)
            {
                Thumbs.ItemsSource = null;
                HidePreviewLoading();
                TxtPreviewHint.Text = "Pilih satu order di tabel label pengiriman. Yang sudah punya PDF akan tampil di sini.";
                TxtPreviewHint.Visibility = Visibility.Visible;
                return;
            }

            if (string.IsNullOrWhiteSpace(rr.PdfPath) || !File.Exists(rr.PdfPath))
            {
                Thumbs.ItemsSource = null;
                HidePreviewLoading();
                TxtPreviewHint.Text =
                    $"Order {rr.OrderSn}: belum ada PDF di disk — klik «Ambil label (PDF)» atau «Ambil PDF» di baris ini.";
                TxtPreviewHint.Visibility = Visibility.Visible;
                return;
            }

            TxtPreviewHint.Visibility = Visibility.Collapsed;
            _resiPreviewStub.File = rr.PdfPath!;
            _resiPreviewStub.Printer = (CmbResiPrinter.SelectedItem as string) ?? Printers.FirstOrDefault() ?? "";
            _lastUserPickedForPreview = _resiPreviewStub;
            if (!_suppressPreviewWhileScrolling)
                StartPreviewLatest(_resiPreviewStub);
        }

        private static bool IsVisualDescendantOf(DependencyObject? node, DependencyObject? ancestor)
        {
            while (node != null)
            {
                if (ReferenceEquals(node, ancestor))
                    return true;
                node = VisualTreeHelper.GetParent(node);
            }
            return false;
        }

        private void EnsureResiGridInSelectedTab()
        {
            if (ResiGrid == null || ResiFilterTabs == null)
                return;
            var idx = ResiFilterTabs.SelectedIndex;
            if (idx < 0)
                return;
            var hosts = new[] { ResiGridHost0, ResiGridHost1, ResiGridHost2 };
            if (idx >= hosts.Length || hosts[idx] == null)
                return;
            var target = hosts[idx];
            if (ReferenceEquals(ResiGrid.Parent, target))
                return;
            if (ResiGrid.Parent is Panel p)
                p.Children.Remove(ResiGrid);
            target.Children.Add(ResiGrid);
        }

        private void ResiFilterTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // SelectionChanged pada DataGrid bubble ke TabControl — jangan anggap sebagai ganti tab.
            if (ResiGrid != null && e.OriginalSource is DependencyObject src &&
                IsVisualDescendantOf(src, ResiGrid))
                return;

            EnsureResiGridInSelectedTab();

            if (!IsLoaded || ResiFilterTabs == null)
                return;

            switch (_resiTabFilter)
            {
                case ResiTabFilter.All:
                    _resiPageIndexAll = _resiPageIndex;
                    break;
                case ResiTabFilter.NotPrintedResi:
                    _resiPageIndexNotPrinted = _resiPageIndex;
                    break;
                case ResiTabFilter.PrintedResi:
                    _resiPageIndexPrinted = _resiPageIndex;
                    break;
            }

            switch (ResiFilterTabs.SelectedIndex)
            {
                case 0:
                    _resiTabFilter = ResiTabFilter.All;
                    _resiPageIndex = _resiPageIndexAll;
                    break;
                case 1:
                default:
                    _resiTabFilter = ResiTabFilter.NotPrintedResi;
                    _resiPageIndex = _resiPageIndexNotPrinted;
                    break;
                case 2:
                    _resiTabFilter = ResiTabFilter.PrintedResi;
                    _resiPageIndex = _resiPageIndexPrinted;
                    break;
            }

            LoadResiPageFromDb(_resiPageIndex);
        }

        private void BtnResiRefresh_Click(object sender, RoutedEventArgs e) => LoadResiPageFromDb(_resiPageIndex);

        private void BtnResiSelectAllPage_Click(object sender, RoutedEventArgs e)
        {
            ResiGrid.Focus();
            ResiGrid.SelectAll();
        }

        private void BtnResiClearChecks_Click(object sender, RoutedEventArgs e) => ResiGrid.UnselectAll();

        /// <summary>
        /// Menjaga anchor Shift+klik di DataGrid setelah <see cref="DataGrid.SelectedItem"/> di-set dari kode.
        /// </summary>
        private void ResiGridSyncCurrentCell(ResiRow? row)
        {
            if (row == null || ResiGrid.Columns.Count == 0)
                return;
            ResiGrid.CurrentCell = new DataGridCellInfo(row, ResiGrid.Columns[0]);
        }

        private List<ResiRow> GetResiRowsTargetedForPdfFetch() =>
            ResiGrid.SelectedItems.Cast<ResiRow>().GroupBy(r => r.OrderSn).Select(g => g.First()).ToList();

        private void ResiBulkFetchUi(Action action)
        {
            if (Dispatcher.CheckAccess())
                action();
            else
                Dispatcher.Invoke(action);
        }

        private void ShowResiBulkFetchOverlay(int total, string? title = null, string? initialDetail = null)
        {
            if (ResiBulkFetchOverlay == null || PbResiBulk == null)
                return;
            var n = Math.Max(1, total);
            PbResiBulk.Maximum = n;
            PbResiBulk.Value = 0;
            if (TxtResiBulkTitle != null)
                TxtResiBulkTitle.Text = string.IsNullOrEmpty(title) ? "Mengunduh label PDF" : title;
            var progress0 = $"0 / {total} selesai diproses · 0 berhasil · 0 gagal";
            TxtResiBulkProgressLine!.Text = progress0;
            TxtResiBulkProgressLine.ToolTip = progress0;
            var detail0 = initialDetail ?? (total <= 1
                ? "Mengunduh 1 label dari Shopee…"
                : $"Akan mengunduh {total} label dari Shopee (berurutan).");
            TxtResiBulkDetail!.Text = detail0;
            TxtResiBulkDetail.ToolTip = detail0;
            ResiBulkFetchOverlay.Visibility = Visibility.Visible;
        }

        private void UpdateResiBulkFetchOverlay(int completedProcessed, int total, int okCount, int failCount, string detail)
        {
            if (PbResiBulk == null)
                return;
            PbResiBulk.Value = Math.Min(completedProcessed, PbResiBulk.Maximum);
            var progressLine =
                $"{completedProcessed} / {total} selesai diproses · {okCount} berhasil · {failCount} gagal";
            TxtResiBulkProgressLine!.Text = progressLine;
            TxtResiBulkProgressLine.ToolTip = progressLine;
            var d = detail.Trim();
            TxtResiBulkDetail!.Text = d;
            TxtResiBulkDetail.ToolTip = string.IsNullOrEmpty(d) ? null : d;
        }

        private void HideResiBulkFetchOverlay()
        {
            if (ResiBulkFetchOverlay != null)
                ResiBulkFetchOverlay.Visibility = Visibility.Collapsed;
        }

        private void BtnResiPrev_Click(object sender, RoutedEventArgs e) => LoadResiPageFromDb(_resiPageIndex - 1);

        private void BtnResiNext_Click(object sender, RoutedEventArgs e) => LoadResiPageFromDb(_resiPageIndex + 1);

        private void ResiGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            e.Handled = true;
            if (MainWorkspaceTabs.SelectedIndex != 1)
                return;
            // Saat multi-select, pratinjau mengikuti fokus utama (baris terakhir dipilih).
            ResiSynchronizePreviewForRow(ResiGrid.SelectedItem as ResiRow);
        }

        private async void BtnResiFetchPdf_Click(object sender, RoutedEventArgs e)
        {
            var rows = GetResiRowsTargetedForPdfFetch();
            if (rows.Count == 0)
            {
                MessageBox.Show(
                    "Sorot satu atau beberapa baris (biru): klik untuk satu baris, Ctrl+klik untuk banyak, Shift+klik untuk rentang — seperti Excel. Lalu klik «Ambil label (PDF)».",
                    "Label pengiriman",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            if (rows.Count == 1)
                await FetchResiForRowAsync(rows[0], showSuccessDialog: true);
            else
                await FetchResiManyAsync(rows);
        }

        private async void ResiRowFetch_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is not ResiRow row)
                return;
            await FetchResiForRowAsync(row, showSuccessDialog: true);
        }

        private static string FormatTrackingInvalidHint()
        {
            return "\n\nNomor resi di sistem Shopee belum valid atau belum tersinkron dari kurir.\n" +
                   "• Ambil label (PDF): (0) get_order_detail → package_number, lalu (1) get_mass_tracking_number — jika sukses lanjut create; tracking dari mass dipakai jika ada.\n" +
                   "• Di Seller Centre: pastikan pengiriman sudah diajukan (pickup/dropoff) sesuai channel.\n" +
                   "• Kurir non‑integrasi: isi nomor resi yang benar di Shopee.\n" +
                   "• Lihat Printing log untuk baris [label] (1) / (2) / (3).\n" +
                   "• Tunggu 1–2 menit setelah arrange / resi muncul di kurir, lalu Shopee Sync + coba lagi.";
        }

        private async Task<(bool Ok, string? Error, string? Path)> FetchResiPdfCoreAsync(
            ResiRow row,
            IProgress<SyncProgress> progress,
            CancellationToken ct)
        {
            try
            {
                var path = await DownloadShopeeResiPdfAsync(row.OrderSn, progress, ct, logFullRequestResponse: true);
                DbUpsertOrderResiPdf(row.OrderSn, path);
                return (true, null, path);
            }
            catch (Exception ex)
            {
                return (false, ex.Message, null);
            }
        }

        private async Task FetchResiForRowAsync(ResiRow row, bool showSuccessDialog)
        {
            if (!IsConnected())
            {
                MessageBox.Show("Hubungkan Shopee dulu (Connect Shopee).", "Label pengiriman", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            IsEnabled = false;
            try
            {
                var progress = new Progress<SyncProgress>(p =>
                {
                    if (!string.IsNullOrWhiteSpace(p.Log))
                        LogPrint(p.Log);
                });
                var (ok, err, path) = await FetchResiPdfCoreAsync(row, progress, CancellationToken.None);
                LoadResiPageFromDb(_resiPageIndex);
                if (ok)
                {
                    if (showSuccessDialog)
                        MessageBox.Show("PDF label tersimpan:\n" + path, "Label pengiriman", MessageBoxButton.OK, MessageBoxImage.Information);
                    var refreshed = ResiRows.FirstOrDefault(r => r.OrderSn == row.OrderSn);
                    if (refreshed != null)
                    {
                        ResiGrid.SelectedItem = refreshed;
                        ResiGridSyncCurrentCell(refreshed);
                        ResiSynchronizePreviewForRow(refreshed);
                    }
                }
                else
                {
                    var msg = "Gagal ambil PDF label pengiriman:\n" + err;
                    var m = err ?? "";
                    if (m.Contains("tracking_number_invalid", StringComparison.OrdinalIgnoreCase) ||
                        m.Contains("tracking number is invalid", StringComparison.OrdinalIgnoreCase) ||
                        m.Contains("Sinkron resi gagal", StringComparison.OrdinalIgnoreCase))
                        msg += FormatTrackingInvalidHint();
                    MessageBox.Show(msg, "Label pengiriman", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            finally
            {
                IsEnabled = true;
            }
        }

        private async Task FetchResiManyAsync(IReadOnlyList<ResiRow> rows)
        {
            if (!IsConnected())
            {
                MessageBox.Show("Hubungkan Shopee dulu (Connect Shopee).", "Label pengiriman", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var total = rows.Count;
            var ok = 0;
            var failed = new List<string>();

            ResiBulkFetchUi(() => ShowResiBulkFetchOverlay(total));

            try
            {
                for (var i = 0; i < rows.Count; i++)
                {
                    var row = rows[i];
                    var idx = i + 1;

                    ResiBulkFetchUi(() => UpdateResiBulkFetchOverlay(
                        i,
                        total,
                        ok,
                        failed.Count,
                        $"Sedang mengambil ({idx}/{total}): {row.OrderSn} — sinkron resi, create dokumen, unduh PDF (API Shopee)…"));

                    var progress = new Progress<SyncProgress>(p =>
                    {
                        if (string.IsNullOrWhiteSpace(p.Log))
                            return;
                        LogPrint(p.Log);
                        var line = p.Log.Trim();
                        if (line.Length > 180)
                            line = line[..180] + "…";
                        ResiBulkFetchUi(() => UpdateResiBulkFetchOverlay(i, total, ok, failed.Count,
                            $"{row.OrderSn}: {line}"));
                    });

                    var (success, err, _) = await FetchResiPdfCoreAsync(row, progress, CancellationToken.None);
                    if (success)
                        ok++;
                    else
                        failed.Add($"{row.OrderSn}: {err}");

                    ResiBulkFetchUi(() => UpdateResiBulkFetchOverlay(
                        i + 1,
                        total,
                        ok,
                        failed.Count,
                        success
                            ? $"Selesai ({idx}/{total}): {row.OrderSn} — PDF tersimpan."
                            : $"Gagal ({idx}/{total}): {row.OrderSn} — {err}"));
                }

                LoadResiPageFromDb(_resiPageIndex);
                ResiSynchronizePreviewForRow(ResiGrid.SelectedItem as ResiRow);

                var summary = $"Selesai mengunduh {rows.Count} order.\nBerhasil: {ok}\nGagal: {failed.Count}";
                if (failed.Count > 0)
                {
                    summary += "\n\n" + string.Join("\n", failed.Take(8));
                    if (failed.Count > 8)
                        summary += $"\n… dan {failed.Count - 8} lainnya.";
                }

                MessageBox.Show(summary, "Label pengiriman", MessageBoxButton.OK,
                    failed.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            finally
            {
                ResiBulkFetchUi(HideResiBulkFetchOverlay);
            }
        }

        private async void BtnResiPrint_Click(object sender, RoutedEventArgs e)
        {
            var rows = GetResiRowsTargetedForPdfFetch();
            if (rows.Count == 0)
            {
                MessageBox.Show(
                    "Sorot satu atau beberapa baris (biru) yang sudah punya PDF label, lalu klik «Cetak label».",
                    "Label pengiriman",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var withPdf = rows.Where(r => !string.IsNullOrWhiteSpace(r.PdfPath) && File.Exists(r.PdfPath)).ToList();
            if (withPdf.Count == 0)
            {
                MessageBox.Show(
                    "Belum ada file PDF untuk baris yang dipilih. Klik «Ambil label (PDF)» dulu.",
                    "Label pengiriman",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (withPdf.Count < rows.Count)
            {
                MessageBox.Show(
                    $"{rows.Count - withPdf.Count} baris tanpa file PDF dilewati. Akan mencetak {withPdf.Count} label.",
                    "Label pengiriman",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }

            if (string.IsNullOrWhiteSpace((CmbResiPrinter.SelectedItem as string)?.Trim()))
            {
                MessageBox.Show("Pilih printer untuk label.", "Label pengiriman", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await PrintResiManyAsync(withPdf);
        }

        private async void ResiRowPrint_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is not ResiRow row)
                return;
            await PrintResiRowAsync(row);
        }

        /// <summary>Cetak satu label tanpa MessageBox; untuk batch. Mengembalikan alasan gagal jika tidak Done.</summary>
        private async Task<(bool Ok, string? FailReason)> PrintResiRowQuietAsync(ResiRow row, bool reloadPage, double pdfPrintScale)
        {
            if (string.IsNullOrWhiteSpace(row.PdfPath) || !File.Exists(row.PdfPath))
                return (false, "PDF tidak ada");

            var printer = (CmbResiPrinter.SelectedItem as string)?.Trim();
            if (string.IsNullOrWhiteSpace(printer))
                return (false, "Printer belum dipilih");

            var temp = new JobRow
            {
                File = row.PdfPath!,
                Printer = printer,
                PageFrom = 1,
                PageTo = 0,
                Copies = 1,
                Duplex = DuplexMode.Simplex,
                Paper = PaperPreset.A6,
                PrintMonochrome = true,
                PrintFromTopLeft = true,
                OrderProcessId = 0,
                Status = "Ready",
                PdfPrintScale = pdfPrintScale
            };

            await PrintAsync(temp, markShopeeLineAsPrinted: false);
            var ok = string.Equals(temp.Status, "Done", StringComparison.OrdinalIgnoreCase);
            if (ok)
            {
                DbSetResiPrinted(row.OrderSn, true);
                if (reloadPage)
                    LoadResiPageFromDb(_resiPageIndex);
            }

            return (ok, ok ? null : (temp.Status ?? "Gagal cetak"));
        }

        private async Task PrintResiRowAsync(ResiRow row, bool reloadPage = true, double? pdfPrintScale = null)
        {
            if (string.IsNullOrWhiteSpace(row.PdfPath) || !File.Exists(row.PdfPath))
            {
                MessageBox.Show("Belum ada file PDF label. Klik Ambil label (PDF) dulu.", "Label pengiriman", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace((CmbResiPrinter.SelectedItem as string)?.Trim()))
            {
                MessageBox.Show("Pilih printer untuk label.", "Label pengiriman", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var scale = pdfPrintScale ?? ResiLabelPdfPrintScale;
            var (ok, err) = await PrintResiRowQuietAsync(row, reloadPage, scale);
            if (!ok)
            {
                MessageBox.Show(
                    string.IsNullOrWhiteSpace(err) ? "Gagal mencetak label." : err,
                    "Label pengiriman",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void ShowResiInlinePrintProgress(int total, string initialDetail)
        {
            if (ResiPrintProgressPanel == null || PbResiPrintProgress == null || TxtResiPrintStatus == null)
                return;
            var n = Math.Max(1, total);
            PbResiPrintProgress.Maximum = n;
            PbResiPrintProgress.Value = 0;
            TxtResiPrintStatus.Text = initialDetail;
            ResiPrintProgressPanel.Visibility = Visibility.Visible;
        }

        private void UpdateResiInlinePrintProgress(int completedProcessed, int total, int okCount, int failCount, string detail)
        {
            if (PbResiPrintProgress == null || TxtResiPrintStatus == null)
                return;
            PbResiPrintProgress.Value = Math.Min(completedProcessed, PbResiPrintProgress.Maximum);
            var d = detail.Trim();
            if (d.Length > 200)
                d = d[..200] + "…";
            TxtResiPrintStatus.Text = $"{completedProcessed}/{total} · berhasil {okCount} · gagal {failCount}" +
                                     (string.IsNullOrEmpty(d) ? "" : $" — {d}");
        }

        private void HideResiInlinePrintProgress()
        {
            if (TxtResiPrintStatus != null)
                TxtResiPrintStatus.Text = string.Empty;
            if (PbResiPrintProgress != null)
                PbResiPrintProgress.Value = 0;
            if (ResiPrintProgressPanel != null)
                ResiPrintProgressPanel.Visibility = Visibility.Collapsed;
        }

        private async Task PrintResiManyAsync(IReadOnlyList<ResiRow> rows)
        {
            var total = rows.Count;
            var ok = 0;
            var failed = new List<string>();
            var scaleHint = ResiLabelPdfPrintScale.ToString("0.##", CultureInfo.CurrentCulture);

            ResiBulkFetchUi(() => ShowResiInlinePrintProgress(
                total,
                total <= 1
                    ? $"Mencetak 1 label (skala {scaleHint}×)…"
                    : $"Mencetak {total} label (skala {scaleHint}×)…"));

            try
            {
                for (var i = 0; i < rows.Count; i++)
                {
                    var row = rows[i];
                    var idx = i + 1;

                    ResiBulkFetchUi(() => UpdateResiInlinePrintProgress(
                        i,
                        total,
                        ok,
                        failed.Count,
                        $"{row.OrderSn}…"));

                    var (success, err) = await PrintResiRowQuietAsync(row, reloadPage: false, ResiLabelPdfPrintScale);

                    if (success)
                        ok++;
                    else
                        failed.Add($"{row.OrderSn}: {err ?? "gagal"}");

                    ResiBulkFetchUi(() => UpdateResiInlinePrintProgress(
                        i + 1,
                        total,
                        ok,
                        failed.Count,
                        success ? $"{row.OrderSn} ok" : $"{row.OrderSn}: {err}"));
                }

                LoadResiPageFromDb(_resiPageIndex);
                ResiSynchronizePreviewForRow(ResiGrid.SelectedItem as ResiRow);

                if (failed.Count > 0)
                    LogPrint("Cetak label — gagal: " + string.Join("; ", failed));
            }
            finally
            {
                ResiBulkFetchUi(HideResiInlinePrintProgress);
            }
        }

        private void ResiRowTogglePrinted_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is not ResiRow row)
                return;
            DbSetResiPrinted(row.OrderSn, !row.ResiPrinted);
            LoadResiPageFromDb(_resiPageIndex);
        }

        




        public class ShopeeConfig
        {
            public string BaseUrl { get; set; } = DefaultShopeeBaseUrl;

            // Credentials from Shopee Open Platform
            public long PartnerId { get; set; }
            public string PartnerKey { get; set; } = "";
            public long ShopId { get; set; }
            public string AccessToken { get; set; } = "";

            // Query defaults
            public string OrderStatus { get; set; } = "READY_TO_SHIP";

            // Unix time seconds
            public long TimeFromUnix { get; set; } = DateTimeOffset.UtcNow.AddDays(-2).ToUnixTimeSeconds();
            public long TimeToUnix { get; set; } = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        }

        private sealed class ShopeeClient
        {
            private readonly ShopeeConfig _cfg;
            private readonly HttpClient _http;

            public ShopeeClient(ShopeeConfig cfg)
            {
                _cfg = cfg;
                _http = new HttpClient
                {
                    BaseAddress = new Uri((_cfg.BaseUrl ?? DefaultShopeeBaseUrl).TrimEnd('/') + "/")
                };
                _http.Timeout = TimeSpan.FromSeconds(45);
            }

            private string Sign(string apiPath, long timestamp)
            {
                // Shopee v2 signature: HMAC-SHA256(partner_key, partner_id + api_path + timestamp + access_token + shop_id)
                // ref: community examples & SDKs
                var baseStr = $"{_cfg.PartnerId}{apiPath}{timestamp}{_cfg.AccessToken}{_cfg.ShopId}";
                using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(_cfg.PartnerKey ?? ""));
                var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(baseStr));
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }

            private async Task<T> GetAsync<T>(string apiPath, Dictionary<string, string?> query, CancellationToken ct = default)
            {
                var ts = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
                query["partner_id"] = _cfg.PartnerId.ToString(CultureInfo.InvariantCulture);
                query["timestamp"] = ts.ToString(CultureInfo.InvariantCulture);
                query["shop_id"] = _cfg.ShopId.ToString(CultureInfo.InvariantCulture);
                query["access_token"] = _cfg.AccessToken;
                query["sign"] = Sign(apiPath, ts);

                var qs = string.Join("&", query
                    .Where(kv => !string.IsNullOrWhiteSpace(kv.Value))
                    .Select(kv => $"{Uri.EscapeDataString(kv.Key)}={Uri.EscapeDataString(kv.Value!)}"));

                var url = apiPath.TrimStart('/') + "?" + qs;

                using var resp = await _http.GetAsync(url, ct);
                var body = await resp.Content.ReadAsStringAsync(ct);

                if (!resp.IsSuccessStatusCode)
                    throw new InvalidOperationException($"HTTP {(int)resp.StatusCode} {resp.ReasonPhrase}\n{body}");

                var obj = JsonSerializer.Deserialize<T>(body, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (obj == null)
                    throw new InvalidOperationException("Failed to parse Shopee response:\n" + body);

                return obj;
            }

            public async Task<List<ShopeeOrderDetail>> GetOrderDetailsAsync(long timeFromUnix, long timeToUnix, string orderStatus)
            {
                var listResp = await GetAsync<ShopeeOrderListResponse>(
                    "/api/v2/order/get_order_list",
                    new Dictionary<string, string?>
                    {
                        ["time_range_field"] = "create_time",
                        ["time_from"] = timeFromUnix.ToString(CultureInfo.InvariantCulture),
                        ["time_to"] = timeToUnix.ToString(CultureInfo.InvariantCulture),
                        ["page_size"] = "50",
                        ["order_status"] = orderStatus
                    });

                var ordersn = listResp?.Response?.OrderList?.Select(x => x.OrderSn).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().ToList()
                             ?? new List<string>();

                if (ordersn.Count == 0) return new List<ShopeeOrderDetail>();

                // Shopee allows ordersn_list comma separated
                var detailResp = await GetAsync<ShopeeOrderDetailResponse>(
                    "/api/v2/order/get_order_detail",
                    new Dictionary<string, string?>
                    {
                        ["order_sn_list"] = string.Join(",", ordersn),
                        ["response_optional_fields"] = "item_list,buyer_username,shipping_carrier"
                    });

                return detailResp?.Response?.OrderList ?? new List<ShopeeOrderDetail>();
            }
        }

        // --- DTOs (Shopee OpenAPI v2) ---
        private sealed class ShopeeOrderListResponse
        {
            [JsonPropertyName("error")] public string? Error { get; set; }
            [JsonPropertyName("message")] public string? Message { get; set; }
            [JsonPropertyName("response")] public ShopeeOrderListData? Response { get; set; }
        }

        private sealed class ShopeeOrderListData
        {
            [JsonPropertyName("order_list")] public List<ShopeeOrderBrief>? OrderList { get; set; }
        }

        private sealed class ShopeeOrderBrief
        {
            [JsonPropertyName("order_sn")] public string? OrderSn { get; set; }
        }

        private sealed class ShopeeOrderDetailResponse
        {
            [JsonPropertyName("error")] public string? Error { get; set; }
            [JsonPropertyName("message")] public string? Message { get; set; }
            [JsonPropertyName("response")] public ShopeeOrderDetailData? Response { get; set; }
        }

        private sealed class ShopeeOrderDetailData
        {
            [JsonPropertyName("order_list")] public List<ShopeeOrderDetail>? OrderList { get; set; }
        }

        public sealed class ShopeeOrderDetail
        {
            [JsonPropertyName("order_sn")] public string? OrderSn { get; set; }
            [JsonPropertyName("create_time")] public long? CreateTimeUnix { get; set; }
            [JsonPropertyName("item_list")] public List<ShopeeOrderItem> ItemList { get; set; } = new();

            [JsonIgnore]
            public DateTime? CreateTimeLocal =>
                CreateTimeUnix.HasValue ? DateTimeOffset.FromUnixTimeSeconds(CreateTimeUnix.Value).LocalDateTime : null;
        }

        public sealed class ShopeeOrderItem
        {
            [JsonPropertyName("item_name")] public string? ItemName { get; set; }

            // parent sku / item sku (availability depends on shop settings)
            [JsonPropertyName("item_sku")] public string? ItemSku { get; set; }

            // variant sku / model sku (availability depends on shop settings)
            [JsonPropertyName("model_sku")] public string? ModelSku { get; set; }

            // some sellers store variant code in this field (fallback)
            [JsonPropertyName("variation_sku")] public string? VariationSku { get; set; }

            [JsonPropertyName("model_name")] public string? ModelName { get; set; }
            [JsonPropertyName("variation_name")] public string? VariationName { get; set; }

            [JsonPropertyName("model_quantity_purchased")] public int ModelQuantity { get; set; } = 1;
        }

        private sealed class ShopeeSyncWindow : Window
        {
            private readonly TextBox _partnerId = new();
            private readonly PasswordBox _partnerKey = new();
            private readonly TextBox _shopId = new();
            private readonly PasswordBox _accessToken = new();
            private readonly TextBox _baseUrl = new();
            private readonly ComboBox _orderStatus = new();
            private readonly TextBox _timeFrom = new();
            private readonly TextBox _timeTo = new();

            public ShopeeConfig Config { get; private set; }

            public ShopeeSyncWindow(ShopeeConfig cfg)
            {
                Config = cfg;

                Title = "Shopee Sync (OpenAPI v2)";
                Width = 560;
                Height = 420;
                WindowStartupLocation = WindowStartupLocation.CenterOwner;
                ResizeMode = ResizeMode.NoResize;

                var root = new Grid { Margin = new Thickness(14) };
                root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                Content = root;

                var hint = new TextBlock
                {
                    Text = "Masukkan credential Shopee Open Platform (v2). Data disimpan di ./config/shopee.json",
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 10)
                };
                root.Children.Add(hint);

                var form = new Grid { Margin = new Thickness(0, 0, 0, 10) };
                Grid.SetRow(form, 1);
                root.Children.Add(form);

                form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) });
                form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                for (int i = 0; i < 8; i++)
                    form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                int r = 0;
                AddRow(form, r++, "Base URL", _baseUrl);
                AddRow(form, r++, "Partner ID", _partnerId);
                AddRow(form, r++, "Partner Key", _partnerKey);
                AddRow(form, r++, "Shop ID", _shopId);
                AddRow(form, r++, "Access Token", _accessToken);

                _orderStatus.ItemsSource = new[] { "READY_TO_SHIP", "PROCESSED", "SHIPPED", "COMPLETED", "CANCELLED", "IN_CANCEL" };
                _orderStatus.SelectedIndex = 0;
                AddRow(form, r++, "Order Status", _orderStatus);

                AddRow(form, r++, "Time From (unix)", _timeFrom);
                AddRow(form, r++, "Time To (unix)", _timeTo);

                // Defaults
                _baseUrl.Text = cfg.BaseUrl ?? DefaultShopeeBaseUrl;
                _partnerId.Text = cfg.PartnerId > 0 ? cfg.PartnerId.ToString(CultureInfo.InvariantCulture) : "";
                _partnerKey.Password = cfg.PartnerKey ?? "";
                _shopId.Text = cfg.ShopId > 0 ? cfg.ShopId.ToString(CultureInfo.InvariantCulture) : "";
                _accessToken.Password = cfg.AccessToken ?? "";
                _orderStatus.SelectedItem = string.IsNullOrWhiteSpace(cfg.OrderStatus) ? "READY_TO_SHIP" : cfg.OrderStatus;

                _timeFrom.Text = (cfg.TimeFromUnix > 0 ? cfg.TimeFromUnix : DateTimeOffset.UtcNow.AddDays(-2).ToUnixTimeSeconds()).ToString(CultureInfo.InvariantCulture);
                _timeTo.Text = (cfg.TimeToUnix > 0 ? cfg.TimeToUnix : DateTimeOffset.UtcNow.ToUnixTimeSeconds()).ToString(CultureInfo.InvariantCulture);

                var buttons = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Right
                };
                Grid.SetRow(buttons, 2);
                root.Children.Add(buttons);

                var btnCancel = new Button { Content = "Cancel", MinWidth = 90, Margin = new Thickness(0, 0, 8, 0) };
                btnCancel.Click += (_, __) => { DialogResult = false; Close(); };
                buttons.Children.Add(btnCancel);

                var btnOk = new Button { Content = "Save & Sync", MinWidth = 120 };
                btnOk.Click += (_, __) =>
                {
                    if (!long.TryParse(_partnerId.Text.Trim(), out var pid) || pid <= 0)
                    {
                        MessageBox.Show("Partner ID tidak valid.");
                        return;
                    }
                    if (!long.TryParse(_shopId.Text.Trim(), out var sid) || sid <= 0)
                    {
                        MessageBox.Show("Shop ID tidak valid.");
                        return;
                    }
                    if (!long.TryParse(_timeFrom.Text.Trim(), out var tf) || tf <= 0 ||
                        !long.TryParse(_timeTo.Text.Trim(), out var tt) || tt <= 0 || tt < tf)
                    {
                        MessageBox.Show("Time range unix tidak valid.");
                        return;
                    }
                    if (string.IsNullOrWhiteSpace(_partnerKey.Password) || string.IsNullOrWhiteSpace(_accessToken.Password))
                    {
                        MessageBox.Show("Partner Key & Access Token wajib diisi.");
                        return;
                    }

                    Config = new ShopeeConfig
                    {
                        BaseUrl = string.IsNullOrWhiteSpace(_baseUrl.Text) ? DefaultShopeeBaseUrl : _baseUrl.Text.Trim(),
                        PartnerId = pid,
                        PartnerKey = _partnerKey.Password,
                        ShopId = sid,
                        AccessToken = _accessToken.Password,
                        OrderStatus = (_orderStatus.SelectedItem?.ToString() ?? "READY_TO_SHIP"),
                        TimeFromUnix = tf,
                        TimeToUnix = tt
                    };

                    DialogResult = true;
                    Close();
                };
                buttons.Children.Add(btnOk);
            }

            private static void AddRow(Grid g, int row, string label, Control input)
            {
                var lbl = new TextBlock { Text = label, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 8) };
                Grid.SetRow(lbl, row);
                Grid.SetColumn(lbl, 0);
                g.Children.Add(lbl);

                input.Margin = new Thickness(0, 0, 0, 8);
                input.MinHeight = 26;
                Grid.SetRow(input, row);
                Grid.SetColumn(input, 1);
                g.Children.Add(input);
            }
        }
private void ShowPreviewLoading(string fullPath)
        {
            PreviewLoadingText.Text = $"Loading:\n{fullPath}";
            PreviewLoadingOverlay.Visibility = Visibility.Visible;
        }

        private void HidePreviewLoading()
        {
            PreviewLoadingOverlay.Visibility = Visibility.Collapsed;
            PreviewLoadingText.Text = "";
        }


        private string ConfigDir => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config");
        private string DefaultDataMapPath => Path.Combine(ConfigDir, "PaperbellDataMap.xlsx");

                // Composite key helpers (keep backward compatibility)


// ✅ New key for Shopee mapping (request): model_sku + item_sku (no separator)
private static string NormKey(string? s)
{
    if (string.IsNullOrWhiteSpace(s)) return "";
    return s.Trim().Replace(" ", "").ToLowerInvariant();
}

private static string KeyModelItem(string? modelSku, string? itemSku)
    => NormKey(modelSku) + NormKey(itemSku);
        private static string NormKeyPart(string? s) => (s ?? "").Trim();

        private static string KeySkuIndukRef(string skuInduk, string noRef)
            => $"IR|{NormKeyPart(skuInduk)}|{NormKeyPart(noRef)}";

        private static string KeySkuIndukOnly(string skuInduk)
            => $"I|{NormKeyPart(skuInduk)}";

        private static string KeyRefVar(string noRef, string variasi)
            => $"RV|{NormKeyPart(noRef)}|{NormKeyPart(variasi)}";

        private static string KeyRefOnly(string noRef)
            => $"R|{NormKeyPart(noRef)}";
public MainWindow()
        {
            InitializeComponent();
            InitPdfPreviewAvailability();

            if (!_pdfPreviewAvailable)
            {
                Thumbs.ItemsSource = null;
                HidePreviewLoading();
            }

            // === FIX koneksi Shopee (pakai system proxy & credential) ===
            System.Net.ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

            var handler = new SocketsHttpHandler
            {
                UseProxy = true,
                Proxy = WebRequest.DefaultWebProxy,
                Credentials = CredentialCache.DefaultCredentials,
                ConnectTimeout = TimeSpan.FromSeconds(20)
            };

            _shopeeHttp = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(60)
            };

            AllRowsView = CollectionViewSource.GetDefaultView(Rows);

            NotPrintedView = new ListCollectionView(Rows);
            NotPrintedView.Filter = o => o is JobRow r && !r.IsPrinted;

            PrintedView = new ListCollectionView(Rows);
            PrintedView.Filter = o => o is JobRow r && r.IsPrinted;



            QueueGrid.ItemsSource = Rows;
            DataContext = this;
            RefreshPrinters();
            
            EnsureDbDirectory();
            SQLitePCL.Batteries.Init();
            InitDatabase();
            LoadAppStateFromDb();
            UpdateShopeeUi();
            // ✅ Ensure DataGrid edits are committed immediately (prevents printing old printer value)
            QueueGrid.CellEditEnding += (s3, e3) =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        QueueGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
                        QueueGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);
                    }
                    catch { }
                }));
            };

            QueueGrid.CurrentCellChanged += (s4, e4) =>
            {
                try
                {
                    QueueGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
                    QueueGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);
                }
                catch { }
            };

            _currentTabFilter = ShopeeTabFilter.NotPrinted;
            _shopeePageIndex = 0;
            _pageIndexNotPrinted = 0;
            
            // ✅ Autoload DataMap after window shown (avoid app exit if error)
            this.Loaded += (s, e) => TryAutoLoadDataMap();

            // ✅ Prevent silent crash: show any unhandled UI exceptions
            Application.Current.DispatcherUnhandledException += (s2, e2) =>
            {
                MessageBox.Show(
                    "Unhandled error:" + e2.Exception,
                    "Application Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                e2.Handled = true;
            };
        }

        private void RefreshViews()
        {
            AllRowsView?.Refresh();
            NotPrintedView?.Refresh();
            PrintedView?.Refresh();
        }

        // =====================
        // ✅ ADD: AutoLoad DataMap
        // =====================
        private void TryAutoLoadDataMap()
        {
            try
            {
                Directory.CreateDirectory(ConfigDir);

                if (!File.Exists(DefaultDataMapPath))
                {
                    // DataMap belum ada; biarkan user load manual lewat tombol Load Map.
                    return;
                }

                // penting untuk ExcelDataReader (xls/xlsx lama)
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                LoadDataMap(DefaultDataMapPath);
                LoadShopeePageFromDb(0);

                // optional: info sukses
                // MessageBox.Show("DataMap loaded:\n" + DefaultDataMapPath);
                BtnLoadMap.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Gagal autoload DataMap.\n\n" +
                    "Path:\n" + DefaultDataMapPath + "\n\n" +
                    "Error:\n" + ex,
                    "AutoLoad DataMap Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        // ✅ OPTIONAL column finder (tidak throw)
        private static string? FindColOptional(DataTable t, params string[] candidates)
        {
            foreach (var c in candidates)
                if (t.Columns.Contains(c)) return c;

            string Norm(string s) => (s ?? "").ToLowerInvariant()
                .Replace(" ", "").Replace(".", "").Replace("_", "").Replace("-", "");

            var cols = t.Columns.Cast<DataColumn>().Select(dc => dc.ColumnName).ToList();
            var normCols = cols.ToDictionary(x => Norm(x), x => x);

            foreach (var cand in candidates)
            {
                var nc = Norm(cand);
                if (normCols.TryGetValue(nc, out var real)) return real;
            }

            return null;
        }

        // ✅ parse kolom "Page" seperti: "1", "1-2", "2-", "-5"
        private static (int from, int to) ParsePageRange(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return (1, 1);
            s = s.Trim();

            // ✅ Banyak file mapper memasukkan page range dalam format "3/4" yang oleh Excel dianggap DATE.
            // Contoh: 2025-03-04 00:00:00  -> artinya halaman 3-4 (Month-Day).
            if (DateTime.TryParse(s, out var dt))
            {
                int a = dt.Month;
                int b = dt.Day;
                if (a >= 1 && a <= 500 && b >= 1 && b <= 500)
                {
                    if (b < a) (a, b) = (b, a);
                    return (a, b);
                }
            }

            // ✅ parse kolom "Page" seperti: "1", "1-2", "2-", "-5"
            var m = Regex.Match(s, @"^\s*(\d+)?\s*(?:-\s*(\d+)?)?\s*$");
            if (!m.Success) return (1, 1);

            int? aa1 = null, bb1 = null;
            if (m.Groups[1].Success && int.TryParse(m.Groups[1].Value, out var aa)) aa1 = aa;
            if (m.Groups[2].Success && int.TryParse(m.Groups[2].Value, out var bb)) bb1 = bb;

            int from = Math.Max(1, aa1 ?? 1);

            // kalau "1" (tidak ada '-')
            if (!s.Contains("-"))
            {
                // guard: angka kebesaran (misal Excel serial date kebaca sebagai 45234)
                if (from > 500) return (1, 1);
                return (from, from);
            }

            // "2-" => sampai akhir (to=0)
            if (aa1.HasValue && !bb1.HasValue) return (from, 0);

            // "-5" => dari 1 sampai 5
            if (!aa1.HasValue && bb1.HasValue) return (1, Math.Max(1, Math.Min(500, bb1.Value)));

            // "1-2"
            int to = Math.Max(0, bb1 ?? from);
            to = Math.Min(500, to);

            if (to > 0 && to < from) (from, to) = (to, from);
            return (from, to);
        }



        // =====================
        // UI handlers
        // =====================

        private void RefreshPrinters_Click(object sender, RoutedEventArgs e) => RefreshPrinters();

        private void PickPdfs_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Multiselect = true,
                Title = "Select PDF files"
            };

            if (dlg.ShowDialog() == true)
            {
                foreach (var p in dlg.FileNames)
                {
                    Rows.Add(new JobRow
                    {
                        Index = Rows.Count + 1,
                        File = p,
                        Printer = Printers.FirstOrDefault() ?? "",
                        PageFrom = 1,
                        PageTo = 1,
                        Copies = 1,
                        Duplex = DuplexMode.Simplex,
                        Paper = PaperPreset.A5,
                        Pages = "1",
                        Status = "Ready",
                        Percent = 0,
                        TotalPages = 0
                    });
                }
            }
        }

        public bool IsPrinted { get; set; }

        private async void StartPreviewLatest(JobRow row)
        {
            if (!_pdfPreviewAvailable)
            {
                row.Status = "Preview unavailable";
                row.Percent = 0;
                Thumbs.ItemsSource = null;
                HidePreviewLoading();
                return;
            }

            var oldRow = _previewRow;

            try
            {
                _previewCts?.Cancel();
                _previewCts?.Dispose();
            }
            catch { }

            _previewCts = new CancellationTokenSource();
            var ct = _previewCts.Token;

            int reqId = ++_previewReqId;
            _previewRow = row;

            ShowPreviewLoading(row.File);

            if (oldRow != null && !ReferenceEquals(oldRow, row))
            {
                if (oldRow.Status == "Loading preview…")
                    oldRow.Status = "Ready";
            }

            try
            {
                await LoadPdfPreviewAsync(row, ct, reqId);
            }
            catch (Exception ex)
            {
                DisablePdfPreview(row, ex);
            }
        }

        private void QueueGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (QueueGrid.SelectedItem is JobRow row)
            {
                _lastUserPickedForPreview = row;
                row.Status = "Ready";

                if (!_suppressPreviewWhileScrolling)
                {
                    _queueGridPreviewDeferredWhileScrolling = false;
                    StartPreviewLatest(row);
                }
                else
                    _queueGridPreviewDeferredWhileScrolling = true;
            }
        }

        // =====================
        // PDF Preview (PdfiumViewer)
        // =====================

        private async Task LoadPdfPreviewAsync(JobRow row, CancellationToken ct, int reqId)
        {
            Thumbs.ItemsSource = null;

            if (!_pdfPreviewAvailable)
            {
                row.Status = "Ready (preview disabled)";
                HidePreviewLoading();
                return;
            }

            if (string.IsNullOrWhiteSpace(row.File) || !File.Exists(row.File))
            {
                row.Status = "No file";
                HidePreviewLoading();
                return;
            }

            row.Status = "Loading preview…";
            row.Percent = 0;

            try
            {
                using var doc = await Task.Run(() =>
                {
                    ct.ThrowIfCancellationRequested();
                    return PdfDocument.Load(row.File);
                }, ct);

                ct.ThrowIfCancellationRequested();

                if (reqId != _previewReqId)
                    return;

                row.TotalPages = doc.PageCount;

                var thumbs = new ObservableCollection<PageThumb>();
                const int thumbTargetWidth = 600;
                const int thumbDpi = 120;

                for (int i = 0; i < doc.PageCount; i++)
                {
                    ct.ThrowIfCancellationRequested();

                    if (reqId != _previewReqId)
                        return;

                    int pageNo = i + 1;

                    var img = await Task.Run(() =>
                    {
                        ct.ThrowIfCancellationRequested();
                        using var bmp = RenderPdfPageSafe(doc, i, thumbTargetWidth, thumbDpi);
                        ct.ThrowIfCancellationRequested();
                        return BitmapToBitmapImage(bmp);
                    }, ct);

                    ct.ThrowIfCancellationRequested();

                    if (reqId != _previewReqId)
                        return;

                    thumbs.Add(new PageThumb
                    {
                        File = row.File,
                        PageNumber = pageNo,
                        TotalPages = doc.PageCount,
                        Image = img,
                        Caption = $"Page {pageNo}"
                    });
                }

                if (reqId != _previewReqId)
                    return;

                Thumbs.ItemsSource = thumbs;
                row.Status = "Ready";
                row.Percent = 0;
                HidePreviewLoading();
            }
            catch (OperationCanceledException)
            {
                if (row.Status == "Loading preview…")
                    row.Status = "Ready";

                HidePreviewLoading();
            }
            catch (Exception ex)
            {
                if (IsPdfiumFatal(ex))
                {
                    DisablePdfPreview(row, ex);
                }
                else
                {
                    row.Status = "Preview error: " + ex.Message;
                    row.Percent = 0;
                    HidePreviewLoading();
                }
            }
            finally
            {
                if (reqId != _previewReqId && row.Status == "Loading preview…")
                    row.Status = "Ready";
            }
        }


        private static Bitmap RenderPdfPageSafe(PdfDocument doc, int pageIndex, int targetWidthPx, int dpi)
        {
            try
            {
                var size = doc.PageSizes[pageIndex]; // points (1/72 inch)

                double widthPx = size.Width / 72.0 * dpi;
                double heightPx = size.Height / 72.0 * dpi;

                double scale = targetWidthPx / widthPx;
                int w = Math.Max(1, (int)Math.Round(widthPx * scale));
                int h = Math.Max(1, (int)Math.Round(heightPx * scale));

                using var img = doc.Render(
                    pageIndex,
                    w,
                    h,
                    dpi,
                    dpi,
                    PdfRenderFlags.Annotations
                );

                return new Bitmap(img);
            }
            catch
            {
                throw;
            }
        }

        private static BitmapImage BitmapToBitmapImage(Bitmap bmp)
        {
            using var ms = new MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            ms.Position = 0;

            var bi = new BitmapImage();
            bi.BeginInit();
            bi.CacheOption = BitmapCacheOption.OnLoad;
            bi.StreamSource = ms;
            bi.EndInit();
            bi.Freeze();
            return bi;
        }

        // Click thumbnail => open modal full size
        private async void Thumb_Click(object sender, MouseButtonEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is not PageThumb pt)
                return;

            if (!_pdfPreviewAvailable)
                return;

            _zoom = 1.0;
            ZoomText.Text = "100%";
            _full = (pt.File, pt.PageNumber, pt.TotalPages);

            try
            {
                FullImage.Source = await RenderFullAsync(pt.File, pt.PageNumber, baseWidthPx: 1400, zoom: _zoom);
                if (FullImage.Source != null)
                    Modal.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                if (IsPdfiumFatal(ex))
                    DisablePdfPreview(null, ex);
                else
                    LogPrint("Failed to render full page: " + ex);
            }
        }

        private async Task<BitmapImage?> RenderFullAsync(string path, int pageNumber, int baseWidthPx, double zoom)
        {
            if (!_pdfPreviewAvailable)
                return null;

            try
            {
                return await Task.Run(() =>
                {
                    using var doc = PdfDocument.Load(path);
                    int pageIndex = pageNumber - 1;
                    int targetWidth = (int)Math.Round(baseWidthPx * zoom);
                    using var bmp = RenderPdfPageSafe(doc, pageIndex, targetWidth, dpi: 150);
                    return BitmapToBitmapImage(bmp);
                });
            }
            catch (Exception ex)
            {
                if (IsPdfiumFatal(ex))
                {
                    Dispatcher.Invoke(() => DisablePdfPreview(null, ex));
                    return null;
                }

                throw;
            }
        }

        private async void ZoomIn_Click(object sender, RoutedEventArgs e)
        {
            if (_full == null || !_pdfPreviewAvailable) return;

            try
            {
                _zoom = Math.Min(3.0, _zoom + 0.25);
                ZoomText.Text = $"{(int)(_zoom * 100)}%";
                FullImage.Source = await RenderFullAsync(_full.Value.file, _full.Value.page, 1400, _zoom);
            }
            catch (Exception ex)
            {
                if (IsPdfiumFatal(ex))
                    DisablePdfPreview(null, ex);
                else
                    LogPrint("Zoom error:\n" + ex.Message);
            }
        }

        private async void ZoomOut_Click(object sender, RoutedEventArgs e)
        {
            if (_full == null || !_pdfPreviewAvailable) return;

            try
            {
                _zoom = Math.Max(0.5, _zoom - 0.25);
                ZoomText.Text = $"{(int)(_zoom * 100)}%";
                FullImage.Source = await RenderFullAsync(_full.Value.file, _full.Value.page, 1400, _zoom);
            }
            catch (Exception ex)
            {
                if (IsPdfiumFatal(ex))
                    DisablePdfPreview(null, ex);
                else
                    LogPrint("Zoom error:\n" + ex.Message);
            }
        }

        private void CloseModal_Click(object sender, RoutedEventArgs e) => Modal.Visibility = Visibility.Collapsed;

        // =====================
        // Printing (SumatraPDF)
        // =====================

        private void PrintRow_Click(object sender, RoutedEventArgs e)
        {
            // ✅ commit any in-cell edits (ComboBox/TextBox) before reading row values
            QueueGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
            QueueGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);
            Keyboard.ClearFocus();

            if ((sender as FrameworkElement)?.DataContext is JobRow r)
                _ = PrintAsync(r);
        }

        private void CancelRow_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is JobRow r)
            {
                if (r.JobGuid != null && _jobCts.TryGetValue(r.JobGuid.Value, out var cts))
                {
                    cts.Cancel();
                    _jobCts.Remove(r.JobGuid.Value);
                }
                r.Status = "Canceled";
                r.HasJob = false;
                r.Percent = 0;
            }
        }

        private void RemoveRow_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is JobRow r)
                Rows.Remove(r);
        }

        private async void PrintSelected_Click(object sender, RoutedEventArgs e)
        {
            QueueGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            QueueGrid.CommitEdit(DataGridEditingUnit.Row, true);
            Keyboard.ClearFocus();

            foreach (var r in Rows.ToList())
                await PrintAsync(r);
        }

        private static string BuildSumatraPrintSettings(JobRow r)
        {
            // Range
            int from = Math.Max(1, r.PageFrom);
            int to = r.PageTo <= 0 ? 0 : Math.Max(from, r.PageTo);
            string range = to <= 0 ? $"{from}-" : (from == to ? $"{from}" : $"{from}-{to}");

            // Duplex
            string duplex = r.Duplex switch
            {
                DuplexMode.DuplexLongEdge => "duplexlong",
                DuplexMode.DuplexShortEdge => "duplexshort",
                _ => "simplex"
            };

            // Default scaling: no scale (sesuai request)
            const string scaling = "noscale";

            // Paper: only include for standard sizes. For true custom like B5 JIS 192x257mm, use a dedicated printer profile in Windows.
            string paper = r.Paper switch
            {
                PaperPreset.A5 => "paper=A5",
                PaperPreset.A6 => "paper=A6",
                PaperPreset.B5JIS => "paper=B5",
                _ => ""
            };

            // Copies (Sumatra supports "3x")
            int copies = Math.Max(1, r.Copies);
            string copiesPart = copies > 1 ? $"{copies}x" : "";

            var parts = new List<string> { range, duplex, scaling };
            if (r.PrintMonochrome)
                parts.Add("monochrome");
            if (!string.IsNullOrWhiteSpace(paper)) parts.Add(paper);
            if (!string.IsNullOrWhiteSpace(copiesPart)) parts.Add(copiesPart);

            return string.Join(",", parts);
        }

        private static void TryApplyJobRowPaperAndColor(PrintDocument pd, JobRow r)
        {
            pd.DefaultPageSettings.Landscape = false;
            if (r.PrintMonochrome)
                pd.DefaultPageSettings.Color = false;

            switch (r.Paper)
            {
                case PaperPreset.A6:
                    TrySetPrinterPaperSize(pd, PaperKind.A6, 413, 583); // ~105×148 mm
                    break;
                case PaperPreset.A5:
                    TrySetPrinterPaperSize(pd, PaperKind.A5, 583, 827); // ~148×210 mm
                    break;
            }
        }

        private static void TrySetPrinterPaperSize(PrintDocument pd, PaperKind kind, int widthHundredthsInch, int heightHundredthsInch)
        {
            foreach (PaperSize ps in pd.PrinterSettings.PaperSizes)
            {
                if (ps.Kind == kind)
                {
                    pd.DefaultPageSettings.PaperSize = ps;
                    return;
                }
            }

            pd.DefaultPageSettings.PaperSize = new PaperSize(kind.ToString(), widthHundredthsInch, heightHundredthsInch);
        }

        /// <summary>Cetak PDF lewat Pdfium + GDI+: isi di area margin dengan faktor <paramref name="contentScale"/>; jika <see cref="JobRow.PrintFromTopLeft"/> kiri-atas, selain itu tengah.</summary>
        private static void PrintPdfToPrinterWithScale(JobRow r, string printerName, double contentScale)
        {
            contentScale = Math.Clamp(contentScale, 0.05, 3.0);

            using var doc = PdfDocument.Load(r.File);
            int pdfTotal = doc.PageCount;
            int from = Math.Max(1, r.PageFrom);
            int to = r.PageTo <= 0 ? pdfTotal : Math.Min(Math.Max(from, r.PageTo), pdfTotal);

            using var pd = new PrintDocument();
            pd.PrinterSettings.PrinterName = printerName;
            pd.PrinterSettings.Copies = (short)Math.Min(999, Math.Max(1, r.Copies));
            pd.PrinterSettings.Duplex = r.Duplex switch
            {
                DuplexMode.DuplexLongEdge => Duplex.Vertical,
                DuplexMode.DuplexShortEdge => Duplex.Horizontal,
                _ => Duplex.Simplex
            };
            TryApplyJobRowPaperAndColor(pd, r);
            // Kurangi margin lunak Windows supaya (0,0) mendekati tepi kiri atas kertas/driver.
            pd.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);

            var current = from;
            pd.PrintPage += (_, e) =>
            {
                int pageIndex = current - 1;
                if (pageIndex < 0 || pageIndex >= doc.PageCount)
                {
                    e.HasMorePages = false;
                    return;
                }

                var size = doc.PageSizes[pageIndex];
                float printableWInch = e.MarginBounds.Width / 100f;
                float printableHInch = e.MarginBounds.Height / 100f;

                float boxWInch = printableWInch * (float)contentScale;
                float boxHInch = printableHInch * (float)contentScale;

                double pageWIn = size.Width / 72.0;
                double pageHIn = size.Height / 72.0;
                if (pageWIn <= 0 || pageHIn <= 0)
                {
                    e.HasMorePages = false;
                    return;
                }

                double ratioPage = pageWIn / pageHIn;
                double ratioBox = boxWInch / boxHInch;
                double drawWIn = ratioPage > ratioBox ? boxWInch : boxHInch * ratioPage;
                double drawHIn = ratioPage > ratioBox ? boxWInch / ratioPage : boxHInch;

                const int dpi = 300;
                int pixW = Math.Max(1, (int)Math.Round(drawWIn * dpi));
                using var bmp = RenderPdfPageSafe(doc, pageIndex, pixW, dpi);

                float drawWUnits = (float)(drawWIn * 100);
                float drawHUnits = (float)(drawHIn * 100);
                // Origin Graphics pada cetak = sudah relatif ke pojok kiri-atas area cetak (margin).
                // Jangan tambahkan MarginBounds.Left/Top lagi — itu membuat gambar tengah/tergeser.
                float x = r.PrintFromTopLeft
                    ? 0f
                    : (e.MarginBounds.Width - drawWUnits) / 2f;
                float y = r.PrintFromTopLeft
                    ? 0f
                    : (e.MarginBounds.Height - drawHUnits) / 2f;

                var g = e.Graphics!;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.DrawImage(bmp, x, y, drawWUnits, drawHUnits);

                e.HasMorePages = current < to;
                current++;
            };

            pd.Print();
        }

        // ✅ Resolve the picked printer name to an actually installed printer name.
        // If Sumatra can't find the printer, it will fall back to Windows default (often Epson).
        // So we MUST validate/resolve, otherwise user selection won't be respected.
        private static string? ResolvePrinterName(string? picked)
        {
            var p = (picked ?? "").Trim();
            if (string.IsNullOrWhiteSpace(p)) return null;

            var installed = System.Drawing.Printing.PrinterSettings.InstalledPrinters
                .Cast<string>()
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            // Exact match
            var exact = installed.FirstOrDefault(x => string.Equals(x, p, StringComparison.OrdinalIgnoreCase));
            if (exact != null) return exact;

            // Try "contains" match (handles cases like "Brother HL-L..." vs "Brother HL-L... (Copy 1)")
            var contains = installed.FirstOrDefault(x => x.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
                        ?? installed.FirstOrDefault(x => p.IndexOf(x, StringComparison.OrdinalIgnoreCase) >= 0);

            return contains;
        }

        private void MarkPrintedInUiAndDb(JobRow r, bool printed)
{
    if (r == null) return;

    // Only persist for Shopee rows saved in DB.
    if (r.OrderProcessId <= 0)
    {
        r.IsPrinted = printed;
        return;
    }

    r.IsPrinted = printed;
    DbSetPrintedById(r.OrderProcessId, printed);
}



        private async Task PrintAsync(JobRow r, bool markShopeeLineAsPrinted = true)
        {
            if (string.IsNullOrWhiteSpace(r.File) || !File.Exists(r.File))
            {
                r.Status = "File missing";
                return;
            }

            if (string.IsNullOrWhiteSpace(r.Printer))
            {
                r.Status = "Select a printer";
                return;
            }

            NormalizePageRangeForPrint(r);

            // ✅ Validate/resolve printer name so Sumatra doesn't fall back to Windows default printer
            var resolvedPrinter = ResolvePrinterName(r.Printer);
            if (string.IsNullOrWhiteSpace(resolvedPrinter))
            {
                r.Status = "Selected printer not found on this PC";
                r.Percent = 0;
                return;
            }
            var pname = resolvedPrinter;

            // Printer rules
            var isEpsonL3210 = pname.IndexOf("L3210", StringComparison.OrdinalIgnoreCase) >= 0;
            // For Epson L3210: simplex only (we’re not setting duplex anyway in Sumatra path)

            // Ensure TotalPages known (needed for parser + progress)
            if (r.TotalPages <= 0 && _pdfPreviewAvailable)
            {
                try
                {
                    using var doc = PdfDocument.Load(r.File);
                    r.TotalPages = doc.PageCount;
                }
                catch (Exception ex)
                {
                    LogPrint("Failed to read PDF page count: " + ex);
                }
            }

            // Compute how many pages will actually be printed for UI progress (e.g. 1-1 should show 1/1, not 1/8)
            int pdfTotal = r.TotalPages > 0 ? r.TotalPages : 0;
            int from = Math.Max(1, r.PageFrom);
            int to = r.PageTo <= 0 ? 0 : Math.Max(from, r.PageTo);

            int effectiveTo = (to <= 0 || pdfTotal <= 0) ? to : Math.Min(to, pdfTotal);

            int printTotalPages;
            if (effectiveTo <= 0)
            {
                // "to end" — if pdfTotal known use it, otherwise fallback to 1 for now
                printTotalPages = pdfTotal > 0 ? Math.Max(1, pdfTotal - from + 1) : 1;
            }
            else
            {
                printTotalPages = Math.Max(1, effectiveTo - from + 1);
            }

            r.PrintTotalPages = printTotalPages;
            r.PrintedPages = 0;

            if (markShopeeLineAsPrinted)
                MarkPrintedInUiAndDb(r, true);

            var usePdfiumGdi = _pdfPreviewAvailable &&
                               (Math.Abs(r.PdfPrintScale - 1.0) > 0.0001 || r.PrintFromTopLeft);
            if (!usePdfiumGdi && Math.Abs(r.PdfPrintScale - 1.0) > 0.0001)
                LogPrint($"PdfPrintScale={r.PdfPrintScale} diminta tapi PDF preview (Pdfium) mati — fallback Sumatra tanpa skala persen.");
            if (!usePdfiumGdi && r.PrintFromTopLeft)
                LogPrint("Cetak label kiri-atas (Pdfium+GDI) tidak tersedia — fallback Sumatra.");

            if (usePdfiumGdi)
            {
                r.Status = $"On progress (0/{r.PrintTotalPages} page)";
                r.Percent = 5;
                r.HasJob = true;
                r.JobGuid = Guid.NewGuid();

                var cts = new CancellationTokenSource();
                _jobCts[r.JobGuid.Value] = cts;

                try
                {
                    _ = MonitorPrintJobAsync(r, cts.Token);
                    LogPrint($"PDFIUM GDI print scale={r.PdfPrintScale} topLeft={r.PrintFromTopLeft} printer='{pname}' file='{r.File}'");

                    await Task.Run(() => PrintPdfToPrinterWithScale(r, pname, r.PdfPrintScale), cts.Token);

                    if (!cts.IsCancellationRequested)
                    {
                        r.Percent = Math.Max(r.Percent, 90);
                        r.Status = $"On progress ({r.PrintTotalPages} page)";
                        await Task.Delay(400);
                        r.Percent = 100;
                        r.Status = "Done";
                    }
                }
                catch (OperationCanceledException)
                {
                    r.Status = "Canceled";
                    r.Percent = 0;
                }
                catch (Exception ex)
                {
                    r.Status = "Print error: " + ex.Message;
                    r.Percent = 0;
                }
                finally
                {
                    r.HasJob = false;
                    if (r.JobGuid != null) _jobCts.Remove(r.JobGuid.Value);
                    r.JobGuid = null;
                }

                return;
            }

            var sumatra = TryFindSumatra();
            if (sumatra == null)
            {
                r.Status = "SumatraPDF not found. Install SumatraPDF to print.";
                r.Percent = 0;
                return;
            }

            // New print job state
            r.Status = $"On progress (0/{r.PrintTotalPages} page)";
            r.Percent = 5;
            r.HasJob = true;
            r.JobGuid = Guid.NewGuid();

            var cts2 = new CancellationTokenSource();
            _jobCts[r.JobGuid.Value] = cts2;

            try
            {
                // Kick off monitoring (best effort)
                _ = MonitorPrintJobAsync(r, cts2.Token);

                // Print via SumatraPDF
                await Task.Run(() =>
                {
                    var printSettings = BuildSumatraPrintSettings(r);
                    var args = $"-print-to \"{pname}\" -print-settings \"{printSettings}\" -silent -exit-on-print \"{r.File}\"";

                    LogPrint($"ROW OrderNo='{r.OrderNo}' File='{r.File}' PickedPrinter='{r.Printer}' ResolvedPrinter='{pname}' Duplex='{r.Duplex}' Paper='{r.Paper}' Pages='{r.PageRange}' Copies='{r.Copies}'");
                    LogPrint("SUMATRA: " + sumatra);
                    LogPrint("ARGS: " + args);

                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = sumatra,
                        Arguments = args,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    using var p = System.Diagnostics.Process.Start(psi);
                    if (p == null)
                    {
                        LogPrint("ERROR: Failed to start Sumatra process.");
                        return;
                    }

                    // Wait a bit so we can capture errors; Sumatra will submit job quickly.
                    if (!p.WaitForExit(8000))
                    {
                        LogPrint("WARN: Sumatra did not exit within 8s (job may still be queued).");
                    }

                    try
                    {
                        var so = p.StandardOutput.ReadToEnd();
                        var se = p.StandardError.ReadToEnd();
                        if (!string.IsNullOrWhiteSpace(so)) LogPrint("STDOUT: " + so.Trim());
                        if (!string.IsNullOrWhiteSpace(se)) LogPrint("STDERR: " + se.Trim());
                        LogPrint("EXIT: " + p.ExitCode);
                    }
                    catch (Exception exRead)
                    {
                        LogPrint("WARN: Failed reading stdout/stderr: " + exRead.Message);
                    }
                }, cts2.Token);
// If monitoring didn’t manage to find job, still finish nicely
                if (!cts2.IsCancellationRequested)
                {
                    r.Percent = Math.Max(r.Percent, 90);
                    r.Status = $"On progress ({r.PrintTotalPages} page)";
                    await Task.Delay(400);
                    r.Percent = 100;
                    r.Status = "Done";
                }
            }
            catch (OperationCanceledException)
            {
                r.Status = "Canceled";
                r.Percent = 0;
            }
            catch (Exception ex)
            {
                r.Status = "Print error: " + ex.Message;
                r.Percent = 0;
            }
            finally
            {
                r.HasJob = false;
                if (r.JobGuid != null) _jobCts.Remove(r.JobGuid.Value);
                r.JobGuid = null;
            }
        }

        private static string? TryFindSumatra()
        {
            // Prefer local tools folder next to the running exe:
            // <output>\tools\SumatraPDF.exe
            var localTools = Path.Combine(AppContext.BaseDirectory, "tools", "SumatraPDF.exe");
            if (File.Exists(localTools)) return localTools;

            string[] candidates =
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "SumatraPDF", "SumatraPDF.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "SumatraPDF", "SumatraPDF.exe"),
            };

            return candidates.FirstOrDefault(File.Exists);
        }

        // =====================
        // Progress monitoring (Windows spooler, best-effort)
        // =====================

        private async Task MonitorPrintJobAsync(JobRow r, CancellationToken ct)
        {
            try
            {
                // Give Sumatra time to submit job
                await Task.Delay(350, ct);

                var server = new LocalPrintServer();
                var qName = ResolvePrinterName(r.Printer) ?? r.Printer;
                var queue = server.GetPrintQueue(qName);

                var fileName = Path.GetFileName(r.File);
                var user = WindowsIdentity.GetCurrent().Name;

                PrintSystemJobInfo? job = null;

                // Try to find matching job for up to ~15 seconds
                var start = DateTime.UtcNow;
                while (!ct.IsCancellationRequested && (DateTime.UtcNow - start) < TimeSpan.FromSeconds(15))
                {
                    queue.Refresh();
                    foreach (var j in queue.GetPrintJobInfoCollection())
                    {
                        // Name is often the document name; Sumatra usually uses the PDF filename.
                        if ((j.Name?.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0)
                        {
                            job = j;
                            break;
                        }
                    }

                    if (job != null) break;
                    await Task.Delay(300, ct);
                }

                if (job == null)
                    return; // Monitoring not available

                // Track until completion/deletion
                int total = r.PrintTotalPages > 0 ? r.PrintTotalPages : (r.TotalPages > 0 ? r.TotalPages : 0);

                while (!ct.IsCancellationRequested)
                {
                    queue.Refresh();
                    try { job.Refresh(); } catch { break; }

                    // Page-based progress if available
                    if (total > 0)
                    {
                        int printed = Math.Min(Math.Max(0, job.NumberOfPagesPrinted), total);
                        int pct = (int)Math.Floor(100.0 * Math.Min(printed, total) / total);
                        r.Percent = Math.Max(r.Percent, Math.Min(99, pct));
                        r.PrintedPages = printed;
                        r.Status = $"On progress ({r.PrintedPages}/{total} page)";
                    }

                    if (job.IsCompleted || job.IsDeleted || (job.JobStatus & PrintJobStatus.Completed) != 0)
                        break;

                    await Task.Delay(350, ct);
                }

                // Mark complete for UI
                if (!ct.IsCancellationRequested)
                {
                    r.PrintedPages = total > 0 ? total : r.PrintedPages;
                    r.Percent = 100;
                    r.Status = "Done";
                }
            }
            catch
            {
                // Ignore monitoring errors (it’s best-effort)
            }
        }

        // =====================
        // Printers list
        // =====================

        /// <summary>Set pilihan printer label: pertahankan <paramref name="previousSelection"/> jika masih ada, lalu default Epson L3210 (nama mengandung <see cref="ResiDefaultPrinterNameContains"/>), lalu printer pertama.</summary>
        private void SelectResiPrinterDefaultOrRestore(string? previousSelection)
        {
            if (CmbResiPrinter == null || Printers.Count == 0)
                return;
            try
            {
                if (!string.IsNullOrWhiteSpace(previousSelection))
                {
                    var prev = previousSelection.Trim();
                    var still = Printers.FirstOrDefault(p => string.Equals(p, prev, StringComparison.OrdinalIgnoreCase));
                    if (still != null)
                    {
                        CmbResiPrinter.SelectedItem = still;
                        return;
                    }
                }

                var l3210 = Printers.FirstOrDefault(p =>
                    p.IndexOf(ResiDefaultPrinterNameContains, StringComparison.OrdinalIgnoreCase) >= 0);
                CmbResiPrinter.SelectedItem = l3210 ?? Printers[0];
            }
            catch
            {
                // ignore
            }
        }

        private void RefreshPrinters()
        {
            var previous = (CmbResiPrinter?.SelectedItem as string)?.Trim();
            Printers.Clear();
            try
            {
                // ✅ Most compatible with SumatraPDF: use InstalledPrinters names
                foreach (string n in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    if (!string.IsNullOrWhiteSpace(n) && !Printers.Contains(n))
                        Printers.Add(n);
                }

                // ✅ Also include print queues (sometimes additional / shared queues appear here)
                var server = new LocalPrintServer();
                foreach (var q in server.GetPrintQueues().OrderBy(q => q.Name))
                {
                    var name = q.Name;
                    if (!string.IsNullOrWhiteSpace(name) && !Printers.Contains(name))
                        Printers.Add(name);
                }
            }
            catch
            {
                // fallback: InstalledPrinters only
                foreach (string n in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    if (!string.IsNullOrWhiteSpace(n) && !Printers.Contains(n))
                        Printers.Add(n);
                }
            }

            SelectResiPrinterDefaultOrRestore(previous);
        }

        // =====================================================================
        // ✅ ADD: Excel Import + Mapping (EXTENSION ONLY, TIDAK ganggu fitur lain)
        // =====================================================================

        /// <summary>
        /// OPTIONAL handler (kalau kamu bikin tombol di XAML):
        /// - pilih DataMap xlsx manual, lalu disalin ke .\config\PaperbellDataMap.xlsx dan di-load.
        /// </summary>
        private void LoadDataMapXlsx_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                Title = "Select PaperbellDataMap.xlsx"
            };

            if (dlg.ShowDialog() != true) return;

            try
            {
                Directory.CreateDirectory(ConfigDir);

                // copy ke config biar autoload next time
                File.Copy(dlg.FileName, DefaultDataMapPath, overwrite: true);

                LoadDataMap(DefaultDataMapPath);

                MessageBox.Show("DataMap loaded:\n" + DefaultDataMapPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Gagal load DataMap:\n" + ex.Message);
            }
        }

        /// <summary>
        /// OPTIONAL handler (kalau kamu bikin tombol di XAML):
        /// import Order.toship.xlsx lalu create rows berdasarkan mapping.
        /// </summary>
        private void ImportOrdersXlsx_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                Title = "Select Order.toship.xlsx"
            };

            if (dlg.ShowDialog() != true) return;

            try
            {
                ImportOrdersAndCreateRows(dlg.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Gagal import orders:\n" + ex.Message);
            }
        }

        private void LoadDataMap(string xlsxPath)
{
    _dataMap.Clear();
    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    using var stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    using var reader = ExcelReaderFactory.CreateReader(stream);

    var ds = reader.AsDataSet(new ExcelDataSetConfiguration
    {
        ConfigureDataTable = _ => new ExcelDataTableConfiguration
        {
            UseHeaderRow = true
        }
    });

    if (ds.Tables.Count == 0)
        throw new InvalidOperationException("DataMap XLSX tidak punya sheet.");

    var t = ds.Tables[0];

        // ✅ Wajib (No. Referensi SKU) + ✅ Optional (SKU Induk / Variasi)
    string colNoRef = FindCol(t, "Variasi", "variasi", "No.Referensi SKU", "Nomor Referensi SKU", "No Referensi SKU", "Nomor Referensi", "Product Code", "ProductCode", "SKU ID", "SKUID", "SKU");
    string? colSkuInduk = FindColOptional(t, "SKU Induk", "SKUInduk", "Induk SKU", "Parent SKU", "SKU Parent", "SKU Induk (Parent)");
    string? colVar = FindColOptional(t, "Nama Variasi", "Variasi", "Variant");

    // ✅ Optional (sesuai file kamu: Folder Path, Printer Name, Page, Copies, Duplex, Size)
    string? colPdf = FindColOptional(t, "FilePath", "File Path", "filepath", "Folder Path", "FolderPath", "PDF Path", "PdfPath", "Path", "File");
    string? colPrinter = FindColOptional(t, "Printer Name", "Printer", "Nama Printer");

    // 🔥 ini yang penting: file kamu punya "Page" bukan PageFrom/PageTo
    string? colPage = FindColOptional(t, "Page", "Pages", "Halaman");
    string? colPageFrom = FindColOptional(t, "PageFrom", "From", "Halaman Dari");
    string? colPageTo = FindColOptional(t, "PageTo", "To", "Halaman Sampai");

    string? colCopies = FindColOptional(t, "Copies", "Copy", "Jumlah Copy");
    string? colDuplex = FindColOptional(t, "Duplex", "DoubleSided");
    string? colPaper = FindColOptional(t, "Size", "Paper", "Kertas");

    foreach (DataRow r in t.Rows)
    {
        var noRef = (r[colNoRef]?.ToString() ?? "").Trim();
        var skuInduk = SafeGet(r, colSkuInduk ?? "").Trim();
        var varr = SafeGet(r, colVar ?? "").Trim();
int pageFrom = 1, pageTo = 1;

        // ✅ Prioritas: kolom "Page" (contoh: "1-2", "2-", dst)
        if (!string.IsNullOrWhiteSpace(colPage) && t.Columns.Contains(colPage))
        {
            (pageFrom, pageTo) = ParsePageRange(r[colPage]?.ToString());
        }
        else
        {
            // fallback kalau suatu saat pakai PageFrom/PageTo
            pageFrom = TryInt(SafeGet(r, colPageFrom ?? ""), 1);
            pageTo = TryInt(SafeGet(r, colPageTo ?? ""), 1);
        }

                var map = new DataMapRow
        {
            SKUInduk = skuInduk,
            NoRef = noRef,
            Variasi = varr,
            FilePath = SafeGet(r, colPdf ?? ""),
            Printer = SafeGet(r, colPrinter ?? ""),
            PageFrom = Math.Max(1, pageFrom),
            PageTo = pageTo < 0 ? 0 : pageTo, // 0 artinya sampai akhir
            Copies = TryInt(SafeGet(r, colCopies ?? ""), 1),
            Duplex = SafeGet(r, colDuplex ?? ""),
            Paper = SafeGet(r, colPaper ?? ""),
            SearchAlias = r["Search Alias"]?.ToString()?.Trim(),
        };

        // ✅ Key utama (request): model_sku + item_sku (NoRef=ModelSKU, SKUInduk=ItemSKU)
if (!string.IsNullOrWhiteSpace(map.NoRef) && !string.IsNullOrWhiteSpace(map.SKUInduk))
    _dataMap[KeyModelItem(map.NoRef, map.SKUInduk)] = map;

// ✅ Fallback keys (biar tetap kompatibel sama mapper lama)
if (!string.IsNullOrWhiteSpace(map.SKUInduk) && !_dataMap.ContainsKey(KeySkuIndukOnly(map.SKUInduk)))
    _dataMap[KeySkuIndukOnly(map.SKUInduk)] = map;

if (!string.IsNullOrWhiteSpace(map.NoRef) && !_dataMap.ContainsKey(KeyRefOnly(map.NoRef)))
    _dataMap[KeyRefOnly(map.NoRef)] = map;

if (!string.IsNullOrWhiteSpace(map.NoRef) || !string.IsNullOrWhiteSpace(map.Variasi))
    _dataMap[KeyRefVar(map.NoRef, map.Variasi)] = map;

RebuildSearchIndex();
}
}

        private void LoadTopOrderProcessFromDb(int top = 20)
        {
            Rows.Clear();

            using var con = OpenDb();
            con.Open();

            using var cmd = con.CreateCommand();
            cmd.CommandText = @"
SELECT order_sn, item_key, item_name, model_name, qty, status, create_time, model_sku, item_sku, printed
FROM order_process
ORDER BY create_time DESC, id DESC
LIMIT $top;
";
            cmd.Parameters.AddWithValue("$top", top);

            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                var orderSn = rd.GetString(0);
                var itemKey = rd.GetString(1);
                var itemName = rd.GetString(2);
                var modelName = rd.GetString(3);
                var qty = rd.GetInt32(4);
                var createT = rd.GetInt64(6);

                _dataMap.TryGetValue(itemKey, out var map);

                JobRow row;
                if (map != null)
                {
                    var pageFrom = Math.Max(1, map.PageFrom);
                    var pageTo = map.PageTo <= 0 ? 0 : Math.Max(1, map.PageTo);

                    row = new JobRow
                    {
                        Index = Rows.Count + 1,
                        OrderNo = orderSn,
                        ProductName = itemName,
                        VariationName = modelName,
                        VariationCode = itemKey, // penting (unik & dedup-friendly)
                        OrderCreatedAt = UnixToLocalDateTime(createT),

                        File = (map.FilePath ?? "").Trim(),
                        Printer = !string.IsNullOrWhiteSpace(map.Printer) ? map.Printer.Trim() : (Printers.FirstOrDefault() ?? ""),
                        PageFrom = pageFrom,
                        PageTo = pageTo,
                        Copies = Math.Max(1, qty) * Math.Max(1, map.Copies),
                        Duplex = ParseDuplex(map.Duplex),
                        Paper = ParsePaper(map.Paper),
                        Pages = $"{pageFrom}{(pageTo <= 0 ? "-" : (pageTo == pageFrom ? "" : "-" + pageTo))}",
                        Status = "Ready",
                        Percent = 0,
                        TotalPages = 0
                    };
                }
                else
                {
                    row = new JobRow
                    {
                        Index = Rows.Count + 1,
                        OrderNo = orderSn,
                        ProductName = itemName,
                        VariationName = modelName,
                        VariationCode = itemKey,
                        OrderCreatedAt = UnixToLocalDateTime(createT),
                        Copies = Math.Max(1, qty),
                        Status = "UNMAPPED (DB)",
                        Percent = 0,
                        TotalPages = 0
                    };
                }

                Rows.Add(row);
            }
        }

        private void ImportOrdersAndCreateRows(string orderXlsxPath)
        {
            if (_dataMap.Count == 0)
            {
                // autoload fallback (kalau belum sempat load)
                if (File.Exists(DefaultDataMapPath))
                    LoadDataMap(DefaultDataMapPath);
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var stream = File.Open(orderXlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var ds = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });

            if (ds.Tables.Count == 0)
                throw new InvalidOperationException("Order XLSX tidak punya sheet.");

            var t = ds.Tables[0];

            string colNoRef = FindCol(t, "No.Referensi SKU", "Nomor Referensi SKU", "No Referensi SKU", "Nomor Referensi", "Product Code", "ProductCode", "SKU ID", "SKUID", "SKU");
            string? colSkuInduk = FindColOptional(t, "SKU Induk", "SKUInduk", "Induk SKU", "Parent SKU", "SKU Parent", "SKU Induk (Parent)");
            string? colVar = FindColOptional(t, "Nama Variasi", "Variasi", "Variant");
string colQty = FindCol(t, "Jumlah", "Qty", "Quantity");
            string colOrderNo = FindCol(t, "No. Pesanan", "No Pesanan", "Order No", "Order Number");
            string? colProductName = FindColOptional(t, "Nama Produk", "Product Name", "NamaProduk", "Produk", "Product");
            string? colCreated = FindColOptional(t, "Waktu Pesanan Dibuat", "Waktu pesanan dibuat", "Order Created", "Created Time", "Created At");

            foreach (DataRow r in t.Rows)
            {
                var noRef = (r[colNoRef]?.ToString() ?? "").Trim();
                var skuInduk = SafeGet(r, colSkuInduk ?? "").Trim();
                var varr = SafeGet(r, colVar ?? "").Trim();
                int qty = TryInt(r[colQty]?.ToString(), 1);
                var orderNo = (r[colOrderNo]?.ToString() ?? "").Trim();
                if (string.IsNullOrWhiteSpace(orderNo))
                    orderNo = $"(no-order-no) row-{Rows.Count + 1}";
                var productName = (!string.IsNullOrWhiteSpace(colProductName) && t.Columns.Contains(colProductName))
                    ? (r[colProductName]?.ToString() ?? "").Trim()
                    : "";
                var createdAt = (!string.IsNullOrWhiteSpace(colCreated) && t.Columns.Contains(colCreated)) ? TryDateTimeCell(r[colCreated]) : null;
                // ✅ Key priority: SKU Induk + NoRef (utama), lalu fallback yang lama
                string? key = null;
                DataMapRow? map = null;

                if (!string.IsNullOrWhiteSpace(skuInduk) && !string.IsNullOrWhiteSpace(noRef))
                {
                    key = KeySkuIndukRef(skuInduk, noRef);
                    _dataMap.TryGetValue(key, out map);
                }

                if (map == null && !string.IsNullOrWhiteSpace(noRef))
                {
                    key = KeyRefVar(noRef, varr);
                    _dataMap.TryGetValue(key, out map);
                }

                if (map == null && !string.IsNullOrWhiteSpace(noRef))
                {
                    key = KeyRefOnly(noRef);
                    _dataMap.TryGetValue(key, out map);
                }

                if (map == null && !string.IsNullOrWhiteSpace(skuInduk))
                {
                    key = KeySkuIndukOnly(skuInduk);
                    _dataMap.TryGetValue(key, out map);
                }
if (map != null)
                {
                    // Pdf path: kalau yang disimpan folder, kamu bisa taruh file pdf full path.
                    // Aku treat kolom itu sebagai "path apa adanya".
                    string pdf = (map.FilePath ?? "").Trim();

                    var row = new JobRow
                    {
                        Index = Rows.Count + 1,
                        OrderNo = orderNo,
                        VariationName = varr,
                        VariationCode = noRef,
                        ProductName = productName,
                        OrderCreatedAt = createdAt,
                        File = pdf,
                        Printer = !string.IsNullOrWhiteSpace(map.Printer) ? map.Printer : (Printers.FirstOrDefault() ?? ""),
                        PageFrom = Math.Max(1, map.PageFrom),
                        PageTo = map.PageTo <= 0 ? 0 : Math.Max(1, map.PageTo),
                        Copies = Math.Max(1, qty) * Math.Max(1, map.Copies),
                        Duplex = ParseDuplex(map.Duplex),
                        Paper = ParsePaper(map.Paper),
                        Pages = $"{Math.Max(1, map.PageFrom)}{(map.PageTo <= 0 ? "-" : (map.PageTo == map.PageFrom ? "" : "-" + map.PageTo))}",     // legacy
                        Status = "Ready",
                        Percent = 0,
                        TotalPages = 0
                    };

                    Rows.Add(row);
                }
                else
                {
                    // Unmapped row tetap dibuat biar kelihatan yang missing
                    Rows.Add(new JobRow
                    {
                        Index = Rows.Count + 1,
                        OrderNo = orderNo,
                        VariationName = varr,
                        VariationCode = noRef,
                        OrderCreatedAt = createdAt,
                        File = "",
                        Printer = "",
                        PageFrom = 1,
                        PageTo = 1,
                        Copies = Math.Max(1, qty),
                        Duplex = DuplexMode.Simplex,
                        Paper = PaperPreset.Default,
                        Pages = "",
                        Status = "UNMAPPED",
                        Percent = 0,
                        TotalPages = 0
                    });
                }
            }
        
            // ✅ Sort: paling baru -> paling lama (berdasarkan "Waktu Pesanan Dibuat")
            {
                var sorted = Rows
                    .OrderByDescending(x => x.OrderCreatedAt ?? DateTime.MinValue)
                    .ToList();

                Rows.Clear();
                int idxNo = 1;
                foreach (var r in sorted)
                {
                    r.Index = idxNo++;
                    Rows.Add(r);
                }
            }
}

        private static DuplexMode ParseDuplex(string? v)
        {
            if (string.IsNullOrWhiteSpace(v)) return DuplexMode.Simplex;

            var s = v.Trim().ToLowerInvariant();

            // ✅ Format baru:
            // - "L" = Long Edge
            // - "N" = No duplex (simplex)
            // ✅ Tetap dukung format lama: "Long", "Short", "DuplexLong", dst.
            if (s == "n" || s == "no" || s == "none" || s == "0" || s == "-" || s == "simplex")
                return DuplexMode.Simplex;

            if (s == "l" || s == "long" || s.Contains("longedge") || s.Contains("duplexlong"))
                return DuplexMode.DuplexLongEdge;

            if (s == "s" || s == "short" || s.Contains("shortedge") || s.Contains("duplexshort"))
                return DuplexMode.DuplexShortEdge;

            // fallback: kalau ada kata "duplex" tapi gak jelas, anggap long edge
            if (s.Contains("duplex") || s.Contains("double"))
                return DuplexMode.DuplexLongEdge;

            return DuplexMode.Simplex;
        }

        private static PaperPreset ParsePaper(string? v)
        {
            if (string.IsNullOrWhiteSpace(v)) return PaperPreset.Default;
            var s = v.Trim().ToUpperInvariant();

            return s switch
            {
                "A4" => PaperPreset.A4,
                "A5" => PaperPreset.A5,
                "A6" => PaperPreset.A6,
                "B5" => PaperPreset.B5JIS,
                "B5JIS" => PaperPreset.B5JIS,
                _ => PaperPreset.Default
            };
        }

        private static int TryInt(string? v, int def)
            => int.TryParse((v ?? "").Trim(), out var x) ? x : def;

        private static DateTime? TryDateTimeCell(object? v)
        {
            if (v == null || v == DBNull.Value) return null;

            // ExcelReader sering return DateTime langsung
            if (v is DateTime dt) return dt;

            // kadang return double (OADate)
            if (v is double d)
            {
                try { return DateTime.FromOADate(d); } catch { }
            }

            var s = v.ToString();
            if (string.IsNullOrWhiteSpace(s)) return null;

            // coba parse normal (mengikuti culture OS)
            if (DateTime.TryParse(s, out var dt2)) return dt2;

            // fallback: parse invariant
            if (DateTime.TryParse(s, System.Globalization.CultureInfo.InvariantCulture,
                                  System.Globalization.DateTimeStyles.AssumeLocal, out var dt3))
                return dt3;

            return null;
        }

        private static string SafeGet(DataRow r, string colName)
        {
            if (string.IsNullOrWhiteSpace(colName)) return "";
            return r.Table.Columns.Contains(colName) ? (r[colName]?.ToString() ?? "") : "";
        }

        private static string FindCol(DataTable t, params string[] candidates)
        {
            foreach (var c in candidates)
            {
                if (t.Columns.Contains(c)) return c;
            }

            // fallback: fuzzy match (case-insensitive, remove spaces/dots)
            string Norm(string s)
            {
                var x = (s ?? "").ToLowerInvariant();
                x = x.Replace(" ", "").Replace(".", "").Replace("_", "").Replace("-", "");
                return x;
            }

            var cols = t.Columns.Cast<DataColumn>().Select(dc => dc.ColumnName).ToList();
            var normCols = cols.ToDictionary(x => Norm(x), x => x);

            foreach (var cand in candidates)
            {
                var nc = Norm(cand);
                if (normCols.TryGetValue(nc, out var real)) return real;
            }

            // kalau kolom tidak wajib, boleh return empty.
            // tapi untuk SKU/Variasi di DataMap & Order, ini harus ada.
            throw new InvalidOperationException(
                "Kolom tidak ditemukan: " + string.Join(" / ", candidates) +
                "\nKolom yang ada: " + string.Join(", ", cols));
        }

        public class DataMapRow
        {
            // New fields
            public string SKUInduk { get; set; } = "";
            public string NoRef { get; set; } = "";
            public string Variasi { get; set; } = "";
            public string FilePath { get; set; } = "";

            public string? SearchAlias { get; set; }

            // Backward-compatible aliases (jangan hapus supaya fungsi lama tetap jalan)
            public string SKU { get => NoRef; set => NoRef = value ?? ""; } // dulu dipakai sebagai No.Referensi SKU
            public string PdfPathOrFolder { get => FilePath; set => FilePath = value ?? ""; } // dulu: Folder Path

            public string Printer { get; set; } = "";
            public int PageFrom { get; set; } = 1;
            public int PageTo { get; set; } = 1; // 0 allowed = to end
            public int Copies { get; set; } = 1;
            public string Duplex { get; set; } = "";
            public string Paper { get; set; } = "";
        }

    }

    // =====================
    // Models (UNCHANGED)
    // =====================



    public enum DuplexMode
    {
        Simplex,
        DuplexLongEdge,
        DuplexShortEdge
    }

    public enum PaperPreset
    {
        Default, // use driver default / PDF size
        A4,
        A5,
        A6,
        B5JIS,   // use a dedicated Windows printer profile for true 192x257mm
    }

    public sealed class ResiRow : INotifyPropertyChanged
    {
        public string OrderSn { get => _orderSn; set { if (_orderSn == value) return; _orderSn = value; On(); } }
        private string _orderSn = "";

        public long CreateTimeUnix
        {
            get => _createTimeUnix;
            set
            {
                if (_createTimeUnix == value) return;
                _createTimeUnix = value;
                On();
                On(nameof(OrderCreatedText));
            }
        }

        private long _createTimeUnix;

        public string OrderCreatedText => CreateTimeUnix > 0
            ? DateTimeOffset.FromUnixTimeSeconds(CreateTimeUnix).ToLocalTime().ToString("yyyy-MM-dd HH:mm",
                CultureInfo.InvariantCulture)
            : "";

        public string? PdfPath
        {
            get => _pdfPath;
            set
            {
                if (_pdfPath == value) return;
                _pdfPath = value;
                On();
                On(nameof(PdfStatusText));
                On(nameof(HasPdfFile));
            }
        }

        private string? _pdfPath;

        public bool HasPdfFile => !string.IsNullOrWhiteSpace(PdfPath) && File.Exists(PdfPath!);

        public string PdfStatusText => HasPdfFile ? "Sudah ada" : "Belum diunduh";

        public bool ResiPrinted
        {
            get => _resiPrinted;
            set
            {
                if (_resiPrinted == value) return;
                _resiPrinted = value;
                On();
                On(nameof(ResiPrintedLabel));
            }
        }

        private bool _resiPrinted;

        public string ResiPrintedLabel => ResiPrinted ? "Sudah dicetak" : "Belum dicetak label";

        public event PropertyChangedEventHandler? PropertyChanged;

        private void On([CallerMemberName] string? n = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    public class JobRow : INotifyPropertyChanged
    {
        public int Index { get => _index; set { _index = value; On(); } }
        public string File { get => _file; set { _file = value; On(); } }
        public string Printer { get => _printer; set { _printer = value; On(); } }

        // sync guard supaya nggak loop
        private bool _syncingPages;

        public string OrderNo { get => _orderNo; set { _orderNo = value; On(); } }
        private string _orderNo = "";


        

        public string ProductName { get => _productName; set { _productName = value; On(); } }
        private string _productName = "";
public string VariationName { get => _variationName; set { _variationName = value; On(); } }
        private string _variationName = "";

        public long OrderProcessId { get => _orderProcessId; set { _orderProcessId = value; On(); On(nameof(CanTogglePrinted)); } }
        private long _orderProcessId;

        // Only Shopee rows from DB can be toggled
        public bool CanTogglePrinted { get => _canTogglePrinted; set { _canTogglePrinted = value; On(); } }
        private bool _canTogglePrinted;

        public bool IsPrinted { get => _isPrinted; set { if (_isPrinted == value) return; _isPrinted = value; On(); } }
        private bool _isPrinted;

        // kode variasi / No. Referensi yang dipakai untuk mapping
        public string VariationCode { get => _variationCode; set { _variationCode = value; On(); } }
        private string _variationCode = "";

        public DateTime? OrderCreatedAt { get => _orderCreatedAt; set { _orderCreatedAt = value; On(); On(nameof(OrderCreatedText)); } }
        private DateTime? _orderCreatedAt;

        public string OrderCreatedText => OrderCreatedAt.HasValue ? OrderCreatedAt.Value.ToString("yyyy-MM-dd HH:mm") : "";




        // Per-row settings
        public int PageFrom
        {
            get => _pageFrom;
            set
            {
                _pageFrom = Math.Max(1, value);
                On();
                On(nameof(PageRange));

                // auto sync legacy Pages (dipakai UI XAML)
                if (!_syncingPages)
                {
                    _syncingPages = true;
                    _pages = PageRange;
                    On(nameof(Pages));
                    _syncingPages = false;
                }
            }
        }



        public int PageTo
        {
            get => _pageTo;
            set
            {
                _pageTo = value < 0 ? 0 : value; // 0 = to end
                On();
                On(nameof(PageRange));

                // auto sync legacy Pages (dipakai UI XAML)
                if (!_syncingPages)
                {
                    _syncingPages = true;
                    _pages = PageRange;
                    On(nameof(Pages));
                    _syncingPages = false;
                }
            }
        }

        // Single field range input (e.g. "2-3", "2", "2-") for UI convenience.
        public string PageRange
        {
            get
            {
                var from = Math.Max(1, PageFrom);
                if (PageTo <= 0) return $"{from}-";
                if (PageTo == from) return $"{from}";
                return $"{from}-{PageTo}";
            }
            set
            {
                var s = (value ?? "").Trim();
                if (string.IsNullOrWhiteSpace(s))
                {
                    return;
                }

                var m = Regex.Match(s, @"^\s*(\d+)?\s*(?:-\s*(\d+)?)?\s*$");
                if (!m.Success)
                {
                    On();
                    return;
                }

                int? a = null, b = null;
                if (m.Groups[1].Success && int.TryParse(m.Groups[1].Value, out var aa)) a = aa;
                if (m.Groups[2].Success && int.TryParse(m.Groups[2].Value, out var bb)) b = bb;

                int from = Math.Max(1, a ?? 1);
                int to;

                if (a.HasValue && !s.Contains("-")) to = from;                 // "N"
                else if (!a.HasValue && b.HasValue) to = Math.Max(0, b.Value); // "-M"
                else if (a.HasValue && !b.HasValue) to = 0;                    // "N-"
                else to = Math.Max(0, b ?? from);                              // "N-M"

                if (to > 0 && to < from) (from, to) = (to, from);

                _pageFrom = from;
                _pageTo = to;

                On(nameof(PageFrom));
                On(nameof(PageTo));
                On();
            }
        }

        // ✅ Legacy text (dipakai UI XAML). Sekarang disinkronin ke PageFrom/PageTo
        // legacy alias
        public string Pages
        {
            get => _pages;
            set
            {
                _pages = value ?? "";
                On();

                if (_syncingPages) return;

                var s = (_pages ?? "").Trim();
                if (string.IsNullOrWhiteSpace(s)) return;

                var m = Regex.Match(s, @"^\s*(\d+)?\s*(?:-\s*(\d+)?)?\s*$");
                if (!m.Success) return;

                int? a = null, b = null;
                if (m.Groups[1].Success && int.TryParse(m.Groups[1].Value, out var aa)) a = aa;
                if (m.Groups[2].Success && int.TryParse(m.Groups[2].Value, out var bb)) b = bb;

                int from = Math.Max(1, a ?? 1);
                int to;

                if (a.HasValue && !s.Contains("-")) to = from;                 // "N"
                else if (!a.HasValue && b.HasValue) to = Math.Max(0, b.Value); // "-M"
                else if (a.HasValue && !b.HasValue) to = 0;                    // "N-"
                else to = Math.Max(0, b ?? from);                              // "N-M"

                if (to > 0 && to < from) (from, to) = (to, from);

                _syncingPages = true;
                _pageFrom = from;
                _pageTo = to;
                On(nameof(PageFrom));
                On(nameof(PageTo));
                On(nameof(PageRange));
                _syncingPages = false;
            }
        }


        public int Copies { get => _copies; set { _copies = Math.Max(1, value); On(); } }
        public DuplexMode Duplex { get => _duplex; set { _duplex = value; On(); } }
        public PaperPreset Paper { get => _paper; set { _paper = value; On(); } }

        /// <summary>1.0 = isi memakai penuh kotak margin (Pdfium+GDI). &lt;1 = diperkecil di dalam margin.</summary>
        public double PdfPrintScale { get; set; } = 1.0;

        /// <summary>Sumatra: monochrome. GDI+ path: <see cref="PageSettings.Color"/> = false.</summary>
        public bool PrintMonochrome { get; set; }

        /// <summary>True: cetak lewat Pdfium+GDI dengan gambar mulai kiri-atas area margin (label Shopee).</summary>
        public bool PrintFromTopLeft { get; set; }

        public bool Selected { get => _selected; set { _selected = value; On(); } }
        public string Status { get => _status; set { _status = value; On(); } }
        public int Percent { get => _percent; set { _percent = value; On(); } }
        public bool HasJob { get => _hasJob; set { _hasJob = value; On(); } }
        public int TotalPages { get => _totalPages; set { _totalPages = value; On(); } }

        public Guid? JobGuid { get => _jobGuid; set { _jobGuid = value; On(); } }

        private int _index;
        private string _file = "";
        private string _printer = "";

        private string _pageRange = "1";
        private int _pageFrom = 1;
        private int _pageTo = 1;
        private int _copies = 1;
        private DuplexMode _duplex = DuplexMode.Simplex;
        private PaperPreset _paper = PaperPreset.Default;

        private string _pages = "";
        private bool _selected;
        private string _status = "";
        private int _percent;
        private bool _hasJob;
        private int _totalPages;
        private Guid? _jobGuid;

        private int _printTotalPages;
        private int _printedPages;

        public event PropertyChangedEventHandler? PropertyChanged;
        private void On([CallerMemberName] string? n = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        public int PrintTotalPages
        {
            get => _printTotalPages;
            set { _printTotalPages = value; On(); }
        }

        public int PrintedPages
        {
            get => _printedPages;
            set { _printedPages = value; On(); }
        }
    }

    public class PageThumb
    {
        public string File { get; set; } = "";
        public int PageNumber { get; set; }
        public int TotalPages { get; set; }
        public BitmapImage Image { get; set; } = default!;
        public string Caption { get; set; } = "";
    }
}
