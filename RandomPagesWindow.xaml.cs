using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using PdfiumViewer;

namespace PaperbellAppDotNet
{
    public sealed class RandomPageMapPick
    {
        public string Display { get; init; } = "";
        public string PdfPath { get; init; } = "";
    }

    public partial class RandomPagesWindow : Window
    {
        private readonly IReadOnlyList<RandomPageMapPick> _planner;
        private readonly IReadOnlyList<RandomPageMapPick> _loose;
        private readonly Func<string, string?> _resolvePrinterContains;

        /// <summary>Baris antrian yang dihasilkan (satu baris per halaman).</summary>
        public IReadOnlyList<JobRow>? GeneratedRows { get; private set; }

        public RandomPagesWindow(
            IReadOnlyList<RandomPageMapPick> planner,
            IReadOnlyList<RandomPageMapPick> loose,
            Func<string, string?> resolvePrinterContains)
        {
            InitializeComponent();
            _planner = planner;
            _loose = loose;
            _resolvePrinterContains = resolvePrinterContains;

            RbPlanner.Checked += (_, _) => RefreshPdfCombo();
            RbLoose.Checked += (_, _) => RefreshPdfCombo();
            Loaded += (_, _) =>
            {
                RefreshPdfCombo();
                UpdateHint();
            };
        }

        private void RefreshPdfCombo()
        {
            var list = RbLoose.IsChecked == true ? _loose : _planner;
            CmbPdf.ItemsSource = list;
            CmbPdf.SelectedIndex = list.Count > 0 ? 0 : -1;
            UpdateHint();
        }

        private void UpdateHint()
        {
            var list = RbLoose.IsChecked == true ? _loose : _planner;
            if (list.Count == 0)
                TxtHint.Text = "Tidak ada baris Data Map dengan Group P atau L dan path PDF — periksa XLSX di folder config.";
            else
                TxtHint.Text = $"Sumber terpilih: {(RbLoose.IsChecked == true ? "Loose Leaf → printer mengandung \"Brother\"" : "Planner → printer mengandung \"Epson\"")}.";
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (CmbPdf.SelectedItem is not RandomPageMapPick pick)
            {
                MessageBox.Show(this, "Pilih PDF sumber.", "Random pages", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var path = (pick.PdfPath ?? "").Trim();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show(this, "File PDF tidak ditemukan:\n" + path, "Random pages", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse((TxtCount.Text ?? "").Trim(), NumberStyles.Integer, CultureInfo.CurrentCulture, out var n) ||
                n < 1)
            {
                MessageBox.Show(this, "Masukkan jumlah halaman (N) berupa bilangan bulat ≥ 1.", "Random pages",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var wantLoose = RbLoose.IsChecked == true;
            var printerToken = wantLoose ? "Brother" : "Epson";
            var resolved = _resolvePrinterContains(printerToken);
            if (string.IsNullOrWhiteSpace(resolved))
            {
                MessageBox.Show(this,
                    $"Tidak ada printer terpasang yang cocok dengan \"{printerToken}\". Pasang driver atau ubah nama printer.",
                    "Random pages", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int total;
            try
            {
                using var doc = PdfDocument.Load(path);
                total = doc.PageCount;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Gagal membaca PDF:\n" + ex.Message, "Random pages", MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            if (n > total)
            {
                MessageBox.Show(this, $"PDF hanya punya {total} halaman; N tidak boleh lebih besar.", "Random pages",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var paper = RbB5.IsChecked == true ? PaperPreset.B5JIS : PaperPreset.A5;
            var pages = PickDistinctRandomPageNumbers(total, n);

            var rows = new List<JobRow>();
            var tag = wantLoose ? "RND-L" : "RND-P";
            foreach (var p in pages)
            {
                rows.Add(new JobRow
                {
                    File = path,
                    Printer = resolved,
                    PageFrom = p,
                    PageTo = p,
                    Copies = 1,
                    Duplex = DuplexMode.Simplex,
                    Paper = paper,
                    Pages = p.ToString(CultureInfo.InvariantCulture),
                    Status = "Ready",
                    Percent = 0,
                    TotalPages = total,
                    OrderNo = tag,
                    ProductName = "Random pages",
                    VariationName = Path.GetFileName(path)
                });
            }

            GeneratedRows = rows;
            DialogResult = true;
        }

        /// <summary>Tanpa duplikat; urutan acak.</summary>
        internal static List<int> PickDistinctRandomPageNumbers(int totalPages, int count, Random? rng = null)
        {
            rng ??= Random.Shared;
            if (count < 1) throw new ArgumentOutOfRangeException(nameof(count));
            if (totalPages < 1) throw new ArgumentOutOfRangeException(nameof(totalPages));
            if (count > totalPages) throw new ArgumentException("count cannot exceed totalPages.");
            return Enumerable.Range(1, totalPages).OrderBy(_ => rng.Next()).Take(count).ToList();
        }
    }
}
