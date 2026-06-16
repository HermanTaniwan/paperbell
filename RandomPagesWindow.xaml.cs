using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PaperbellAppDotNet
{
    public sealed class RandomPageMapPick
    {
        public string Display { get; init; } = "";
        public string PdfPath { get; init; } = "";
        public string SearchText { get; init; } = "";
    }

    public partial class RandomPagesWindow : Window
    {
        private readonly IReadOnlyList<RandomPageMapPick> _planner;
        private readonly IReadOnlyList<RandomPageMapPick> _loose;
        private readonly Func<string, string?> _resolvePrinterContains;

        private bool _generateInProgress;

        /// <summary>Satu baris antrian: PDF gabungan halaman acak (satu kali print).</summary>
        public JobRow? GeneratedRow { get; private set; }

        public RandomPagesWindow(
            IReadOnlyList<RandomPageMapPick> planner,
            IReadOnlyList<RandomPageMapPick> loose,
            Func<string, string?> resolvePrinterContains)
        {
            InitializeComponent();
            _planner = planner;
            _loose = loose;
            _resolvePrinterContains = resolvePrinterContains;

            RbPlanner.Checked += (_, _) => UpdatePoolUi();
            RbLoose.Checked   += (_, _) => UpdatePoolUi();
            RbA5.Checked      += (_, _) => UpdatePoolUi();
            RbB5.Checked      += (_, _) => UpdatePoolUi();

            TxtExclude.TextChanged += (_, _) => UpdatePoolUi();
            Loaded += (_, _) => UpdatePoolUi();
        }

        private static string[] ParseExcludeTokens(string? raw)
        {
            return (raw ?? "")
                .Split(new[] { ',', ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => x.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static List<RandomPageMapPick> ApplyExclude(
            IReadOnlyList<RandomPageMapPick> source,
            IReadOnlyList<string> tokens)
        {
            if (tokens.Count == 0)
                return source.ToList();

            var outList = new List<RandomPageMapPick>(source.Count);
            foreach (var pick in source)
            {
                var hay = (pick.SearchText ?? pick.Display ?? "").Trim();
                var excluded = tokens.Any(t => hay.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0);
                if (!excluded)
                    outList.Add(pick);
            }

            return outList;
        }

        /// <summary>
        /// Filter pool berdasarkan paper size: A5 â†’ hanya PDF dengan "A5" di nama file,
        /// B5 â†’ hanya PDF dengan "B5". File tanpa marker ukuran diikutkan di semua mode.
        /// Jika pool < N setelah filter, BuildMergedJob akan loop/repeat PDF yang ada.
        /// </summary>
        private static List<RandomPageMapPick> ApplyPaperFilter(
            IReadOnlyList<RandomPageMapPick> source, PaperPreset paper)
        {
            var key      = paper == PaperPreset.A5 ? "A5" : "B5";
            var opposite = paper == PaperPreset.A5 ? "B5" : "A5";
            return source.Where(x =>
            {
                var fn = Path.GetFileName((x.PdfPath ?? "").Trim());
                var hasKey      = fn.Contains(key,      StringComparison.OrdinalIgnoreCase);
                var hasOpposite = fn.Contains(opposite, StringComparison.OrdinalIgnoreCase);
                // Ikutkan kalau file punya marker yang dipilih, atau tidak punya marker sama sekali.
                return hasKey || (!hasKey && !hasOpposite);
            }).ToList();
        }

        private void UpdatePoolUi()
        {
            var paper    = RbB5.IsChecked == true ? PaperPreset.B5JIS : PaperPreset.A5;
            var baseList = RbLoose.IsChecked == true ? _loose : _planner;
            if (baseList.Count == 0)
            {
                TxtPoolInfo.Text = "Tidak ada file PDF untuk grup ini di Data Map.";
                TxtHint.Text = "Pastikan ada baris dengan kolom Group yang sesuai dan path PDF terisi.";
                return;
            }

            var excludeTokens = ParseExcludeTokens(TxtExclude.Text);
            var list = ApplyPaperFilter(ApplyExclude(baseList, excludeTokens), paper);

            var paperLabel = paper == PaperPreset.A5 ? "A5" : "B5";
            var suffixExcluded = (excludeTokens.Length > 0 || list.Count < baseList.Count)
                ? $" (filter {paperLabel}" +
                  (excludeTokens.Length > 0 ? $" + exclude {excludeTokens.Length} kata" : "") +
                  $"; tersisa {list.Count} dari {baseList.Count})"
                : "";

            if (RbLoose.IsChecked == true)
            {
                TxtPoolInfo.Text =
                    list.Count == 1
                        ? "Satu file PDF di pool â€” hanya N = 1. Loose Leaf: satu spread bolak-balik acak: halaman ganjil berpasangan dengan berikutnya (mis. 3â€“4); file minimal 2 halaman."
                        : $"{list.Count} file di pool â€” Loose Leaf: N PDF acak, tiap PDF satu spread (hal. pertama ganjil + berikutnya, mis. 3â€“4) â†’ total 2Ã—N halaman. N â‰¤ {list.Count}, tiap PDF minimal 2 halaman.";
                GbCount.Header = "N (jumlah PDF)";
            }
            else
            {
                TxtPoolInfo.Text =
                    list.Count == 1
                        ? "Satu file PDF di pool â€” hanya N = 1 (satu halaman acak). Satu PDF, cetak sekali."
                        : $"{list.Count} file di pool â€” Planner: pilih acak N PDF berbeda, tiap PDF 1 halaman acak â†’ total N halaman. N â‰¤ {list.Count}.";
                GbCount.Header = "N (jumlah PDF)";
            }
            TxtPoolInfo.Text += suffixExcluded;

            if (!_generateInProgress)
            {
                TxtHint.Text = RbLoose.IsChecked == true
                    ? "Printer default: nama printer mengandung \"Brother\" (Loose Leaf)."
                    : "Printer default: nama printer mengandung \"L3210\" (Epson L3210, Planner).";
            }
        }

        private async void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (_generateInProgress)
                return;

            var paper = RbB5.IsChecked == true ? PaperPreset.B5JIS : PaperPreset.A5;
            var baseList = RbLoose.IsChecked == true ? _loose : _planner;
            var excludeTokens = ParseExcludeTokens(TxtExclude.Text);
            var list = ApplyPaperFilter(ApplyExclude(baseList, excludeTokens), paper);
            if (list.Count == 0)
            {
                MessageBox.Show(this,
                    excludeTokens.Length > 0
                        ? "Setelah exclude, tidak ada PDF yang tersisa untuk grup ini."
                        : "Tidak ada PDF untuk grup ini.",
                    "Random pages", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse((TxtCount.Text ?? "").Trim(), NumberStyles.Integer, CultureInfo.CurrentCulture, out var n) ||
                n < 1)
            {
                MessageBox.Show(this, "Masukkan N (jumlah PDF) berupa bilangan bulat â‰¥ 1.", "Random pages",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Tidak ada batasan n > list.Count â€” BuildMergedJob akan repeat PDF
            // yang ada jika pool lebih kecil dari n.

            var wantLoose = RbLoose.IsChecked == true;
            var printerToken = wantLoose ? "Brother" : "L3210";
            var resolved = _resolvePrinterContains(printerToken);
            if (string.IsNullOrWhiteSpace(resolved))
            {
                MessageBox.Show(this,
                    $"Tidak ada printer terpasang yang cocok dengan \"{printerToken}\". Pasang driver atau ubah nama printer.",
                    "Random pages", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var tag = wantLoose ? "RND-L" : "RND-P";

            var listSnapshot = list.ToArray();

            _generateInProgress = true;
            RootLayout.IsEnabled = false;
            Mouse.OverrideCursor = Cursors.Wait;
            var totalPagesOut = wantLoose ? 2 * n : n;
            TxtHint.Text =
                $"Membaca & menggabungkan {totalPagesOut} halaman ({n} PDF)â€¦ (file rusak/hilang dilewati)";

            try
            {
                var (row, pickSummary) = await Task.Run(
                    () => BuildMergedJob(listSnapshot, n, paper, resolved, tag, wantLoose))
                    .ConfigureAwait(true);

                GeneratedRow = row;
                TxtHint.Text = "Terpilih: " + pickSummary;
                DialogResult = true;
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(this, ex.Message, "Random pages", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "Gagal menggabungkan PDF:\n" + ex.Message +
                    "\n\nJika PDF terproteksi/aneh, coba file lain atau duplikat tanpa proteksi.",
                    "Random pages", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _generateInProgress = false;
                RootLayout.IsEnabled = true;
                Mouse.OverrideCursor = null;
                UpdatePoolUi();
            }
        }

        /// <summary>
        /// Acak satu pasangan spread bolak-balik: (k, k+1) dengan k ganjil (mis. 3 dan 4).
        /// Ada âŒŠpageCount/2âŒ‹ kemungkinan (1â€“2, 3â€“4, â€¦).
        /// </summary>
        private static (int First, int Second) PickRandomOddSpreadPair(int pageCount, Random rng)
        {
            if (pageCount < 2)
                throw new ArgumentOutOfRangeException(nameof(pageCount));

            var spreadCount = pageCount / 2;
            var idx = rng.Next(spreadCount);
            var first = 1 + 2 * idx;
            return (first, first + 1);
        }

        /// <summary>Dijalankan di thread pool â€” jangan sentuh UI.</summary>
        private static (JobRow Row, string PickSummary) BuildMergedJob(
            IReadOnlyList<RandomPageMapPick> list,
            int n,
            PaperPreset paper,
            string resolvedPrinter,
            string tag,
            bool looseLeafBolakBalik)
        {
            var candidatePaths = list
                .Select(x => (x.PdfPath ?? "").Trim())
                .Where(p => p.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var existsPaths = candidatePaths.Where(File.Exists).ToArray();
            if (existsPaths.Length == 0)
                throw new InvalidOperationException("Tidak ada file PDF yang ditemukan di disk. Periksa path di Data Map.");

            // Buka SEMUA PDF yang tersedia di pool (bukan hanya N),
            // karena jika pool < N kita akan repeat/ulang PDF yang ada.
            var poolOrder = existsPaths.OrderBy(_ => Random.Shared.Next()).ToArray();
            var sources = new ConcurrentDictionary<string, PdfDocument>(StringComparer.OrdinalIgnoreCase);

            var minPagesPerFile = looseLeafBolakBalik ? 2 : 1;

            void TryOpen(string path)
            {
                try
                {
                    var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import);
                    if (doc.PageCount < minPagesPerFile)
                    {
                        doc.Dispose();
                        return;
                    }
                    sources[path] = doc;
                }
                catch { }
            }

            Parallel.ForEach(poolOrder, TryOpen);

            if (sources.Count == 0)
            {
                var extra = looseLeafBolakBalik ? " Untuk Loose Leaf tiap PDF harus punya minimal 2 halaman." : "";
                throw new InvalidOperationException(
                    $"Tidak ada PDF yang bisa dibaca dari pool.{extra}");
            }

            // Buat daftar N path yang akan dipakai.
            // Kalau pool < N: repeat PDF yang ada secara acak sampai memenuhi N.
            var availablePaths = sources.Keys.OrderBy(_ => Random.Shared.Next()).ToArray();
            var pickedPaths = new List<string>(n);
            while (pickedPaths.Count < n)
            {
                var shuffled = availablePaths.OrderBy(_ => Random.Shared.Next()).ToArray();
                foreach (var p in shuffled)
                {
                    if (pickedPaths.Count >= n) break;
                    pickedPaths.Add(p);
                }
            }
            // Acak urutan final agar PDF yang repeat tidak selalu berurutan
            pickedPaths = pickedPaths.OrderBy(_ => Random.Shared.Next()).ToList();
            var totalOut = looseLeafBolakBalik ? 2 * n : n;
            var segments = new List<(string Path, int Page)>(totalOut);
            foreach (var path in pickedPaths)
            {
                var doc = sources[path];
                if (looseLeafBolakBalik)
                {
                    var (p1, p2) = PickRandomOddSpreadPair(doc.PageCount, Random.Shared);
                    segments.Add((path, p1));
                    segments.Add((path, p2));
                }
                else
                {
                    var pageNum = Random.Shared.Next(1, doc.PageCount + 1);
                    segments.Add((path, pageNum));
                }
            }

            // Buat ringkasan pilihan untuk ditampilkan di UI
            var pickParts = new List<string>(n);
            if (looseLeafBolakBalik)
            {
                for (int si = 0; si < segments.Count; si += 2)
                {
                    var fname = Path.GetFileNameWithoutExtension(segments[si].Path);
                    pickParts.Add($"{fname} pg.{segments[si].Page}+{segments[si + 1].Page}");
                }
            }
            else
            {
                foreach (var (path, page) in segments)
                    pickParts.Add($"{Path.GetFileNameWithoutExtension(path)} pg.{page}");
            }
            var pickSummary = string.Join(" | ", pickParts);

            var dictCopy = new Dictionary<string, PdfDocument>(sources, StringComparer.OrdinalIgnoreCase);
            sources.Clear();
            var mergedPath = RandomPagesPdfMerge.MergePreOpenedToTempFile(dictCopy, segments, paper);

            var row = new JobRow
            {
                File = mergedPath,
                Printer = resolvedPrinter,
                PageFrom = 1,
                PageTo = totalOut,
                Copies = 1,
                Duplex = looseLeafBolakBalik ? DuplexMode.DuplexLongEdge : DuplexMode.Simplex,
                Paper = paper,
                Pages = totalOut == 1 ? "1" : $"1-{totalOut}",
                Status = "Ready",
                Percent = 0,
                TotalPages = totalOut,
                OrderNo = tag,
                ProductName = "Random pages",
                VariationName = pickSummary,
                DeleteTempMergedPdfAfterUse = true
            };
            return (row, pickSummary);
        }
    }
}
