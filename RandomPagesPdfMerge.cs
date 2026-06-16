using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PaperbellAppDotNet
{
    /// <summary>Menyatukan halaman dari PDF yang sudah dibuka (satu kali buka per file).</summary>
    internal static class RandomPagesPdfMerge
    {
        /// <summary>Gabungkan halaman ke file temp lalu dispose semua <paramref name="sources"/>.</summary>
        public static string MergePreOpenedToTempFile(
            Dictionary<string, PdfDocument> sources,
            IReadOnlyList<(string SourcePath, int PageOneBased)> segments,
            PaperPreset targetPaper)
        {
            if (sources == null || sources.Count == 0)
                throw new ArgumentException("Tidak ada PDF sumber yang terbuka.", nameof(sources));
            if (segments == null || segments.Count == 0)
                throw new ArgumentException("Minimal satu segmen halaman.", nameof(segments));

            var temp = Path.Combine(Path.GetTempPath(), $"paperbell_random_{Guid.NewGuid():N}.pdf");

            // Satu slot MemoryStream per segmen (urutan output dijaga lewat index).
            var pageStreams = new MemoryStream[segments.Count];

            try
            {
                // Buka tiap file sumber TEPAT SATU KALI, ekstrak semua halaman yang
                // dibutuhkan ke MemoryStream masing-masing sebelum file ditutup.
                // Ini menghindari masalah cache/reference internal PdfSharp ketika
                // file yang sama dibuka lebih dari sekali dalam satu proses.
                var indexedSegments = segments
                    .Select((s, i) => (s.SourcePath, s.PageOneBased, OutIdx: i))
                    .GroupBy(s => s.SourcePath, StringComparer.OrdinalIgnoreCase);

                foreach (var group in indexedSegments)
                {
                    if (!sources.TryGetValue(group.Key, out var cachedDoc))
                        throw new InvalidOperationException($"Sumber hilang di cache: {group.Key}");

                    using var freshSrc = PdfReader.Open(group.Key, PdfDocumentOpenMode.Import);

                    foreach (var (_, pageOneBased, outIdx) in group)
                    {
                        if (pageOneBased < 1 || pageOneBased > freshSrc.PageCount)
                            throw new InvalidOperationException(
                                $"Halaman {pageOneBased} di luar rentang 1..{freshSrc.PageCount} " +
                                $"({Path.GetFileName(group.Key)}).");

                        using var onePageDoc = new PdfDocument();
                        onePageDoc.AddPage(freshSrc.Pages[pageOneBased - 1]);

                        var ms = new MemoryStream();
                        onePageDoc.Save(ms, closeStream: false);
                        ms.Position = 0;
                        pageStreams[outIdx] = ms;
                    }
                }

                // Rakit output final dari MemoryStream yang sudah terisolasi.
                PdfDocument? output = null;
                var loadedPageDocs = new List<PdfDocument>(segments.Count);
                try
                {
                    output = new PdfDocument();
                    foreach (var ms in pageStreams)
                    {
                        var pageDoc = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
                        loadedPageDocs.Add(pageDoc);
                        output.AddPage(pageDoc.Pages[0]);
                    }
                    output.Save(temp);
                    return temp;
                }
                finally
                {
                    output?.Dispose();
                    foreach (var doc in loadedPageDocs)
                        doc.Dispose();
                }
            }
            finally
            {
                foreach (var ms in pageStreams)
                    ms?.Dispose();
                foreach (var doc in sources.Values)
                    doc.Dispose();
                sources.Clear();
            }
        }
    }
}
