using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace PaperbellAppDotNet
{
    public partial class CuciDarahPrintDialog : Window
    {
        public const string DefaultPdfPath = @"H:\My Drive\Paperbell\PINK.pdf";

        public string PdfPath => TxtPdfPath.Text.Trim();

        public string PaperName =>
            (CmbPaper.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "A5";

        public CuciDarahPrintDialog()
        {
            InitializeComponent();
            TxtPdfPath.Text = DefaultPdfPath;
            TxtPdfPath.CaretIndex = TxtPdfPath.Text.Length;
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Pilih dokumen Cuci Darah",
                Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*",
                CheckFileExists = true
            };

            var currentPath = PdfPath;
            if (File.Exists(currentPath))
            {
                dialog.InitialDirectory = Path.GetDirectoryName(currentPath);
                dialog.FileName = Path.GetFileName(currentPath);
            }

            if (dialog.ShowDialog(this) == true)
                TxtPdfPath.Text = dialog.FileName;
        }

        private void Print_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(PdfPath))
            {
                MessageBox.Show(this, "Pilih file PDF terlebih dahulu.", "Cuci Darah",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                TxtPdfPath.Focus();
                return;
            }

            if (!File.Exists(PdfPath))
            {
                MessageBox.Show(this, $"Dokumen tidak ditemukan:\n{PdfPath}", "Cuci Darah",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                TxtPdfPath.Focus();
                return;
            }

            DialogResult = true;
        }
    }
}
