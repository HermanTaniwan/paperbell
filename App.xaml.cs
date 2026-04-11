using System.Configuration;
using System.Data;
using System.Windows;

namespace Paperbell_App
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            this.DispatcherUnhandledException += (s, e) =>
            {
                MessageBox.Show(e.Exception.ToString(), "Unhandled UI Error");
                e.Handled = true;
            };

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                MessageBox.Show(e.ExceptionObject?.ToString() ?? "(null)", "Unhandled Domain Error");
            };
        }
    }

}
