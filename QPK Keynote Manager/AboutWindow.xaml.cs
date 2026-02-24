using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;

namespace QPK_Keynote_Manager
{
    public partial class AboutWindow : Window
    {
        public string ReleaseDate { get; }
        public string Description { get; }
        public string GithubUri { get; }

        public AboutWindow()
        {
            InitializeComponent();

            // Use build date / today's date — simplest “current release date” behavior
            ReleaseDate = DateTime.Today.ToString("yyyy-MM-dd");

            Description =
                "QPK Keynote Manager provides a fast, find/replace workflow " +
                "for Revit keynotes, sheet names, and view titles/names, with preview, " +
                "selective replacement, and bulk replacement to help keep documentation consistent.";

            GithubUri = "https://github.com/QPK-Design/QPK-Keynote-Manager/";

            DataContext = this;
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            }
            catch
            {
                MessageBox.Show(
                    "Unable to open the link. Please copy/paste it into your browser:\n\n" + e.Uri.AbsoluteUri,
                    "QPK Keynote Manager",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }

            e.Handled = true;
        }
    }
}
