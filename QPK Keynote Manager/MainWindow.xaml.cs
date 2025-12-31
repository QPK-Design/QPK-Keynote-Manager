using System;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QPK_Keynote_Manager
{
    public partial class MainWindow : Window
    {
        private readonly UIDocument _uidoc;
        private readonly Document _doc;

        private readonly ExternalEvent _replaceAllEvent;
        private readonly ExternalEvent _replaceSelectedEvent;

        public MainViewModel VM { get; }

        public MainWindow(UIDocument uidoc)
        {
            InitializeComponent();

            _uidoc = uidoc ?? throw new ArgumentNullException(nameof(uidoc));
            _doc = _uidoc.Document;

            // MVVM
            VM = new MainViewModel(_uidoc);
            DataContext = VM;

            // ExternalEvents (handlers should be in their own .cs files)
            _replaceAllEvent = ExternalEvent.Create(new ReplaceAllHandler(this));
            _replaceSelectedEvent = ExternalEvent.Create(new ReplaceSelectedHandler(this));

            // Optional: close cleanup
            Closed += MainWindow_Closed;
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            // Nothing required here right now.
            // If later you add modeless window tracking, do it here.
        }

        /// <summary>
        /// Used by external event handlers to fetch currently selected row in the DataGrid.
        /// </summary>
        internal IReplaceRow GetSelectedRow()
        {
            return VM?.SelectedResult;
        }

        private void PreviewFindReplace_Click(object sender, RoutedEventArgs e)
        {
            if (VM?.SelectedScope == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Select a preview scope first.");
                return;
            }

            switch (VM.SelectedScope.Kind)
            {
                case FindReplaceScopeKind.SheetNames:
                    VM.PreviewSheetNames();
                    TaskDialog.Show("QPK Keynote Manager",
                        $"Preview complete.\n\nFound {VM.SheetNameResults.Count} sheet name(s) to change.");
                    break;

                case FindReplaceScopeKind.Keynotes:
                    VM.PreviewKeynotes();
                    MessageBox.Show(
                        $"Scanned schedules and found {VM.KeynoteResults.Count} keynote comment(s) to change.",
                        "Preview Complete",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    break;


                case FindReplaceScopeKind.ViewTitles:
                    TaskDialog.Show("QPK Keynote Manager",
                        "View Titles preview is not wired yet in this step.");
                    break;

                default:
                    TaskDialog.Show("QPK Keynote Manager", "Unknown scope selected.");
                    break;
            }
        }

        private void ReplaceSelected_Click(object sender, RoutedEventArgs e)
        {
            _replaceSelectedEvent?.Raise();
        }

        private void ReplaceAll_Click(object sender, RoutedEventArgs e)
        {
            _replaceAllEvent?.Raise();
        }
    }
}
