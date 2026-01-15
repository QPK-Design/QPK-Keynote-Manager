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

            // ExternalEvents
            _replaceAllEvent = ExternalEvent.Create(new ReplaceAllHandler(this));
            _replaceSelectedEvent = ExternalEvent.Create(new ReplaceSelectedHandler(this));

            // Optional cleanup hook
            Closed += MainWindow_Closed;
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            // Nothing required here right now.
        }

        /// <summary>
        /// Used by external event handlers to fetch currently selected row in the DataGrid.
        /// </summary>
        internal IReplaceRow GetSelectedRow()
        {
            return VM?.SelectedResult;
        }

        /// <summary>
        /// Sets the type comment text (BuiltInParameter.ALL_MODEL_TYPE_COMMENTS first, then fallbacks).
        /// Returns true if Revit parameter was successfully set; false if no-op or cannot set.
        /// </summary>
        internal bool SetTypeComment(Element tElem, string newText, bool isCaseSensitive)
        {
            if (tElem == null) return false;

            newText ??= string.Empty;

            // 1) Built-in Type Comments
            Parameter p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null && !p.IsReadOnly)
            {
                string current = p.AsString() ?? string.Empty;

                bool equal = isCaseSensitive
                    ? current.Equals(newText, StringComparison.Ordinal)
                    : current.Equals(newText, StringComparison.OrdinalIgnoreCase);

                if (equal)
                    return false; // no-op (already matches)

                return p.Set(newText);
            }

            // 2) Fallback named parameters
            string[] names = { "Type Comments", "Comments", "Comment", "COMMENT" };
            foreach (string nm in names)
            {
                p = tElem.LookupParameter(nm);
                if (p != null && !p.IsReadOnly)
                {
                    string current = p.AsString() ?? string.Empty;

                    bool equal = isCaseSensitive
                        ? current.Equals(newText, StringComparison.Ordinal)
                        : current.Equals(newText, StringComparison.OrdinalIgnoreCase);

                    if (equal)
                        return false; // no-op

                    return p.Set(newText);
                }
            }

            return false;
        }

        private void PreviewFindReplace_Click(object sender, RoutedEventArgs e)
        {
            if (VM == null)
                return;

            // Scan ALL enabled scopes (checkboxes), not the dropdown selection
            VM.PreviewEnabledScopes();

            // Summary
            int k = VM.KeynoteResults?.Count ?? 0;
            int s = VM.SheetNameResults?.Count ?? 0;
            int v = VM.ViewTitleResults?.Count ?? 0;

            TaskDialog.Show("QPK Keynote Manager",
                $"Preview complete.\n\n" +
                $"Keynotes: {k}\n" +
                $"Sheet Names: {s}\n" +
                $"View Titles: {v}\n\n" +
                $"Use the dropdown to review each scope's results.");
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
