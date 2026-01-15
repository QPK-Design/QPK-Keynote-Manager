using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
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
        private void ResultsDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Only react if the double-click happened on an actual DataGridRow
            if (!IsDoubleClickOnRow(e.OriginalSource as DependencyObject))
                return;

            var row = VM?.SelectedResult as IReplaceRow;
            if (row == null)
                return;

            var sheet = ResolveSheetFromRow(row);
            if (sheet == null)
            {
                TaskDialog.Show("QPK Keynote Manager",
                    "This row is not associated with a sheet (or the sheet could not be found).");
                return;
            }

            try
            {
                // Preferred in modern Revit
                _uidoc.RequestViewChange(sheet);
            }
            catch
            {
                try
                {
                    // Fallback
                    _uidoc.ActiveView = sheet;
                }
                catch
                {
                    TaskDialog.Show("QPK Keynote Manager",
                        "Unable to open the sheet. Revit may be in a state that prevents view changes.");
                    return;
                }
            }

            try
            {
                // Optional: select the sheet in the model
                _uidoc.Selection.SetElementIds(new[] { sheet.Id });
            }
            catch
            {
                // harmless if selection fails
            }
        }

        private void HelpAbout_Click(object sender, RoutedEventArgs e)
        {
            var w = new AboutWindow();
            w.Owner = this;
            w.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            w.ShowDialog();
            
        }

        private static bool IsDoubleClickOnRow(DependencyObject originalSource)
        {
            if (originalSource == null) return false;

            DependencyObject current = originalSource;

            while (current != null)
            {
                if (current is DataGridRow)
                    return true;

                if (current is DataGridColumnHeader || current is ScrollBar)
                    return false;

                current = GetParentSafe(current);
            }

            return false;
        }

        private static DependencyObject GetParentSafe(DependencyObject child)
        {
            if (child == null) return null;

            // If it's a Visual/Visual3D, use VisualTreeHelper
            if (child is Visual || child is System.Windows.Media.Media3D.Visual3D)
                return VisualTreeHelper.GetParent(child);

            // If it's a FrameworkContentElement (Run, Paragraph, etc.)
            if (child is FrameworkContentElement fce)
                return fce.Parent;

            // Generic ContentElement
            if (child is ContentElement ce)
                return ContentOperations.GetParent(ce);

            return null;
        }


        private ViewSheet ResolveSheetFromRow(IReplaceRow row)
        {
            if (row == null) return null;

            // 1) If the row type has a SheetId property, use it (SheetName + ViewTitles rows)
            var sheetId = TryGetSheetIdViaReflection(row);
            if (sheetId != ElementId.InvalidElementId)
            {
                var sheet = _doc.GetElement(sheetId) as ViewSheet;
                if (sheet != null) return sheet;
            }

            // 2) Fallback: parse from the "Sheet" display string (used by Keynotes results)
            // expected format: "A101 - Sheet Name" or "Not on a Sheet"
            var sheetNumber = ParseSheetNumber(row.Sheet);
            if (string.IsNullOrWhiteSpace(sheetNumber))
                return null;

            var match = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .FirstOrDefault(s => string.Equals(s.SheetNumber, sheetNumber, System.StringComparison.OrdinalIgnoreCase));

            return match;
        }

        private static ElementId TryGetSheetIdViaReflection(IReplaceRow row)
        {
            try
            {
                var prop = row.GetType().GetProperty("SheetId", BindingFlags.Public | BindingFlags.Instance);
                if (prop == null) return ElementId.InvalidElementId;

                var val = prop.GetValue(row, null);
                if (val is ElementId eid) return eid;
            }
            catch
            {
                // ignore
            }

            return ElementId.InvalidElementId;
        }

        private static string ParseSheetNumber(string sheetDisplay)
        {
            if (string.IsNullOrWhiteSpace(sheetDisplay)) return null;
            if (sheetDisplay.StartsWith("Not on a Sheet")) return null;

            // Most of your rows use: "{SheetNumber} - {SheetName}"
            int idx = sheetDisplay.IndexOf(" - ");
            if (idx <= 0) return null;

            return sheetDisplay.Substring(0, idx).Trim();
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
