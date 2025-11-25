using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QPK_Keynote_Manager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly UIDocument _uidoc;
        private readonly Document _doc;

        private readonly ExternalEvent _replaceAllEvent;
        private readonly ExternalEvent _replaceSelectedEvent;

        public ObservableCollection<ReplaceResult> ReplaceResults { get; }
            = new ObservableCollection<ReplaceResult>();

        // Bound to the Find / Replace TextBoxes (optional, since we read from TextBoxes directly)
        public string FindText { get; set; }
        public string ReplaceText { get; set; }

        public MainWindow(UIDocument uidoc)
        {
            _uidoc = uidoc;
            _doc = uidoc.Document;

            InitializeComponent();
            DataContext = this;   // enables bindings to FindText / ReplaceText / ReplaceResults

            // Create ExternalEvents for modeless operations
            _replaceAllEvent = ExternalEvent.Create(new ReplaceAllHandler(this));
            _replaceSelectedEvent = ExternalEvent.Create(new ReplaceSelectedHandler(this));
        }

        #region Data Model

        public class ReplaceResult
        {
            public ElementId TypeId { get; set; }      // For applying changes
            public string StringFound { get; set; }    // Original type comment
            public string StringReplaced { get; set; } // New type comment
            public string Sheet { get; set; }          // Sheet name/number (if on sheet)
            public string ScheduleName { get; set; }   // Schedule that produced this row
        }

        #endregion

        #region Helpers – Type Comments

        private string GetTypeComment(Element tElem)
        {
            if (tElem == null) return string.Empty;

            // Try built-in "Type Comments"
            Parameter p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null)
            {
                string s = p.AsString();
                if (!string.IsNullOrEmpty(s)) return s;
            }

            // Fallback common names
            string[] names = { "Type Comments", "Comments", "Comment", "COMMENT" };
            foreach (string nm in names)
            {
                p = tElem.LookupParameter(nm);
                if (p != null)
                {
                    string s = p.AsString();
                    if (!string.IsNullOrEmpty(s)) return s;
                }
            }

            return string.Empty;
        }

        internal bool SetTypeComment(Element tElem, string newText)
        {
            if (tElem == null) return false;

            // Prefer built-in
            Parameter p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null && !p.IsReadOnly)
                return p.Set(newText);

            // Fallback common names
            string[] names = { "Type Comments", "Comments", "Comment", "COMMENT" };
            foreach (string nm in names)
            {
                p = tElem.LookupParameter(nm);
                if (p != null && !p.IsReadOnly)
                    return p.Set(newText);
            }

            return false;
        }

        #endregion

        #region Helpers – Schedule / Elements

        private IEnumerable<ViewSchedule> CollectSchedules(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTemplate);
        }

        private bool ScheduleShowsCommentField(ViewSchedule vs)
        {
            try
            {
                var def = vs.Definition;
                for (int i = 0; i < def.GetFieldCount(); i++)
                {
                    var f = def.GetField(i);
                    string name = f.GetName();
                    if (!string.IsNullOrEmpty(name) &&
                        (name.Equals("Comment", StringComparison.OrdinalIgnoreCase) ||
                         name.Equals("Comments", StringComparison.OrdinalIgnoreCase)))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // ignore and return false
            }
            return false;
        }

        private string GetSheetNameForSchedule(ViewSchedule vs)
        {
            // If schedule is placed on one or more sheets, return "S101 - My Sheet" for the first one
            try
            {
                var viewports = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .Where(vp => vp.ViewId == vs.Id)
                    .ToList();

                if (!viewports.Any())
                    return string.Empty;

                var sheet = _doc.GetElement(viewports[0].SheetId) as ViewSheet;
                if (sheet == null) return string.Empty;

                return $"{sheet.SheetNumber} - {sheet.Name}";
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Get elements that appear in this schedule. Using a view-specific collector.
        /// </summary>
        private IEnumerable<Element> GetElementsForSchedule(ViewSchedule vs)
        {
            try
            {
                return new FilteredElementCollector(_doc, vs.Id)
                    .WhereElementIsNotElementType()
                    .ToElements();
            }
            catch
            {
                return Enumerable.Empty<Element>();
            }
        }

        #endregion

        #region Exposed helpers for handlers

        internal ReplaceResult GetSelectedResult()
        {
            return ResultsDataGrid.SelectedItem as ReplaceResult;
        }

        #endregion

        #region UI Handlers

        // PREVIEW – scan schedules, populate DataGrid with old/new values + sheet
        private void PreviewFindReplace_Click(object sender, RoutedEventArgs e)
        {
            string search = FindTextBox.Text ?? string.Empty;
            string replace = ReplaceTextBox.Text ?? string.Empty;

            if (string.IsNullOrWhiteSpace(search))
            {
                MessageBox.Show("Enter text to find.", "Find & Replace",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            ReplaceResults.Clear();

            var seenTypeIds = new HashSet<ElementId>();
            int scheduleCount = 0;
            int matchCount = 0;

            foreach (var vs in CollectSchedules(_doc))
            {
                if (!ScheduleShowsCommentField(vs))
                    continue;

                scheduleCount++;
                string sheetName = GetSheetNameForSchedule(vs);

                IEnumerable<Element> elements = GetElementsForSchedule(vs);

                foreach (var elem in elements)
                {
                    if (elem == null) continue;

                    ElementId typeId = elem.GetTypeId();
                    if (typeId == null || typeId == ElementId.InvalidElementId)
                        continue;

                    // Only preview each Type once
                    if (seenTypeIds.Contains(typeId))
                        continue;

                    var tElem = _doc.GetElement(typeId);
                    string oldComment = GetTypeComment(tElem);
                    if (string.IsNullOrEmpty(oldComment))
                        continue;

                    if (!oldComment.Contains(search))
                        continue;

                    string newComment = oldComment.Replace(search, replace);
                    if (newComment == oldComment)
                        continue;

                    seenTypeIds.Add(typeId);
                    matchCount++;

                    ReplaceResults.Add(new ReplaceResult
                    {
                        TypeId = typeId,
                        StringFound = oldComment,
                        StringReplaced = newComment,
                        Sheet = sheetName,
                        ScheduleName = vs.Name
                    });
                }
            }

            MessageBox.Show(
                $"Scanned {scheduleCount} schedule(s).\nFound {matchCount} type(s) with matching comments.",
                "Preview Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        // These now just raise ExternalEvents – no Transactions here.
        private void ReplaceAll_Click(object sender, RoutedEventArgs e)
        {
            _replaceAllEvent.Raise();
        }

        private void ReplaceSelected_Click(object sender, RoutedEventArgs e)
        {
            _replaceSelectedEvent.Raise();
        }

        #endregion
    }

    #region External Event Handlers

    /// <summary>
    /// Handles "Replace All" inside Revit's API context.
    /// </summary>
    public class ReplaceAllHandler : IExternalEventHandler
    {
        private readonly MainWindow _window;

        public ReplaceAllHandler(MainWindow window)
        {
            _window = window;
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
            {
                TaskDialog.Show("QPK Keynote Manager",
                    "No active document.");
                return;
            }

            Document doc = uidoc.Document;
            var results = _window.ReplaceResults;
            if (results == null || !results.Any())
            {
                TaskDialog.Show("QPK Keynote Manager",
                    "No preview results to apply. Click Preview first.");
                return;
            }

            using (var tx = new Transaction(doc, "Replace Type Comments (All)"))
            {
                tx.Start();

                int changed = 0;
                foreach (var r in results)
                {
                    var tElem = doc.GetElement(r.TypeId);
                    if (tElem == null) continue;

                    if (_window.SetTypeComment(tElem, r.StringReplaced))
                        changed++;
                }

                tx.Commit();

                TaskDialog.Show("QPK Keynote Manager",
                    $"Updated type comments on {changed} type(s).");
            }
        }

        public string GetName()
        {
            return "QPK Keynote Manager – Replace All";
        }
    }

    /// <summary>
    /// Handles "Replace Selected" inside Revit's API context.
    /// </summary>
    public class ReplaceSelectedHandler : IExternalEventHandler
    {
        private readonly MainWindow _window;

        public ReplaceSelectedHandler(MainWindow window)
        {
            _window = window;
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
            {
                TaskDialog.Show("QPK Keynote Manager",
                    "No active document.");
                return;
            }

            Document doc = uidoc.Document;

            var selected = _window.GetSelectedResult();
            if (selected == null)
            {
                TaskDialog.Show("QPK Keynote Manager",
                    "Select a row in the results grid first.");
                return;
            }

            using (var tx = new Transaction(doc, "Replace Type Comment (Selected)"))
            {
                tx.Start();

                var tElem = doc.GetElement(selected.TypeId);
                bool ok = _window.SetTypeComment(tElem, selected.StringReplaced);

                if (ok)
                {
                    tx.Commit();
                    TaskDialog.Show(
                        "QPK Keynote Manager",
                        $"Updated type comment for TypeId {selected.TypeId.IntegerValue}.");
                }
                else
                {
                    tx.RollBack();
                    TaskDialog.Show(
                        "QPK Keynote Manager",
                        "Failed to set type comment (parameter may be read-only).");
                }
            }
        }

        public string GetName()
        {
            return "QPK Keynote Manager – Replace Selected";
        }
    }

    #endregion
}
