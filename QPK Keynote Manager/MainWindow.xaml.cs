using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
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

        // Collection bound to the DataGrid
        public ObservableCollection<ReplaceResult> ReplaceResults { get; }
            = new ObservableCollection<ReplaceResult>();

        // Bound to the Find / Replace TextBoxes (optional; we also read from the boxes directly)
        public string FindText { get; set; }
        public string ReplaceText { get; set; }

        public MainWindow(UIDocument uidoc)
        {
            _uidoc = uidoc;
            _doc = uidoc.Document;

            InitializeComponent();

            // Bind this instance as DataContext so bindings work:
            //   - FindText, ReplaceText (if you use them)
            //   - ReplaceResults for the DataGrid
            DataContext = this;
        }

        // Model for each row in the DataGrid
        public class ReplaceResult
        {
            public ElementId TypeId { get; set; }      // For applying changes
            public string StringFound { get; set; }    // Original type comment
            public string StringReplaced { get; set; } // New type comment
            public string Sheet { get; set; }          // Sheet number/name (if schedule is placed)
            public string ScheduleName { get; set; }   // Schedule this came from (optional)
        }

        #region Helpers: get/set type comments

        private string GetTypeComment(Element tElem)
        {
            if (tElem == null) return string.Empty;

            // Try built-in "Type Comments"
            var p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null)
            {
                var s = p.AsString();
                if (!string.IsNullOrEmpty(s)) return s;
            }

            // Fallback common names
            string[] names = { "Type Comments", "Comments", "Comment", "COMMENT" };
            foreach (var nm in names)
            {
                p = tElem.LookupParameter(nm);
                if (p != null)
                {
                    var s = p.AsString();
                    if (!string.IsNullOrEmpty(s)) return s;
                }
            }

            return string.Empty;
        }

        private bool SetTypeComment(Element tElem, string newText)
        {
            if (tElem == null) return false;

            // Prefer built-in
            var p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null && !p.IsReadOnly)
                return p.Set(newText);

            // Fallback common names
            string[] names = { "Type Comments", "Comments", "Comment", "COMMENT" };
            foreach (var nm in names)
            {
                p = tElem.LookupParameter(nm);
                if (p != null && !p.IsReadOnly)
                    return p.Set(newText);
            }

            return false;
        }

        #endregion

        #region Helpers: schedules & sheet name

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
                // ignore and treat as false
            }

            return false;
        }

        private string GetSheetNameForSchedule(ViewSchedule vs)
        {
            // If schedule placed on sheet(s), return "S101 - My Sheet", etc.
            try
            {
                var viewports = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .Where(vp => vp.ViewId == vs.Id)
                    .ToList();

                if (!viewports.Any())
                    return string.Empty;

                // Take the first sheet for display
                var sheet = _doc.GetElement(viewports[0].SheetId) as ViewSheet;
                if (sheet == null) return string.Empty;

                return $"{sheet.SheetNumber} - {sheet.Name}";
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion

        #region Button Handlers

        private IEnumerable<Element> GetElementsForSchedule(ViewSchedule vs)
        {
            if (vs == null)
                return Enumerable.Empty<Element>();

            var def = vs.Definition;
            if (def == null)
                return Enumerable.Empty<Element>();

            ElementId catId = def.CategoryId;
            if (catId == null || catId == ElementId.InvalidElementId)
                return Enumerable.Empty<Element>();

            // Collect all non-type elements in the same category as the schedule
            var collector = new FilteredElementCollector(_doc)
                .OfCategoryId(catId)
                .WhereElementIsNotElementType();

            return collector.Cast<Element>();
        }

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

                // 🚫 OLD (doesn't exist in your API):
                // FilteredElementCollector coll;
                // try
                // {
                //     coll = vs.GetFilteredElementCollector();
                // }
                // catch
                // {
                //     coll = new FilteredElementCollector(_doc, vs.Id)
                //         .WhereElementIsNotElementType();
                // }

                // ✅ NEW: use our helper
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


        // REPLACE ALL – apply all previewed replacements
        private void ReplaceAll_Click(object sender, RoutedEventArgs e)
        {
            if (!ReplaceResults.Any())
            {
                MessageBox.Show("No preview results to apply. Click Preview first.",
                    "Replace All", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            using (var tx = new Transaction(_doc, "Replace Type Comments (All)"))
            {
                tx.Start();

                int changed = 0;
                foreach (var r in ReplaceResults)
                {
                    var tElem = _doc.GetElement(r.TypeId);
                    if (tElem == null) continue;

                    if (SetTypeComment(tElem, r.StringReplaced))
                        changed++;
                }

                tx.Commit();

                MessageBox.Show(
                    $"Updated type comments on {changed} type(s).",
                    "Replace All",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }

        // REPLACE SELECTED – only apply for the currently selected row
        private void ReplaceSelected_Click(object sender, RoutedEventArgs e)
        {
            var selected = ResultsDataGrid.SelectedItem as ReplaceResult;
            if (selected == null)
            {
                MessageBox.Show("Select a row in the results grid first.",
                    "Replace Selected", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            using (var tx = new Transaction(_doc, "Replace Type Comment (Selected)"))
            {
                tx.Start();

                var tElem = _doc.GetElement(selected.TypeId);
                bool ok = SetTypeComment(tElem, selected.StringReplaced);

                tx.Commit();

                if (ok)
                {
                    MessageBox.Show(
                        $"Updated type comment for TypeId {selected.TypeId.IntegerValue}.",
                        "Replace Selected",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show(
                        "Failed to set type comment (parameter may be read-only).",
                        "Replace Selected",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
        }

        #endregion
    }
}
