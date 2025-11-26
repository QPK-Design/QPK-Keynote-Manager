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

        // Optional backing for bindings (we still read directly from TextBoxes)
        public string FindText { get; set; }
        public string ReplaceText { get; set; }

        public MainWindow(UIDocument uidoc)
        {
            _uidoc = uidoc;
            _doc = uidoc.Document;

            InitializeComponent();
            DataContext = this;

            _replaceAllEvent = ExternalEvent.Create(new ReplaceAllHandler(this));
            _replaceSelectedEvent = ExternalEvent.Create(new ReplaceSelectedHandler(this));
        }

        #region Data Model

        public class ReplaceResult
        {
            public ElementId TypeId { get; set; }

            // Full comments used for actual Revit parameter updates
            public string FullOldComment { get; set; }
            public string FullNewComment { get; set; }

            // Context around the first match in the comment
            public string FoundPrefix { get; set; }   // words before target
            public string FoundWord { get; set; }     // target word (e.g. "hello")
            public string FoundSuffix { get; set; }   // words after target

            public string ReplPrefix { get; set; }
            public string ReplWord { get; set; }      // replacement word (e.g. "this")
            public string ReplSuffix { get; set; }

            public string Sheet { get; set; }
            public string ScheduleName { get; set; }

            // Handy combined strings (not used for coloring, but available)
            public string StringFound => string.Concat(FoundPrefix, FoundWord, FoundSuffix);
            public string StringReplaced => string.Concat(ReplPrefix, ReplWord, ReplSuffix);
        }

        #endregion

        #region Helpers – Type Comments

        private string GetTypeComment(Element tElem)
        {
            if (tElem == null) return string.Empty;

            Parameter p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null)
            {
                string s = p.AsString();
                if (!string.IsNullOrEmpty(s)) return s;
            }

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

            Parameter p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null && !p.IsReadOnly)
                return p.Set(newText);

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
            }
            return false;
        }

        private string GetSheetNameForSchedule(ViewSchedule vs)
        {
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

        /// <summary>
        /// Build the small context string around the first occurrence of <paramref name="search"/>.
        /// Returns pieces for the "found" and "replaced" context strings.
        /// </summary>
        private void BuildContextStrings(
            string fullOld,
            string search,
            string replace,
            out string foundPrefix,
            out string foundWord,
            out string foundSuffix,
            out string replPrefix,
            out string replWord,
            out string replSuffix)
        {
            foundPrefix = foundWord = foundSuffix = string.Empty;
            replPrefix = replWord = replSuffix = string.Empty;

            if (string.IsNullOrWhiteSpace(fullOld) || string.IsNullOrEmpty(search))
                return;

            var tokens = fullOld
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .ToArray();

            if (tokens.Length == 0) return;

            int hitIndex = -1;
            for (int i = 0; i < tokens.Length; i++)
            {
                // Strip simple punctuation when checking for equality
                string core = tokens[i].Trim(',', '.', ';', ':', '!', '?');
                if (string.Equals(core, search, StringComparison.Ordinal))
                {
                    hitIndex = i;
                    break;
                }
            }

            if (hitIndex == -1)
            {
                // Fallback: we didn't find a clean word match, just bail to entire sentence.
                foundPrefix = fullOld;
                replPrefix = fullOld.Replace(search, replace);
                return;
            }

            int start = Math.Max(hitIndex - 2, 0);
            int end = Math.Min(hitIndex + 2, tokens.Length - 1);

            var beforeTokens = tokens.Skip(start).Take(hitIndex - start).ToArray();
            var afterTokens = tokens.Skip(hitIndex + 1).Take(end - hitIndex).ToArray();

            foundPrefix = beforeTokens.Length > 0 ? string.Join(" ", beforeTokens) + " " : string.Empty;
            foundWord = tokens[hitIndex];
            foundSuffix = afterTokens.Length > 0 ? " " + string.Join(" ", afterTokens) : string.Empty;

            replPrefix = foundPrefix;
            replWord = replace;
            replSuffix = foundSuffix;
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

                    // Build context strings for display
                    BuildContextStrings(
                        oldComment,
                        search,
                        replace,
                        out string fPrefix,
                        out string fWord,
                        out string fSuffix,
                        out string rPrefix,
                        out string rWord,
                        out string rSuffix);

                    seenTypeIds.Add(typeId);
                    matchCount++;

                    ReplaceResults.Add(new ReplaceResult
                    {
                        TypeId = typeId,
                        FullOldComment = oldComment,
                        FullNewComment = newComment,
                        FoundPrefix = fPrefix,
                        FoundWord = fWord,
                        FoundSuffix = fSuffix,
                        ReplPrefix = rPrefix,
                        ReplWord = rWord,
                        ReplSuffix = rSuffix,
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

                    if (_window.SetTypeComment(tElem, r.FullNewComment))
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
                bool ok = _window.SetTypeComment(tElem, selected.FullNewComment);

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
