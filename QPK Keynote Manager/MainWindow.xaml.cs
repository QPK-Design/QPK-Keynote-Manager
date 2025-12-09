using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QPK_Keynote_Manager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// Future Functions:
    /// - Case Sensitive Toggle
    /// - Add text note search
    /// - Add sheet name search
    /// - Add view name search
    /// - Make column for "schedule vs text vs sheet name vs view name"
    /// - Add spell checking function globally
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
        private bool isCaseSensitive = false; // Default off


        public MainWindow(UIDocument uidoc)
        {
            _uidoc = uidoc;
            _doc = uidoc.Document;

            InitializeComponent();
            DataContext = this;

            _replaceAllEvent = ExternalEvent.Create(new ReplaceAllHandler(this));
            _replaceSelectedEvent = ExternalEvent.Create(new ReplaceSelectedHandler(this));

            caseSensitiveToggle.Checked += CaseSensitiveToggle_Checked;
            caseSensitiveToggle.Unchecked += CaseSensitiveToggle_Checked;
        }

        #region Data Model

        public class ReplaceResult : INotifyPropertyChanged
        {
            public ElementId TypeId { get; set; }

            public string Number { get; set; }

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

            // 🔹 add this (see next section)
            public string ScheduleName { get; set; }

            private bool _isApplied;
            public bool IsApplied
            {
                get => _isApplied;
                set
                {
                    if (_isApplied != value)
                    {
                        _isApplied = value;
                        OnPropertyChanged();
                    }
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
            private void OnPropertyChanged([CallerMemberName] string name = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        #endregion


        #region Helpers – Type Comments

        private string GetSheetNameForElement(Element elem)
        {
            if (elem == null)
                return "Not on a Sheet";

            ElementId ownerViewId = elem.OwnerViewId;
            if (ownerViewId == null || ownerViewId == ElementId.InvalidElementId)
                return "Not on a Sheet";

            var ownerView = _doc.GetElement(ownerViewId) as View;
            if (ownerView == null)
                return "Not on a Sheet";

            // Find the viewport that places this view on a sheet
            var viewport = new FilteredElementCollector(_doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .FirstOrDefault(vp => vp.ViewId == ownerView.Id);

            if (viewport == null)
                return "Not on a Sheet";

            var sheet = _doc.GetElement(viewport.SheetId) as ViewSheet;
            if (sheet == null)
                return "Not on a Sheet";

            return $"{sheet.SheetNumber} - {sheet.Name}";
        }

        private void CaseSensitiveToggle_Checked(object sender, RoutedEventArgs e)
        {
            // Store the state for use in search functions
            isCaseSensitive = caseSensitiveToggle.IsChecked.GetValueOrDefault();
        }

        private string GetTypeComment(Element tElem, string newText)
        {
            if (tElem == null) return string.Empty;

            Parameter p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null)
            {
                string s = p.AsString();
                if (!string.IsNullOrEmpty(s) && CompareStrings(s, newText, isCaseSensitive))
                    return s;
            }

            string[] names = { "Type Comments", "Comments", "Comment", "COMMENT" };
            foreach (string nm in names)
            {
                p = tElem.LookupParameter(nm);
                if (p != null)
                {
                    string s = p.AsString();
                    if (!string.IsNullOrEmpty(s) && CompareStrings(s, newText, isCaseSensitive))
                        return s;
                }
            }

            return string.Empty;
        }


        private bool CompareStrings(string str1, string str2, bool caseSensitive)
        {
            return caseSensitive ? str1 == str2 : str1.Equals(str2, StringComparison.OrdinalIgnoreCase);
        }

        internal bool SetTypeComment(Element tElem, string newText)
        {
            if (tElem == null) return false;

            Parameter p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null && !p.IsReadOnly)
            {
                // Here you may want to check the case-sensitive logic
                if (!isCaseSensitive || p.AsString().Equals(newText, StringComparison.OrdinalIgnoreCase))
                    return p.Set(newText);
            }

            // Check alternative parameters
            string[] names = { "Type Comments", "Comments", "Comment", "COMMENT" };
            foreach (string nm in names)
            {
                p = tElem.LookupParameter(nm);
                if (p != null && !p.IsReadOnly)
                {
                    // Check for case-sensitivity before setting
                    if (!isCaseSensitive || p.AsString().Equals(newText, StringComparison.OrdinalIgnoreCase))
                        return p.Set(newText);
                }
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
                // Strip simple punctuation when checking for a match
                string token = tokens[i];
                string core = token.Trim(',', '.', ';', ':', '!', '?');

                // Treat any token that *contains* the search text as a hit
                // This will catch DRAIN, DRAINS, DRAIN., etc.
                if (core.IndexOf(search, StringComparison.Ordinal) >= 0)
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
            string matchedToken = tokens[hitIndex];
            foundWord = matchedToken;
            foundSuffix = afterTokens.Length > 0 ? " " + string.Join(" ", afterTokens) : string.Empty;

            replPrefix = foundPrefix;
            // Apply the same replacement logic to just the matched token
            replWord = matchedToken.Replace(search, replace);
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

                    // --- get NUMBER from the TYPE parameter (unchanged) ---
                    string number = string.Empty;
                    Parameter numParam = null;

                    if (tElem != null)
                    {
                        numParam = tElem.LookupParameter("NUMBER")
                               ?? tElem.LookupParameter("Number")
                               ?? tElem.LookupParameter("number");
                    }

                    if (numParam != null)
                    {
                        if (numParam.StorageType == StorageType.String)
                            number = numParam.AsString();
                        else
                            number = numParam.AsValueString();
                    }
                    // --- end NUMBER lookup ---

                    string oldComment = GetTypeComment(tElem, search);
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

                    // NEW: sheet for this keynote instance’s owner view
                    string sheetName = GetSheetNameForElement(elem);

                    ReplaceResults.Add(new ReplaceResult
                    {
                        TypeId = typeId,
                        Number = number,
                        FullOldComment = oldComment,
                        FullNewComment = newComment,
                        FoundPrefix = fPrefix,
                        FoundWord = fWord,
                        FoundSuffix = fSuffix,
                        ReplPrefix = rPrefix,
                        ReplWord = rWord,
                        ReplSuffix = rSuffix,
                        Sheet = sheetName,         // now "Not on a Sheet" or actual sheet
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
                        r.IsApplied = true;  // <--- mark row as applied
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
                    selected.IsApplied = true;   // <--- visually mark the row

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
