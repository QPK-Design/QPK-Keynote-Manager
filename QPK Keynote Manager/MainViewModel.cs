using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Text;


namespace QPK_Keynote_Manager
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly UIDocument _uidoc;
        private readonly Document _doc;

        public MainViewModel(UIDocument uidoc)
        {
            _uidoc = uidoc;
            _doc = uidoc.Document;

            // Default checkbox states
            _isKeynotesEnabled = true;
            _isSheetNamesEnabled = false;
            _isViewTitlesEnabled = false;

            RefreshAvailableScopes();
            SelectedScope = AvailableScopes.FirstOrDefault();
        }

        // ---------------- UI State ----------------
        private bool _isCaseSensitive;
        public bool IsCaseSensitive
        {
            get => _isCaseSensitive;
            set { _isCaseSensitive = value; OnPropertyChanged(); }
        }

        private bool _isKeynotesEnabled;
        public bool IsKeynotesEnabled
        {
            get => _isKeynotesEnabled;
            set { _isKeynotesEnabled = value; OnPropertyChanged(); RefreshAvailableScopes(); }
        }

        private bool _isSheetNamesEnabled;
        public bool IsSheetNamesEnabled
        {
            get => _isSheetNamesEnabled;
            set { _isSheetNamesEnabled = value; OnPropertyChanged(); RefreshAvailableScopes(); }
        }

        private bool _isViewTitlesEnabled;
        public bool IsViewTitlesEnabled
        {
            get => _isViewTitlesEnabled;
            set { _isViewTitlesEnabled = value; OnPropertyChanged(); RefreshAvailableScopes(); }
        }

        public ObservableCollection<FindReplaceScope> AvailableScopes { get; } =
            new ObservableCollection<FindReplaceScope>();

        private FindReplaceScope _selectedScope;
        public FindReplaceScope SelectedScope
        {
            get => _selectedScope;
            set
            {
                _selectedScope = value;
                OnPropertyChanged();
                UpdateCurrentResults();
            }
        }

        // ---------------- Search Text ----------------
        private string _findText;
        public string FindText
        {
            get => _findText;
            set { _findText = value; OnPropertyChanged(); }
        }

        private string _replaceText;
        public string ReplaceText
        {
            get => _replaceText;
            set { _replaceText = value; OnPropertyChanged(); }
        }

        // ---------------- Results per-scope ----------------

        public ObservableCollection<ReplaceResult> KeynoteResults { get; } =
            new ObservableCollection<ReplaceResult>();


        public ObservableCollection<SheetNameReplaceRow> SheetNameResults { get; } =
            new ObservableCollection<SheetNameReplaceRow>();

        private ObservableCollection<IReplaceRow> _currentResults = new ObservableCollection<IReplaceRow>();
        public ObservableCollection<IReplaceRow> CurrentResults
        {
            get => _currentResults;
            private set { _currentResults = value; OnPropertyChanged(); }
        }

        private IReplaceRow _selectedResult;
        public IReplaceRow SelectedResult
        {
            get => _selectedResult;
            set { _selectedResult = value; OnPropertyChanged(); }
        }

        private void RefreshAvailableScopes()
        {
            var prevKind = SelectedScope?.Kind;

            AvailableScopes.Clear();

            if (IsKeynotesEnabled) AvailableScopes.Add(new FindReplaceScope(FindReplaceScopeKind.Keynotes, "Keynotes"));
            if (IsSheetNamesEnabled) AvailableScopes.Add(new FindReplaceScope(FindReplaceScopeKind.SheetNames, "Sheet Names"));
            if (IsViewTitlesEnabled) AvailableScopes.Add(new FindReplaceScope(FindReplaceScopeKind.ViewTitles, "View Titles"));

            if (!AvailableScopes.Any())
            {
                // Always keep at least one to avoid null UI state
                AvailableScopes.Add(new FindReplaceScope(FindReplaceScopeKind.Keynotes, "Keynotes"));
                _isKeynotesEnabled = true;
                OnPropertyChanged(nameof(IsKeynotesEnabled));
            }

            SelectedScope = AvailableScopes.FirstOrDefault(s => s.Kind == prevKind) ?? AvailableScopes.FirstOrDefault();
        }

        private void UpdateCurrentResults()
        {
            if (SelectedScope == null) return;

            switch (SelectedScope.Kind)
            {
                case FindReplaceScopeKind.Keynotes:
                    CurrentResults = new ObservableCollection<IReplaceRow>(KeynoteResults.Cast<IReplaceRow>());
                    break;

                case FindReplaceScopeKind.SheetNames:
                    CurrentResults = new ObservableCollection<IReplaceRow>(SheetNameResults.Cast<IReplaceRow>());
                    break;

                default:
                    CurrentResults = new ObservableCollection<IReplaceRow>();
                    break;
            }
        }

        // ---------------- Sheet Name Preview ----------------
        public void PreviewSheetNames()
        {
            SheetNameResults.Clear();

            string search = FindText ?? string.Empty;
            string replace = ReplaceText ?? string.Empty;

            if (string.IsNullOrWhiteSpace(search))
                return;

            var comparison = IsCaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            var sheets = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            foreach (var sheet in sheets)
            {
                string oldName = sheet.Name ?? string.Empty;
                if (oldName.IndexOf(search, comparison) < 0)
                    continue;

                string newName = IsCaseSensitive
                    ? oldName.Replace(search, replace)
                    : ReplaceIgnoreCase(oldName, search, replace);

                newName = (newName ?? string.Empty).Trim();

                if (string.IsNullOrWhiteSpace(newName) || string.Equals(newName, oldName, comparison))
                    continue;

                SheetNameResults.Add(new SheetNameReplaceRow
                {
                    SheetId = sheet.Id,
                    SheetNumber = sheet.SheetNumber,
                    Sheet = $"{sheet.SheetNumber} - {oldName}",
                    FoundText = oldName,
                    ReplacedText = newName
                });
            }

            UpdateCurrentResults();
        }

        public void PreviewKeynotes()
        {
            KeynoteResults.Clear();

            string search = FindText ?? string.Empty;
            string replace = ReplaceText ?? string.Empty;

            if (string.IsNullOrWhiteSpace(search))
                return;

            var seenTypeIds = new HashSet<ElementId>();
            int matchCount = 0;

            var schedules = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTemplate)
                .ToList();

            foreach (var vs in schedules)
            {
                if (!ScheduleShowsCommentField(vs))
                    continue;

                IEnumerable<Element> elems;
                try
                {
                    elems = new FilteredElementCollector(_doc, vs.Id)
                        .WhereElementIsNotElementType()
                        .ToElements();
                }
                catch
                {
                    continue;
                }

                foreach (var elem in elems)
                {
                    if (elem == null) continue;

                    ElementId typeId = elem.GetTypeId();
                    if (typeId == ElementId.InvalidElementId) continue;
                    if (seenTypeIds.Contains(typeId)) continue;

                    var tElem = _doc.GetElement(typeId);
                    if (tElem == null) continue;

                    string oldComment = GetTypeComment(tElem);
                    if (string.IsNullOrWhiteSpace(oldComment)) continue;
                    if (!Contains(oldComment, search, IsCaseSensitive)) continue;

                    string newComment = IsCaseSensitive
                        ? oldComment.Replace(search, replace)
                        : ReplaceIgnoreCase(oldComment, search, replace);

                    if (string.Equals(newComment, oldComment, StringComparison.Ordinal))
                        continue;

                    // --- NEW: type NUMBER ---
                    string typeNumber = GetTypeNumber(tElem);

                    // --- NEW: element-based sheet ---
                    string sheetForElem = GetSheetNameForElement(elem);

                    // --- context strings for green/red UI ---
                    BuildContextStrings(
                        oldComment, search, replace, IsCaseSensitive,
                        out string fPrefix, out string fWord, out string fSuffix,
                        out string rPrefix, out string rWord, out string rSuffix);

                    matchCount++;

                    KeynoteResults.Add(new ReplaceResult
                    {
                        Number = typeNumber,
                        TypeId = typeId,
                        FullOldComment = oldComment,
                        FullNewComment = newComment,

                        FoundPrefix = fPrefix,
                        FoundWord = fWord,
                        FoundSuffix = fSuffix,

                        ReplPrefix = rPrefix,
                        ReplWord = rWord,
                        ReplSuffix = rSuffix,

                        Sheet = sheetForElem,
                        ScheduleName = vs.Name
                    });

                    seenTypeIds.Add(typeId);
                }
            }

            UpdateCurrentResults();
        }


        private static bool Contains(string haystack, string needle, bool caseSensitive)
        {
            if (string.IsNullOrEmpty(haystack) || string.IsNullOrEmpty(needle)) return false;
            return haystack.IndexOf(needle, caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string GetTypeComment(Element tElem)
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
            catch { }
            return false;
        }

        private static string GetTypeNumber(Element tElem)
        {
            if (tElem == null) return string.Empty;

            Parameter numParam =
                tElem.LookupParameter("NUMBER")
                ?? tElem.LookupParameter("Number")
                ?? tElem.LookupParameter("number");

            if (numParam == null) return string.Empty;

            try
            {
                if (numParam.StorageType == StorageType.String)
                    return numParam.AsString() ?? string.Empty;

                return numParam.AsValueString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

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

        private string GetSheetNameForSchedule(ViewSchedule vs)
        {
            try
            {
                var vp = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .FirstOrDefault(x => x.ViewId == vs.Id);

                if (vp == null) return string.Empty;

                var sheet = _doc.GetElement(vp.SheetId) as ViewSheet;
                if (sheet == null) return string.Empty;

                return $"{sheet.SheetNumber} - {sheet.Name}";
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Builds small context around the FIRST word-token match of search (2 tokens before/after).
        /// If no clean token match is found, falls back to entire string in prefix.
        /// </summary>
        private static void BuildContextStrings(
            string fullOld,
            string search,
            string replace,
            bool caseSensitive,
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

            var cmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            int hitIndex = -1;
            for (int i = 0; i < tokens.Length; i++)
            {
                string core = tokens[i].Trim(',', '.', ';', ':', '!', '?');
                if (string.Equals(core, search, cmp))
                {
                    hitIndex = i;
                    break;
                }
            }

            if (hitIndex == -1)
            {
                // fallback: show whole string; color word won’t be perfect, but preview still works
                foundPrefix = fullOld;
                replPrefix = caseSensitive ? fullOld.Replace(search, replace) : ReplaceIgnoreCase(fullOld, search, replace);
                replWord = string.Empty;
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
            replWord = replace;          // this is what will be red
            replSuffix = foundSuffix;
        }



        private static string ReplaceIgnoreCase(string input, string oldValue, string newValue)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(oldValue)) return input;

            int idx = input.IndexOf(oldValue, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return input;

            var sb = new System.Text.StringBuilder();
            int start = 0;
            while (idx >= 0)
            {
                sb.Append(input, start, idx - start);
                sb.Append(newValue);
                start = idx + oldValue.Length;
                idx = input.IndexOf(oldValue, start, StringComparison.OrdinalIgnoreCase);
            }
            sb.Append(input, start, input.Length - start);
            return sb.ToString();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
