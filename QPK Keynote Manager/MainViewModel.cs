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

        public ObservableCollection<ViewTitleNameReplaceRow> ViewTitleResults { get; } = new ObservableCollection<ViewTitleNameReplaceRow>();


        private IReplaceRow _selectedResult;
        public IReplaceRow SelectedResult
        {
            get => _selectedResult;
            set { _selectedResult = value; OnPropertyChanged(); }
        }

        private void RefreshAvailableScopes()
        {
            FindReplaceScopeKind? prevKind = SelectedScope?.Kind;

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

                case FindReplaceScopeKind.ViewTitles:
                    CurrentResults = new ObservableCollection<IReplaceRow>(ViewTitleResults.Cast<IReplaceRow>());
                    break;

                default:
                    CurrentResults = new ObservableCollection<IReplaceRow>();
                    break;
            }
        }



        // ---------------- Sheet Name Preview ----------------
        public void PreviewSheetNames(bool updateCurrentResults = true)
        {
            SheetNameResults.Clear();

            string search = (FindText ?? string.Empty).Trim();
            string replace = ReplaceText ?? string.Empty;

            if (string.IsNullOrWhiteSpace(search))
                return;

            StringComparison comparison = IsCaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            List<ViewSheet> sheets = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            foreach (ViewSheet? sheet in sheets)
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

                // --- context strings for green/red UI (match keynote UX) ---
                BuildWordAwareContextStrings(
                    oldName, search, replace, IsCaseSensitive,
                    out string? fPre, out string? fLeft, out string? fMid, out string? fRight, out string? fPost,
                    out string? rPre, out string? rLeft, out string? rMid, out string? rRight, out string? rPost);

                SheetNameResults.Add(new SheetNameReplaceRow
                {
                    SheetId = sheet.Id,
                    SheetNumber = sheet.SheetNumber,
                    Sheet = $"{sheet.SheetNumber} - {oldName}",

                    FoundText = oldName,
                    ReplacedText = newName,

                    // NEW: Found segments
                    FoundPreText = fPre,
                    FoundWordLeft = fLeft,
                    FoundWordMid = fMid,
                    FoundWordRight = fRight,
                    FoundPostText = fPost,

                    // NEW: Replaced segments
                    ReplPreText = rPre,
                    ReplWordLeft = rLeft,
                    ReplWordMid = rMid,
                    ReplWordRight = rRight,
                    ReplPostText = rPost
                });
            }

            if (updateCurrentResults)
                UpdateCurrentResults();
        }

        public void PreviewKeynotes(bool updateCurrentResults = true)
        {
            KeynoteResults.Clear();

            string search = FindText ?? string.Empty;
            string replace = ReplaceText ?? string.Empty;

            if (string.IsNullOrWhiteSpace(search))
                return;

            HashSet<ElementId> seenTypeIds = new HashSet<ElementId>();
            int matchCount = 0;

            List<ViewSchedule> schedules = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTemplate)
                .ToList();

            foreach (ViewSchedule? vs in schedules)
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

                foreach (Element elem in elems)
                {
                    if (elem == null) continue;

                    ElementId typeId = elem.GetTypeId();
                    if (typeId == ElementId.InvalidElementId) continue;
                    if (seenTypeIds.Contains(typeId)) continue;

                    Element tElem = _doc.GetElement(typeId);
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
                    BuildWordAwareContextStrings(
                        oldComment, search, replace, IsCaseSensitive,
                        out string? fPre, out string? fLeft, out string? fMid, out string? fRight, out string? fPost,
                        out string? rPre, out string? rLeft, out string? rMid, out string? rRight, out string? rPost);

                    matchCount++;

                    KeynoteResults.Add(new ReplaceResult
                    {
                        Number = typeNumber,
                        TypeId = typeId,
                        FullOldComment = oldComment,
                        FullNewComment = newComment,

                        // NEW word-aware bindings
                        FoundPreText = fPre,
                        FoundWordLeft = fLeft,
                        FoundWordMid = fMid,
                        FoundWordRight = fRight,
                        FoundPostText = fPost,

                        ReplPreText = rPre,
                        ReplWordLeft = rLeft,
                        ReplWordMid = rMid,
                        ReplWordRight = rRight,
                        ReplPostText = rPost,

                        Sheet = sheetForElem,
                        ScheduleName = vs.Name
                    });

                    seenTypeIds.Add(typeId);
                }
            }

            if (updateCurrentResults)
                UpdateCurrentResults();
        }

        public void PreviewViewTitles(bool updateCurrentResults = true)
        {
            ViewTitleResults.Clear();

            string search = FindText ?? string.Empty;
            string replace = ReplaceText ?? string.Empty;

            if (string.IsNullOrWhiteSpace(search))
                return;

            StringComparison comparison = IsCaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            // Viewports -> View + Sheet
            List<Viewport> viewports = new FilteredElementCollector(_doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .ToList();

            foreach (Viewport? vp in viewports)
            {
                View? view = _doc.GetElement(vp.ViewId) as View;
                ViewSheet? sheet = _doc.GetElement(vp.SheetId) as ViewSheet;
                if (view == null) continue;

                string sheetLabel = (sheet == null)
                    ? "Not on a Sheet"
                    : $"{sheet.SheetNumber} - {sheet.Name}";

                // ---------------- VN (View.Name) ----------------
                string oldName = view.Name ?? string.Empty;
                if (oldName.IndexOf(search, comparison) >= 0)
                {
                    string newName = IsCaseSensitive
                        ? oldName.Replace(search, replace)
                        : ReplaceIgnoreCase(oldName, search, replace);

                    newName = (newName ?? string.Empty).Trim();

                    if (!string.IsNullOrWhiteSpace(newName) &&
                        !string.Equals(newName, oldName, comparison))
                    {
                        BuildWordAwareContextStrings(
                            oldName, search, replace, IsCaseSensitive,
                            out string? fPre, out string? fLeft, out string? fMid, out string? fRight, out string? fPost,
                            out string? rPre, out string? rLeft, out string? rMid, out string? rRight, out string? rPost);

                        ViewTitleResults.Add(new ViewTitleNameReplaceRow
                        {
                            ViewId = view.Id,
                            SheetId = sheet?.Id ?? ElementId.InvalidElementId,

                            Mode = "VN",
                            ModeToolTip = "View Name",

                            // For UI + IReplaceRow
                            FoundText = oldName,
                            ReplacedText = newName,

                            // Keep for apply (your handlers use FullNewText)
                            FullOldText = oldName,
                            FullNewText = newName,

                            // Found segments
                            FoundPreText = fPre,
                            FoundWordLeft = fLeft,
                            FoundWordMid = fMid,
                            FoundWordRight = fRight,
                            FoundPostText = fPost,

                            // Replaced segments
                            ReplPreText = rPre,
                            ReplWordLeft = rLeft,
                            ReplWordMid = rMid,
                            ReplWordRight = rRight,
                            ReplPostText = rPost,

                            Sheet = sheetLabel
                        });
                    }
                }

                // ---------------- VT (Title on Sheet) ----------------
                // BuiltInParameter.VIEW_DESCRIPTION is "Title on Sheet"
                string paramTitle = GetTitleOnSheet(view); // may be blank
                string effectiveTitle = !string.IsNullOrWhiteSpace(paramTitle) ? paramTitle : oldName;

                if (!string.IsNullOrEmpty(effectiveTitle) &&
                    effectiveTitle.IndexOf(search, comparison) >= 0)
                {
                    string newTitle = IsCaseSensitive
                        ? effectiveTitle.Replace(search, replace)
                        : ReplaceIgnoreCase(effectiveTitle, search, replace);

                    newTitle = (newTitle ?? string.Empty).Trim();

                    if (!string.IsNullOrWhiteSpace(newTitle) &&
                        !string.Equals(newTitle, effectiveTitle, comparison))
                    {
                        BuildWordAwareContextStrings(
                            effectiveTitle, search, replace, IsCaseSensitive,
                            out string? fPre, out string? fLeft, out string? fMid, out string? fRight, out string? fPost,
                            out string? rPre, out string? rLeft, out string? rMid, out string? rRight, out string? rPost);

                        ViewTitleResults.Add(new ViewTitleNameReplaceRow
                        {
                            ViewId = view.Id,
                            SheetId = sheet?.Id ?? ElementId.InvalidElementId,

                            Mode = "VT",
                            ModeToolTip = "View Title",

                            FoundText = effectiveTitle,
                            ReplacedText = newTitle,

                            FullOldText = effectiveTitle,
                            FullNewText = newTitle,

                            FoundPreText = fPre,
                            FoundWordLeft = fLeft,
                            FoundWordMid = fMid,
                            FoundWordRight = fRight,
                            FoundPostText = fPost,

                            ReplPreText = rPre,
                            ReplWordLeft = rLeft,
                            ReplWordMid = rMid,
                            ReplWordRight = rRight,
                            ReplPostText = rPost,

                            Sheet = sheetLabel
                        });
                    }
                }

            }
            if (updateCurrentResults)
                UpdateCurrentResults();
        }

        private static string GetTitleOnSheet(View view)
        {
            if (view == null) return string.Empty;

            Parameter p = view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION);
            if (p == null) return string.Empty;

            try { return p.AsString() ?? string.Empty; }
            catch { return string.Empty; }
        }

        /// <summary>
        /// Substring-based highlight: prefix + (matched search) + suffix.
        /// This matches your green/red UI requirement more precisely than the token-based keynote helper.
        /// </summary>
        private static void BuildSubstringContextStrings(
            string fullOld,
            string search,
            string replace,
            bool caseSensitive,
            out string? foundPrefix,
            out string? foundWord,
            out string? foundSuffix,
            out string? replPrefix,
            out string? replWord,
            out string? replSuffix)
        {
            foundPrefix = foundWord = foundSuffix = string.Empty;
            replPrefix = replWord = replSuffix = string.Empty;

            if (string.IsNullOrEmpty(fullOld) || string.IsNullOrEmpty(search))
                return;

            StringComparison cmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            int idx = fullOld.IndexOf(search, cmp);
            if (idx < 0)
            {
                // fallback: show entire string if something unexpected happens
                foundPrefix = fullOld;
                replPrefix = caseSensitive ? fullOld.Replace(search, replace) : ReplaceIgnoreCase(fullOld, search, replace);
                return;
            }

            foundPrefix = fullOld.Substring(0, idx);
            foundWord = fullOld.Substring(idx, search.Length);
            foundSuffix = fullOld.Substring(idx + search.Length);

            replPrefix = foundPrefix;
            replWord = replace; // red
            replSuffix = foundSuffix;
        }



        private static bool IsWordChar(char c) =>
            // Treat letters/digits/_ as word characters
            // (This works great for LIGHTING, DOORS, etc.)
            char.IsLetterOrDigit(c) || c == '_';

        private void BuildWordAwareContextStrings(
            string? original,
            string search,
            string replace,
            bool caseSensitive,
            out string? fPre, out string? fLeft, out string? fMid, out string? fRight, out string? fPost,
            out string? rPre, out string? rLeft, out string? rMid, out string? rRight, out string? rPost)
        {
            fPre = original ?? "";
            fLeft = fMid = fRight = fPost = "";

            rPre = original ?? "";
            rLeft = rMid = rRight = rPost = "";

            if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(search))
                return;

            StringComparison comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            int idx = original.IndexOf(search, comparison);
            if (idx < 0)
                return;

            int matchStart = idx;
            int matchEnd = idx + search.Length; // exclusive

            // Find word boundaries around the match
            int wordStart = matchStart;
            while (wordStart > 0 && IsWordChar(original[wordStart - 1]))
                wordStart--;

            int wordEnd = matchEnd;
            while (wordEnd < original.Length && IsWordChar(original[wordEnd]))
                wordEnd++;

            // Found segments
            fPre = original.Substring(0, wordStart);
            fLeft = original.Substring(wordStart, matchStart - wordStart);
            fMid = original.Substring(matchStart, search.Length);
            fRight = original.Substring(matchEnd, wordEnd - matchEnd);
            fPost = original.Substring(wordEnd);

            // Build replaced string (single replacement at first match, consistent with highlighting)
            string repl = original.Remove(matchStart, search.Length).Insert(matchStart, replace ?? "");

            // Replaced segments:
            // wordStart stays the same, matchStart stays the same, but mid length changes to replace.Length
            int replMidLen = (replace ?? "").Length;
            int replMatchEnd = matchStart + replMidLen;

            // Word end shifts by delta in length
            int delta = replMidLen - search.Length;
            int replWordEnd = wordEnd + delta;

            rPre = repl.Substring(0, wordStart);
            rLeft = repl.Substring(wordStart, matchStart - wordStart);
            rMid = repl.Substring(matchStart, replMidLen);
            rRight = repl.Substring(replMatchEnd, replWordEnd - replMatchEnd);
            rPost = repl.Substring(replWordEnd);
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
                ScheduleDefinition def = vs.Definition;
                for (int i = 0; i < def.GetFieldCount(); i++)
                {
                    ScheduleField f = def.GetField(i);
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

            View? ownerView = _doc.GetElement(ownerViewId) as View;
            if (ownerView == null)
                return "Not on a Sheet";

            Viewport viewport = new FilteredElementCollector(_doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .FirstOrDefault(vp => vp.ViewId == ownerView.Id);

            if (viewport == null)
                return "Not on a Sheet";

            ViewSheet? sheet = _doc.GetElement(viewport.SheetId) as ViewSheet;
            if (sheet == null)
                return "Not on a Sheet";

            return $"{sheet.SheetNumber} - {sheet.Name}";
        }

        private string GetSheetNameForSchedule(ViewSchedule vs)
        {
            try
            {
                Viewport vp = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .FirstOrDefault(x => x.ViewId == vs.Id);

                if (vp == null) return string.Empty;

                ViewSheet? sheet = _doc.GetElement(vp.SheetId) as ViewSheet;
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
            out string? foundPrefix,
            out string? foundWord,
            out string? foundSuffix,
            out string? replPrefix,
            out string? replWord,
            out string? replSuffix)
        {
            foundPrefix = foundWord = foundSuffix = string.Empty;
            replPrefix = replWord = replSuffix = string.Empty;

            if (string.IsNullOrWhiteSpace(fullOld) || string.IsNullOrEmpty(search))
                return;

            string[] tokens = fullOld
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .ToArray();

            if (tokens.Length == 0) return;

            StringComparison cmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

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

            string[] beforeTokens = tokens.Skip(start).Take(hitIndex - start).ToArray();
            string[] afterTokens = tokens.Skip(hitIndex + 1).Take(end - hitIndex).ToArray();

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

            StringBuilder sb = new System.Text.StringBuilder();
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

        public IEnumerable<FindReplaceScope> GetAvailableScopes()
        {
            return AvailableScopes;
        }

        public void PreviewEnabledScopes(IEnumerable<FindReplaceScope> availableScopes)
        {
            // Remember what the user is currently viewing in the dropdown
            FindReplaceScopeKind? keepSelected = SelectedScope?.Kind;

            // Run scans based on checkboxes
            if (IsKeynotesEnabled)
                PreviewKeynotes(updateCurrentResults: false);

            if (IsSheetNamesEnabled)
                PreviewSheetNames(updateCurrentResults: false);

            if (IsViewTitlesEnabled)
                PreviewViewTitles(updateCurrentResults: false);

            // Restore whatever the user was viewing
            if (keepSelected != null)
                SelectedScope = availableScopes.FirstOrDefault(s => s.Kind == keepSelected) ?? SelectedScope;

            // Finally, refresh the grid for the dropdown-selected scope
            UpdateCurrentResults();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
