using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Documents;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Globalization;

namespace QPK_Keynote_Manager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// Features:
    /// - Find & Replace with case sensitivity
    /// - Spell Check with custom dictionary
    /// - Accept/Ignore spelling suggestions
    /// - Per-project custom word storage
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private readonly UIDocument _uidoc;
        private readonly Document _doc;

        private readonly ExternalEvent _replaceAllEvent;
        private readonly ExternalEvent _replaceSelectedEvent;
        private readonly ExternalEvent _applySpellingEvent;

        public ObservableCollection<ReplaceResult> ReplaceResults { get; }
            = new ObservableCollection<ReplaceResult>();

        public ObservableCollection<SpellingResult> SpellingResults { get; }
            = new ObservableCollection<SpellingResult>();

        private HashSet<string> _customDictionary = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private HashSet<string> _englishDictionary = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public string FindText { get; set; }
        public string ReplaceText { get; set; }
        private bool isCaseSensitive = false;

        // Spell check mode tracking
        private bool _isSpellCheckMode = false;
        public bool IsSpellCheckMode
        {
            get => _isSpellCheckMode;
            set
            {
                if (_isSpellCheckMode != value)
                {
                    _isSpellCheckMode = value;
                    OnPropertyChanged();
                    UpdateUIMode();
                }
            }
        }

        private int _currentSpellingIndex = -1;
        public int CurrentSpellingIndex
        {
            get => _currentSpellingIndex;
            set
            {
                if (_currentSpellingIndex != value)
                {
                    _currentSpellingIndex = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(CanNavigateNext));
                    OnPropertyChanged(nameof(CanNavigatePrevious));
                    OnPropertyChanged(nameof(CurrentSpellingDisplayNumber));
                }
            }
        }

        public int CurrentSpellingDisplayNumber => CurrentSpellingIndex + 1;

        public bool CanNavigateNext => CurrentSpellingIndex < SpellingResults.Count - 1;
        public bool CanNavigatePrevious => CurrentSpellingIndex > 0;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        public MainWindow(UIDocument uidoc)
        {
            _uidoc = uidoc;
            _doc = uidoc.Document;

            InitializeComponent();
            DataContext = this;

            _replaceAllEvent = ExternalEvent.Create(new ReplaceAllHandler(this));
            _replaceSelectedEvent = ExternalEvent.Create(new ReplaceSelectedHandler(this));
            _applySpellingEvent = ExternalEvent.Create(new ApplySpellingHandler(this));

            caseSensitiveToggle.Checked += CaseSensitiveToggle_Checked;
            caseSensitiveToggle.Unchecked += CaseSensitiveToggle_Checked;

            LoadCustomDictionary();
            LoadEnglishDictionary();
        }

        #region Data Model

        public class ReplaceResult : INotifyPropertyChanged
        {
            public ElementId TypeId { get; set; }
            public string Number { get; set; }
            public string FullOldComment { get; set; }
            public string FullNewComment { get; set; }
            public string FoundPrefix { get; set; }
            public string FoundWord { get; set; }
            public string FoundSuffix { get; set; }
            public string ReplPrefix { get; set; }
            public string ReplWord { get; set; }
            public string ReplSuffix { get; set; }
            public string Sheet { get; set; }
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

        public class SpellingResult : INotifyPropertyChanged
        {
            public ElementId TypeId { get; set; }
            public string Number { get; set; }
            public string FullComment { get; set; }
            public string FullCorrectedComment { get; set; }
            public string MisspelledWord { get; set; }
            public ObservableCollection<string> Suggestions { get; set; } = new ObservableCollection<string>();

            private string _selectedSuggestion;
            public string SelectedSuggestion
            {
                get => _selectedSuggestion;
                set
                {
                    if (_selectedSuggestion != value)
                    {
                        _selectedSuggestion = value;
                        OnPropertyChanged();
                    }
                }
            }

            public string Sheet { get; set; }
            public string ScheduleName { get; set; }
            public int WordStartIndex { get; set; }
            public int WordLength { get; set; }

            private bool _isIgnored;
            public bool IsIgnored
            {
                get => _isIgnored;
                set
                {
                    if (_isIgnored != value)
                    {
                        _isIgnored = value;
                        OnPropertyChanged();
                    }
                }
            }

            private bool _isAccepted;
            public bool IsAccepted
            {
                get => _isAccepted;
                set
                {
                    if (_isAccepted != value)
                    {
                        _isAccepted = value;
                        OnPropertyChanged();
                    }
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
            private void OnPropertyChanged([CallerMemberName] string name = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        #endregion

        #region Custom Dictionary Management

        private void LoadEnglishDictionary()
        {
            // Try to load dictionary from a text file
            // Place "dictionary.txt" in the same folder as your DLL
            // Or specify a full path

            try
            {
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string assemblyDir = System.IO.Path.GetDirectoryName(assemblyPath);
                string dictionaryPath = System.IO.Path.Combine(assemblyDir, "dictionary.txt");

                if (System.IO.File.Exists(dictionaryPath))
                {
                    // Load all words from file (one word per line)
                    var words = System.IO.File.ReadAllLines(dictionaryPath)
                        .Where(line => !string.IsNullOrWhiteSpace(line))
                        .Select(line => line.Trim());

                    _englishDictionary = new HashSet<string>(words, StringComparer.OrdinalIgnoreCase);

                    StatusText.Text = $"Loaded {_englishDictionary.Count} words from dictionary";
                }
                else
                {
                    // If file doesn't exist, load a minimal fallback dictionary
                    LoadFallbackDictionary();
                    StatusText.Text = "Dictionary file not found - using fallback dictionary. Place 'dictionary.txt' in addon folder.";
                }
            }
            catch (Exception ex)
            {
                LoadFallbackDictionary();
                StatusText.Text = $"Error loading dictionary: {ex.Message}";
            }
        }

        private void LoadFallbackDictionary()
        {
            // Minimal fallback dictionary if file is not available
            var commonWords = new[]
            {
                // Articles, pronouns, conjunctions
                "a", "an", "the", "and", "or", "but", "if", "because", "as", "until", "while",
                "of", "at", "by", "for", "with", "about", "against", "between", "into", "through",
                "during", "before", "after", "above", "below", "to", "from", "up", "down", "in",
                "out", "on", "off", "over", "under", "this", "that", "these", "those",
                
                // Common verbs
                "be", "am", "is", "are", "was", "were", "been", "have", "has", "had", "do",
                "does", "did", "will", "would", "should", "could", "can", "may", "might",
                "make", "take", "come", "go", "see", "get", "give", "use", "find", "tell",
                
                // Common nouns
                "time", "year", "day", "way", "work", "part", "place", "case", "point",
                "wall", "floor", "ceiling", "roof", "door", "window", "room", "building",
                
                // Common adjectives
                "good", "new", "first", "last", "long", "great", "small", "large", "old", "different"
            };

            _englishDictionary = new HashSet<string>(commonWords, StringComparer.OrdinalIgnoreCase);
        }

        private void LoadCustomDictionary()
        {
            try
            {
                // Try to load from project information
                var param = _doc.ProjectInformation.LookupParameter("QPK_CustomDictionary");
                if (param != null && param.StorageType == StorageType.String)
                {
                    string dictData = param.AsString();
                    if (!string.IsNullOrEmpty(dictData))
                    {
                        var words = dictData.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        _customDictionary = new HashSet<string>(words, StringComparer.OrdinalIgnoreCase);
                    }
                }
            }
            catch
            {
                // If parameter doesn't exist, we'll create it when saving
            }
        }

        private void SaveCustomDictionary()
        {
            try
            {
                using (var tx = new Transaction(_doc, "Save Custom Dictionary"))
                {
                    tx.Start();

                    var param = _doc.ProjectInformation.LookupParameter("QPK_CustomDictionary");
                    if (param == null)
                    {
                        // Try to create the parameter (this may fail if not set up in project)
                        // In production, you'd want to ensure this shared parameter exists
                    }
                    else
                    {
                        string dictData = string.Join("|", _customDictionary);
                        param.Set(dictData);
                    }

                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Custom Dictionary", $"Could not save custom dictionary: {ex.Message}");
            }
        }

        public void AddToCustomDictionary(string word)
        {
            if (!string.IsNullOrWhiteSpace(word))
            {
                _customDictionary.Add(word);
                SaveCustomDictionary();
            }
        }

        public bool IsInCustomDictionary(string word)
        {
            return _customDictionary.Contains(word);
        }

        #endregion

        #region Spell Check Functions

        private void RunSpellCheck()
        {
            SpellingResults.Clear();
            CurrentSpellingIndex = -1;

            var seenTypeIds = new HashSet<ElementId>();
            int scheduleCount = 0;
            int errorCount = 0;

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

                    // Get NUMBER from the TYPE parameter
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

                    string comment = GetTypeComment(tElem, string.Empty);
                    if (string.IsNullOrWhiteSpace(comment))
                        continue;

                    // Check spelling on this comment
                    var spellingErrors = FindSpellingErrors(comment);
                    if (spellingErrors.Any())
                    {
                        seenTypeIds.Add(typeId);
                        string sheetName = GetSheetNameForElement(elem);

                        foreach (var error in spellingErrors)
                        {
                            errorCount++;
                            SpellingResults.Add(new SpellingResult
                            {
                                TypeId = typeId,
                                Number = number,
                                FullComment = comment,
                                FullCorrectedComment = comment,
                                MisspelledWord = error.Word,
                                Suggestions = new ObservableCollection<string>(error.Suggestions),
                                SelectedSuggestion = error.Suggestions.FirstOrDefault(),
                                Sheet = sheetName,
                                ScheduleName = vs.Name,
                                WordStartIndex = error.StartIndex,
                                WordLength = error.Length
                            });
                        }
                    }
                }
            }

            if (SpellingResults.Any())
            {
                CurrentSpellingIndex = 0;
                MessageBox.Show(
                    $"Found {errorCount} spelling error(s) in {scheduleCount} schedule(s).",
                    "Spell Check Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show(
                    $"No spelling errors found in {scheduleCount} schedule(s).",
                    "Spell Check Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }

        private List<SpellingError> FindSpellingErrors(string text)
        {
            var errors = new List<SpellingError>();
            if (string.IsNullOrWhiteSpace(text))
                return errors;

            // Split text into words and check each
            var words = System.Text.RegularExpressions.Regex.Split(text, @"\b");
            int currentIndex = 0;

            foreach (var word in words)
            {
                if (!string.IsNullOrWhiteSpace(word) &&
                    System.Text.RegularExpressions.Regex.IsMatch(word, @"^[a-zA-Z]+$"))
                {
                    // Skip if in custom dictionary
                    if (!IsInCustomDictionary(word))
                    {
                        // Check if word is misspelled
                        if (!IsWordCorrect(word))
                        {
                            var suggestions = GetSuggestions(word);
                            errors.Add(new SpellingError
                            {
                                Word = word,
                                StartIndex = currentIndex,
                                Length = word.Length,
                                Suggestions = suggestions
                            });
                        }
                    }
                }
                currentIndex += word.Length;
            }

            return errors;
        }

        private bool IsWordCorrect(string word)
        {
            try
            {
                // Check if word is in English dictionary first
                if (_englishDictionary.Contains(word))
                    return true;

                // Words that are all caps (3+ chars), all numbers, or very short might be acronyms/codes
                if (word.Length <= 2)
                    return true;

                if (word.Length >= 3 && word.All(char.IsUpper))
                    return true;

                if (word.All(char.IsDigit))
                    return true;

                // Common construction/architectural abbreviations not in main dictionary
                var technicalTerms = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "gypbd", "conc", "mtl", "alum", "galv", "thk", "wd", "clg", "qty",
                    "typ", "min", "max", "dia", "lbs", "psf", "psi", "btu", "hvac",
                    "mep", "rcp", "dwg", "spec", "nts", "sim", "approx", "elev",
                    "dwgs", "specs", "mech", "elec", "arch", "struct", "dims",
                    "horiz", "vert", "nom", "req", "min", "max", "approx", "dia",
                    "thru", "w", "ctrs", "ea", "incl", "excl", "cont", "reinf",
                    "galv", "alum", "ss", "ci", "pvc", "abs", "cpvc", "hdpe",
                    "epdm", "tpo", "sbs", "gyp", "clg", "susp", "acous", "insul",
                    "vapor", "retarder", "dampproofing", "waterproofing", "cwp"
                };

                if (technicalTerms.Contains(word))
                    return true;

                // If not found in any dictionary, it's potentially misspelled
                return false;
            }
            catch
            {
                return true; // If we can't check, assume it's correct
            }
        }

        private List<string> GetSuggestions(string word)
        {
            // In production, use a proper spell check library like NHunspell
            // For now, return some basic suggestions
            var suggestions = new List<string>();

            // Add some common corrections
            var commonCorrections = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "teh", "the" },
                { "adn", "and" },
                { "hte", "the" },
                { "taht", "that" },
                { "thier", "their" },
                { "occured", "occurred" },
                { "recieve", "receive" }
            };

            if (commonCorrections.TryGetValue(word, out string correction))
            {
                suggestions.Add(correction);
            }

            // Add the word itself as an option to add to dictionary
            suggestions.Add($"[Add '{word}' to dictionary]");

            return suggestions;
        }

        private class SpellingError
        {
            public string Word { get; set; }
            public int StartIndex { get; set; }
            public int Length { get; set; }
            public List<string> Suggestions { get; set; }
        }

        #endregion

        #region UI Mode Management

        private void UpdateUIMode()
        {
            if (IsSpellCheckMode)
            {
                // Show spell check UI, hide find/replace UI
                FindReplacePanel.Visibility = System.Windows.Visibility.Collapsed;
                SpellCheckPanel.Visibility = System.Windows.Visibility.Visible;
                ResultsDataGrid.Visibility = System.Windows.Visibility.Collapsed;
                SpellingDataGrid.Visibility = System.Windows.Visibility.Visible;
            }
            else
            {
                // Show find/replace UI, hide spell check UI
                FindReplacePanel.Visibility = System.Windows.Visibility.Visible;
                SpellCheckPanel.Visibility = System.Windows.Visibility.Collapsed;
                ResultsDataGrid.Visibility = System.Windows.Visibility.Visible;
                SpellingDataGrid.Visibility = System.Windows.Visibility.Collapsed;
            }
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
            isCaseSensitive = caseSensitiveToggle.IsChecked.GetValueOrDefault();
        }

        internal string GetTypeComment(Element tElem, string search)
        {
            if (tElem == null) return string.Empty;

            Parameter p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null)
            {
                string s = p.AsString();
                if (!string.IsNullOrEmpty(s))
                {
                    if (string.IsNullOrEmpty(search) || CompareStrings(s, search, isCaseSensitive))
                        return s;
                }
            }

            string[] names = { "Type Comments", "Comments", "Comment", "COMMENT" };
            foreach (string nm in names)
            {
                p = tElem.LookupParameter(nm);
                if (p != null)
                {
                    string s = p.AsString();
                    if (!string.IsNullOrEmpty(s))
                    {
                        if (string.IsNullOrEmpty(search) || CompareStrings(s, search, isCaseSensitive))
                            return s;
                    }
                }
            }

            return string.Empty;
        }

        private bool CompareStrings(string str1, string str2, bool caseSensitive)
        {
            if (string.IsNullOrEmpty(str1) || string.IsNullOrEmpty(str2))
                return false;

            var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            return str1.IndexOf(str2, comparison) >= 0;
        }

        private static string ReplaceWithComparison(string input, string oldValue, string newValue, bool caseSensitive)
        {
            if (string.IsNullOrEmpty(oldValue) || string.IsNullOrEmpty(input))
                return input;

            if (caseSensitive)
                return input.Replace(oldValue, newValue);

            var comparison = StringComparison.OrdinalIgnoreCase;
            int idx = input.IndexOf(oldValue, comparison);
            if (idx < 0) return input;

            var sb = new StringBuilder();
            int start = 0;
            while (idx >= 0)
            {
                sb.Append(input, start, idx - start);
                sb.Append(newValue);
                start = idx + oldValue.Length;
                idx = input.IndexOf(oldValue, start, comparison);
            }
            sb.Append(input, start, input.Length - start);
            return sb.ToString();
        }

        internal bool SetTypeComment(Element tElem, string newText)
        {
            if (tElem == null) return false;

            Parameter p = tElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null && !p.IsReadOnly)
            {
                string current = p.AsString() ?? string.Empty;
                bool equal = isCaseSensitive
                    ? current.Equals(newText, StringComparison.Ordinal)
                    : current.Equals(newText, StringComparison.OrdinalIgnoreCase);

                if (equal)
                    return false;

                return p.Set(newText);
            }

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
                        return false;

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
            var comparison = isCaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            for (int i = 0; i < tokens.Length; i++)
            {
                string token = tokens[i];
                string core = token.Trim(',', '.', ';', ':', '!', '?');

                if (core.IndexOf(search, comparison) >= 0)
                {
                    hitIndex = i;
                    break;
                }
            }

            if (hitIndex == -1)
            {
                foundPrefix = fullOld;
                replPrefix = ReplaceWithComparison(fullOld, search, replace, isCaseSensitive);
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
            replWord = ReplaceWithComparison(matchedToken, search, replace, isCaseSensitive);
            replSuffix = foundSuffix;
        }

        #endregion

        #region Exposed helpers for handlers

        internal ReplaceResult GetSelectedResult()
        {
            return ResultsDataGrid.SelectedItem as ReplaceResult;
        }

        internal SpellingResult GetCurrentSpellingResult()
        {
            if (CurrentSpellingIndex >= 0 && CurrentSpellingIndex < SpellingResults.Count)
                return SpellingResults[CurrentSpellingIndex];
            return null;
        }

        #endregion

        #region UI Handlers - Find & Replace

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

                    string oldComment = GetTypeComment(tElem, search);
                    if (string.IsNullOrEmpty(oldComment))
                        continue;

                    var comparison = isCaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                    if (oldComment.IndexOf(search, comparison) < 0)
                        continue;

                    string newComment = ReplaceWithComparison(oldComment, search, replace, isCaseSensitive);
                    if (newComment == oldComment)
                        continue;

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

        #region UI Handlers - Spell Check

        private void RunSpellCheck_Click(object sender, RoutedEventArgs e)
        {
            RunSpellCheck();
        }

        private void NextSpelling_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentSpellingIndex < SpellingResults.Count - 1)
            {
                CurrentSpellingIndex++;
            }
        }

        private void PreviousSpelling_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentSpellingIndex > 0)
            {
                CurrentSpellingIndex--;
            }
        }

        private void AcceptSpelling_Click(object sender, RoutedEventArgs e)
        {
            var current = GetCurrentSpellingResult();
            if (current == null) return;

            // Update the corrected comment with the selected suggestion
            if (!string.IsNullOrEmpty(current.SelectedSuggestion))
            {
                // Check if user wants to add to dictionary
                if (current.SelectedSuggestion.StartsWith("[Add"))
                {
                    AddToCustomDictionary(current.MisspelledWord);
                    current.IsIgnored = true;
                    MessageBox.Show($"Added '{current.MisspelledWord}' to custom dictionary.",
                        "Dictionary Updated", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    // Replace the misspelled word with the suggestion
                    string before = current.FullComment.Substring(0, current.WordStartIndex);
                    string after = current.FullComment.Substring(current.WordStartIndex + current.WordLength);
                    current.FullCorrectedComment = before + current.SelectedSuggestion + after;
                    current.IsAccepted = true;
                }
            }

            // Move to next error
            if (CurrentSpellingIndex < SpellingResults.Count - 1)
            {
                CurrentSpellingIndex++;
            }
        }

        private void IgnoreSpelling_Click(object sender, RoutedEventArgs e)
        {
            var current = GetCurrentSpellingResult();
            if (current == null) return;

            current.IsIgnored = true;

            // Move to next error
            if (CurrentSpellingIndex < SpellingResults.Count - 1)
            {
                CurrentSpellingIndex++;
            }
        }

        private void AddToDictionary_Click(object sender, RoutedEventArgs e)
        {
            var current = GetCurrentSpellingResult();
            if (current == null) return;

            AddToCustomDictionary(current.MisspelledWord);
            current.IsIgnored = true;

            MessageBox.Show($"Added '{current.MisspelledWord}' to custom dictionary.",
                "Dictionary Updated", MessageBoxButton.OK, MessageBoxImage.Information);

            // Move to next error
            if (CurrentSpellingIndex < SpellingResults.Count - 1)
            {
                CurrentSpellingIndex++;
            }
        }

        private void ApplyAllSpelling_Click(object sender, RoutedEventArgs e)
        {
            _applySpellingEvent.Raise();
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
                    {
                        changed++;
                        r.IsApplied = true;
                    }
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
                    selected.IsApplied = true;
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

    public class ApplySpellingHandler : IExternalEventHandler
    {
        private readonly MainWindow _window;

        public ApplySpellingHandler(MainWindow window)
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
            var results = _window.SpellingResults;

            if (results == null || !results.Any())
            {
                TaskDialog.Show("QPK Keynote Manager",
                    "No spelling corrections to apply. Run Spell Check first.");
                return;
            }

            // Group by TypeId to avoid multiple updates to same element
            var grouped = results
                .Where(r => r.IsAccepted && !r.IsIgnored)
                .GroupBy(r => r.TypeId)
                .ToList();

            using (var tx = new Transaction(doc, "Apply Spelling Corrections"))
            {
                tx.Start();

                int changed = 0;
                foreach (var group in grouped)
                {
                    var tElem = doc.GetElement(group.Key);
                    if (tElem == null) continue;

                    // Get the current comment
                    string currentComment = _window.GetTypeComment(tElem, string.Empty);
                    string correctedComment = currentComment;

                    // Apply all corrections for this element (in reverse order to maintain indices)
                    var orderedCorrections = group.OrderByDescending(r => r.WordStartIndex).ToList();
                    foreach (var correction in orderedCorrections)
                    {
                        if (!string.IsNullOrEmpty(correction.FullCorrectedComment))
                        {
                            correctedComment = correction.FullCorrectedComment;
                        }
                    }

                    if (correctedComment != currentComment)
                    {
                        if (_window.SetTypeComment(tElem, correctedComment))
                        {
                            changed++;
                        }
                    }
                }

                tx.Commit();

                TaskDialog.Show("QPK Keynote Manager",
                    $"Applied spelling corrections to {changed} type(s).");
            }
        }

        public string GetName()
        {
            return "QPK Keynote Manager – Apply Spelling Corrections";
        }
    }

    #endregion
}