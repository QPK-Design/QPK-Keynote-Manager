using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;


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
