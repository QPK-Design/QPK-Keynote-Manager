using System.ComponentModel;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;

namespace QPK_Keynote_Manager
{
    // THIS IS THE MODEL
    public class ReplaceResult : IReplaceRow, INotifyPropertyChanged
    {
        public string? Number { get; set; }
        public string? Sheet { get; set; }

        public string? ScheduleName { get; set; }
        public ElementId? TypeId { get; set; }

        public string? FullOldComment { get; set; }
        public string? FullNewComment { get; set; }

        public string? FoundPrefix { get; set; }
        public string? FoundWord { get; set; }
        public string? FoundSuffix { get; set; }

        public string? ReplPrefix { get; set; }
        public string? ReplWord { get; set; }
        public string? ReplSuffix { get; set; }

        // === NEW: word-aware highlighting segments (Found) ===
        public string? FoundPreText { get; set; }
        public string? FoundWordLeft { get; set; }   // gray - same word, before match
        public string? FoundWordMid { get; set; }    // green - the match
        public string? FoundWordRight { get; set; }  // gray - same word, after match
        public string? FoundPostText { get; set; }

        // === NEW: word-aware highlighting segments (Replaced) ===
        public string? ReplPreText { get; set; }
        public string? ReplWordLeft { get; set; }    // gray
        public string? ReplWordMid { get; set; }     // red - replacement
        public string? ReplWordRight { get; set; }   // gray
        public string? ReplPostText { get; set; }

        public string StringFound => string.Concat(FoundPrefix, FoundWord, FoundSuffix);

        public string StringReplaced => string.Concat(ReplPrefix, ReplWord, ReplSuffix);

        private bool _isApplied;
        public bool IsApplied
        {
            get => _isApplied;
            set
            {
                if (_isApplied == value) return;
                _isApplied = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
