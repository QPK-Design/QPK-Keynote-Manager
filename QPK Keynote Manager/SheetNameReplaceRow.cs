using System.ComponentModel;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;

namespace QPK_Keynote_Manager
{
    //THIS IS THE SHEET MODEL.
    public class SheetNameReplaceRow : IReplaceRow, INotifyPropertyChanged
    {
        public ElementId? SheetId { get; set; }
        public string? SheetNumber { get; set; }
        public string? Sheet { get; set; }
        public string? FoundText { get; set; }
        public string? ReplacedText { get; set; }

        // === NEW: Found text highlighting parts ===
        public string? FoundPrefix { get; set; }
        public string? FoundWord { get; set; }
        public string? FoundSuffix { get; set; }

        // === NEW: Replaced text highlighting parts ===
        public string? ReplPrefix { get; set; }
        public string? ReplWord { get; set; }
        public string? ReplSuffix { get; set; }

        // Found (original) - word-aware segments
        public string? FoundPreText { get; set; }
        public string? FoundWordLeft { get; set; }   // part of the same word before match (gray)
        public string? FoundWordMid { get; set; }    // match itself (green)
        public string? FoundWordRight { get; set; }  // remainder of the same word after match (gray)
        public string? FoundPostText { get; set; }

        // Replaced - word-aware segments
        public string? ReplPreText { get; set; }
        public string? ReplWordLeft { get; set; }    // (gray)
        public string? ReplWordMid { get; set; }     // replacement (red)
        public string? ReplWordRight { get; set; }   // (gray)
        public string? ReplPostText { get; set; }

        // Required by IReplaceRow
        public string StringFound => FoundText;

        public string StringReplaced => ReplacedText;

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
