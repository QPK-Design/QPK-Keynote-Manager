using System.ComponentModel;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;

namespace QPK_Keynote_Manager
{
    /// <summary>
    /// Row used for the "View Titles" scope (actually: View Name + Title on Sheet).
    /// Mode:
    ///   "VN" = View Name
    ///   "VT" = View Title (Title on Sheet)
    /// </summary>
    public class ViewTitleNameReplaceRow : IReplaceRow, INotifyPropertyChanged
    {
        public ElementId ViewId { get; set; }
        public ElementId SheetId { get; set; }

        // "VN" or "VT"
        public string Mode { get; set; }
        public string ModeToolTip { get; set; }  // "View Name" or "View Title"

        // Full strings (for tooltip + apply)
        public string FoundText { get; set; }     // original (effective) text shown
        public string ReplacedText { get; set; }  // proposed (effective) replacement shown

        // Optional: keep these for apply logic consistency
        public string FullOldText { get; set; }
        public string FullNewText { get; set; }

        // === Word-aware highlighting segments (Found) ===
        public string FoundPreText { get; set; }
        public string FoundWordLeft { get; set; }
        public string FoundWordMid { get; set; }     // green
        public string FoundWordRight { get; set; }
        public string FoundPostText { get; set; }

        // === Word-aware highlighting segments (Replaced) ===
        public string ReplPreText { get; set; }
        public string ReplWordLeft { get; set; }
        public string ReplWordMid { get; set; }      // red
        public string ReplWordRight { get; set; }
        public string ReplPostText { get; set; }

        public string Sheet { get; set; }

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
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
