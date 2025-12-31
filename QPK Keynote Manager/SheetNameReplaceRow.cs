using System.ComponentModel;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;

namespace QPK_Keynote_Manager
{
    //THIS IS THE SHEET MODEL.
    public class SheetNameReplaceRow : IReplaceRow, INotifyPropertyChanged
    {
        public ElementId SheetId { get; set; }
        public string SheetNumber { get; set; }
        public string Sheet { get; set; }
        public string FoundText { get; set; }
        public string ReplacedText { get; set; }

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
