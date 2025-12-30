// Pseudocode plan:
// 1. Define a model `ReplaceResult` that represents one find/replace hit.
//    - Include existing properties (Number, Sheet, FoundPrefix, FoundWord, FoundSuffix,
//      ReplPrefix, ReplWord, ReplSuffix, IsApplied).
//    - Add a new property `ObjectType` to hold the Revit object type (e.g., "Schedule").
//    - Implement INotifyPropertyChanged so UI updates when values change.
// 2. Implement `FindReplaceViewModel` with:
//    - ObservableCollection<ReplaceResult> ReplaceResults property bound to the DataGrid.
//    - FindText and ReplaceText properties.
//    - A constructor or method to populate sample data with `ObjectType = "Schedule"`.
//    - A comment/example showing where to set `ObjectType` when scanning Revit objects.
// 3. Usage notes:
//    - Change the XAML DataGridTextColumn binding for the Type column to `Binding="{Binding ObjectType}"`.
//    - In `MainWindow.xaml.cs` set `DataContext = new FindReplaceViewModel();` (or inject it).
//    - When scanning Revit items, detect the object type and set `ReplaceResult.ObjectType` accordingly.
//
// The code below implements the model and viewmodel foundation.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace QPK_Keynote_Manager.ViewModels
{
    public class ReplaceResult : INotifyPropertyChanged
    {
        private int _number;
        private string _sheet = string.Empty;
        private string _foundPrefix = string.Empty;
        private string _foundWord = string.Empty;
        private string _foundSuffix = string.Empty;
        private string _replPrefix = string.Empty;
        private string _replWord = string.Empty;
        private string _replSuffix = string.Empty;
        private bool _isApplied;
        private string _objectType = string.Empty;

        public int Number
        {
            get => _number;
            set => SetField(ref _number, value);
        }

        public string Sheet
        {
            get => _sheet;
            set => SetField(ref _sheet, value);
        }

        public string FoundPrefix
        {
            get => _foundPrefix;
            set => SetField(ref _foundPrefix, value);
        }

        public string FoundWord
        {
            get => _foundWord;
            set => SetField(ref _foundWord, value);
        }

        public string FoundSuffix
        {
            get => _foundSuffix;
            set => SetField(ref _foundSuffix, value);
        }

        public string ReplPrefix
        {
            get => _replPrefix;
            set => SetField(ref _replPrefix, value);
        }

        public string ReplWord
        {
            get => _replWord;
            set => SetField(ref _replWord, value);
        }

        public string ReplSuffix
        {
            get => _replSuffix;
            set => SetField(ref _replSuffix, value);
        }

        public bool IsApplied
        {
            get => _isApplied;
            set => SetField(ref _isApplied, value);
        }

        // New property: Revit object type (e.g., "Schedule", "TitleBlock", "TextElement")
        public string ObjectType
        {
            get => _objectType;
            set => SetField(ref _objectType, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        protected void OnPropertyChanged(string? propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class FindReplaceViewModel : INotifyPropertyChanged
    {
        private string _findText = string.Empty;
        private string _replaceText = string.Empty;

        public ObservableCollection<ReplaceResult> ReplaceResults { get; } = new ObservableCollection<ReplaceResult>();

        public string FindText
        {
            get => _findText;
            set => SetField(ref _findText, value);
        }

        public string ReplaceText
        {
            get => _replaceText;
            set => SetField(ref _replaceText, value);
        }

        public FindReplaceViewModel()
        {
            // Populate initial/sample results. In production, replace with Revit scanning logic.
            PopulateSampleResults();
        }

        private void PopulateSampleResults()
        {
            ReplaceResults.Clear();

            ReplaceResults.Add(new ReplaceResult
            {
                Number = 1,
                Sheet = "KeySchedule1",
                FoundPrefix = "Before ",
                FoundWord = "OLD",
                FoundSuffix = " text",
                ReplPrefix = "Before ",
                ReplWord = "NEW",
                ReplSuffix = " text",
                IsApplied = false,
                ObjectType = "Schedule" // default for now
            });

            ReplaceResults.Add(new ReplaceResult
            {
                Number = 2,
                Sheet = "KeySchedule2",
                FoundPrefix = "",
                FoundWord = "Example",
                FoundSuffix = "",
                ReplPrefix = "",
                ReplWord = "Sample",
                ReplSuffix = "",
                IsApplied = false,
                ObjectType = "Schedule"
            });
        }

        // Example usage when scanning Revit objects:
        // foreach (var scanned in scannedItems)
        // {
        //     var detectedType = DetectRevitObjectType(scanned); // returns "Schedule", "TitleBlock", etc.
        //     var result = new ReplaceResult
        //     {
        //         Number = nextNumber++,
        //         Sheet = scanned.SheetName,
        //         FoundPrefix = scanned.Prefix,
        //         FoundWord = scanned.Word,
        //         FoundSuffix = scanned.Suffix,
        //         ReplPrefix = scanned.Prefix,
        //         ReplWord = scanned.Replacement,
        //         ReplSuffix = scanned.Suffix,
        //         IsApplied = false,
        //         ObjectType = detectedType
        //     };
        //     ReplaceResults.Add(result);
        // }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        protected void OnPropertyChanged(string? propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}