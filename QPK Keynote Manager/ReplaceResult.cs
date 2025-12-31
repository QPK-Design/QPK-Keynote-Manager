using Autodesk.Revit.DB;

namespace QPK_Keynote_Manager
{
    /// <summary>
    /// Result row for Keynotes find/replace preview.
    /// Supports word-level coloring via FoundPrefix/FoundWord/FoundSuffix + ReplPrefix/ReplWord/ReplSuffix.
    /// </summary>
    public class ReplaceResult : IReplaceRow
    {
        // Keynotes-only column (optional; show/hide via scope)
        public string Number { get; set; }

        // Required by IReplaceRow
        public string Sheet { get; set; }

        // Useful metadata
        public string ScheduleName { get; set; }
        public ElementId TypeId { get; set; }

        // Full text used for applying change (Revit parameter update)
        public string FullOldComment { get; set; }
        public string FullNewComment { get; set; }

        // Context pieces for UI coloring
        public string FoundPrefix { get; set; }
        public string FoundWord { get; set; }
        public string FoundSuffix { get; set; }

        public string ReplPrefix { get; set; }
        public string ReplWord { get; set; }
        public string ReplSuffix { get; set; }

        // Convenience strings
        public string StringFound => string.Concat(FoundPrefix, FoundWord, FoundSuffix);
        public string StringReplaced => string.Concat(ReplPrefix, ReplWord, ReplSuffix);
    }
}
