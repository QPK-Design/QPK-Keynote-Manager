using Autodesk.Revit.DB;

namespace QPK_Keynote_Manager
{
    /// <summary>
    /// Result row for Keynotes find/replace preview.
    /// (Minimal implementation for MVVM + DataGrid binding, expandable later.)
    /// </summary>
    public class ReplaceResult : IReplaceRow
    {
        // Keynotes-only column (shown when scope = Keynotes)
        public int Number { get; set; }

        // Common columns
        public string Sheet { get; set; }                 // required by IReplaceRow
        public string StringFound { get; set; }
        public string StringReplaced { get; set; }

        // Optional metadata (useful later for applying changes)
        public ElementId ElementId { get; set; }          // instance id or schedule row id
        public ElementId TypeId { get; set; }             // type id if needed
        public string Notes { get; set; }                 // optional debug/status
    }
}
