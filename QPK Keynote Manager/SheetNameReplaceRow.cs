using Autodesk.Revit.DB;

namespace QPK_Keynote_Manager
{
//THIS IS THE SHEET MODEL.
    public class SheetNameReplaceRow : IReplaceRow
    {
        public ElementId SheetId { get; set; }
        public string SheetNumber { get; set; }   // optional
        public string Sheet { get; set; }         // display column ("A101 - Level 1 Plan")

        public string FoundText { get; set; }     // old sheet name
        public string ReplacedText { get; set; }  // proposed new sheet name
    }
}
