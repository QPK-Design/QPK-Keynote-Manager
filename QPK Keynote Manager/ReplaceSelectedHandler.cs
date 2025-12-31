// ReplaceSelectedHandler.cs
using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QPK_Keynote_Manager
{
    public class ReplaceSelectedHandler : IExternalEventHandler
    {
        private readonly MainWindow _window;

        public ReplaceSelectedHandler(MainWindow window)
        {
            _window = window ?? throw new ArgumentNullException(nameof(window));
        }

        public void Execute(UIApplication app)
        {
            var uidoc = app?.ActiveUIDocument;
            var doc = uidoc?.Document;

            if (doc == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "No active Revit document found.");
                return;
            }

            var vm = _window.VM;
            if (vm?.SelectedScope == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Select a scope first.");
                return;
            }

            var selected = _window.GetSelectedRow();
            if (selected == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Select a row first.");
                return;
            }

            try
            {
                switch (vm.SelectedScope.Kind)
                {
                    case FindReplaceScopeKind.SheetNames:
                        if (selected is SheetNameReplaceRow srow)
                            ApplySelectedSheetNameChange(doc, srow);
                        else
                            TaskDialog.Show("QPK Keynote Manager", "Selected row is not a Sheet Name result.");
                        break;

                    case FindReplaceScopeKind.Keynotes:
                        if (selected is ReplaceResult krow)
                            ApplySelectedKeynoteChange(doc, krow);
                        else
                            TaskDialog.Show("QPK Keynote Manager", "Selected row is not a Keynote result.");
                        break;

                    case FindReplaceScopeKind.ViewTitles:
                        TaskDialog.Show("QPK Keynote Manager",
                            "Replace Selected for View Titles is not wired yet in this step.");
                        break;

                    default:
                        TaskDialog.Show("QPK Keynote Manager", "Unknown scope selected.");
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("QPK Keynote Manager", $"Error:\n{ex}");
            }
        }

        private static void ApplySelectedSheetNameChange(Document doc, SheetNameReplaceRow row)
        {
            if (row == null || row.SheetId == null || row.SheetId == ElementId.InvalidElementId)
            {
                TaskDialog.Show("QPK Keynote Manager", "Selected row is missing a valid SheetId.");
                return;
            }

            var sheet = doc.GetElement(row.SheetId) as ViewSheet;
            if (sheet == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Could not find the sheet element for the selected row.");
                return;
            }

            var p = sheet.get_Parameter(BuiltInParameter.SHEET_NAME);
            if (p == null || p.IsReadOnly)
            {
                TaskDialog.Show("QPK Keynote Manager", "Sheet Name parameter is missing or read-only.");
                return;
            }

            var newName = row.ReplacedText?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(newName))
            {
                TaskDialog.Show("QPK Keynote Manager", "Proposed sheet name is empty. Skipping.");
                return;
            }

            using (var tx = new Transaction(doc, "QPK Find & Replace — Sheet Name (Selected)"))
            {
                tx.Start();
                bool ok = p.Set(newName);
                if (ok) tx.Commit();
                else tx.RollBack();

                TaskDialog.Show("QPK Keynote Manager",
                    ok
                        ? $"Updated sheet name:\n\n{row.FoundText}  →  {newName}"
                        : "Failed to update sheet name.");
            }
        }

        private static void ApplySelectedKeynoteChange(Document doc, ReplaceResult row)
        {
            if (row == null || row.TypeId == null || row.TypeId == ElementId.InvalidElementId)
            {
                TaskDialog.Show("QPK Keynote Manager", "Selected keynote row is missing a valid TypeId.");
                return;
            }

            var typeElem = doc.GetElement(row.TypeId);
            if (typeElem == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Could not find the Type element for the selected row.");
                return;
            }

            var p = TryGetTypeCommentsParam(typeElem);
            if (p == null || p.IsReadOnly)
            {
                TaskDialog.Show("QPK Keynote Manager", "Type Comments parameter is missing or read-only.");
                return;
            }

            var newComment = (row.FullNewComment ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(newComment))
            {
                TaskDialog.Show("QPK Keynote Manager", "Proposed comment is empty. Skipping.");
                return;
            }

            using (var tx = new Transaction(doc, "QPK Find & Replace — Keynote (Selected)"))
            {
                tx.Start();

                bool ok = p.Set(newComment);

                if (ok) tx.Commit();
                else tx.RollBack();

                TaskDialog.Show("QPK Keynote Manager",
                    ok
                        ? "Updated Type Comments for selected keynote."
                        : "Failed to update Type Comments for selected keynote.");
            }
        }

        private static Parameter TryGetTypeCommentsParam(Element typeElem)
        {
            if (typeElem == null) return null;

            // Best: BuiltInParameter
            var p = typeElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null) return p;

            // Fallbacks (in case of custom/shared params)
            p = typeElem.LookupParameter("Type Comments")
                ?? typeElem.LookupParameter("Comments")
                ?? typeElem.LookupParameter("Comment");

            return p;
        }

        public string GetName()
        {
            return "QPK Keynote Manager - Replace Selected Handler";
        }
    }
}
