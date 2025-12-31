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
            // ✅ EARLY-OUT GUARD (runs once per click)
            if (row.IsApplied)
                return;

            if (row == null || row.SheetId == null || row.SheetId == ElementId.InvalidElementId)
            {
                TaskDialog.Show("QPK Keynote Manager", "Selected row is missing a valid SheetId.");
                return;
            }

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
                if (ok)
                {
                    tx.Commit();
                    row.IsApplied = true;  // ✅ mark row green
                }
                else
                {
                    tx.RollBack();
                }


                TaskDialog.Show("QPK Keynote Manager",
                    ok
                        ? $"Updated sheet name:\n\n{row.FoundText}  →  {newName}"
                        : "Failed to update sheet name.");
            }
        }

        private void ApplySelectedKeynoteChange(Document doc, ReplaceResult row)
        {
            if (row == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Select a keynote row first.");
                return;
            }

            // ✅ EARLY-OUT GUARD (runs once per click)
            if (row.IsApplied)
                return;

            if (row.TypeId == null || row.TypeId == ElementId.InvalidElementId)
            {
                TaskDialog.Show("QPK Keynote Manager", "Selected keynote row is missing a valid TypeId.");
                return;
            }

            var tElem = doc.GetElement(row.TypeId);
            if (tElem == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Could not find the Type element for the selected row.");
                return;
            }

            string newText = (row.FullNewComment ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(newText))
            {
                TaskDialog.Show("QPK Keynote Manager", "Proposed comment is empty. Skipping.");
                return;
            }

            bool isCaseSensitive = _window?.VM?.IsCaseSensitive ?? false;

            using (var tx = new Transaction(doc, "Replace Type Comment (Selected)"))
            {
                tx.Start();

                bool ok = _window.SetTypeComment(tElem, newText, isCaseSensitive);

                if (ok)
                {
                    tx.Commit();
                    row.IsApplied = true; // ✅ mark row green
                    TaskDialog.Show("QPK Keynote Manager", "Updated Type Comments for selected keynote.");
                }
                else
                {
                    tx.RollBack();
                    TaskDialog.Show(
                        "QPK Keynote Manager",
                        "Failed to set Type Comments.\n\n" +
                        "Possible reasons:\n" +
                        "- Parameter is read-only\n" +
                        "- Parameter not found (Type Comments / Comments)\n" +
                        "- New text matches existing value");
                }
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
