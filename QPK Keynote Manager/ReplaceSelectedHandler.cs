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
                        TaskDialog.Show("QPK Keynote Manager",
                            "Replace Selected for Keynotes is not wired yet in this step.");
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

                if (ok)
                    tx.Commit();
                else
                    tx.RollBack();

                TaskDialog.Show("QPK Keynote Manager",
                    ok
                        ? $"Updated sheet name:\n\n{row.FoundText}  →  {newName}"
                        : "Failed to update sheet name.");
            }
        }

        public string GetName()
        {
            return "QPK Keynote Manager - Replace Selected Handler";
        }
    }
}
