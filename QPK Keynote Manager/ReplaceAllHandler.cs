// ReplaceAllHandler.cs
using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QPK_Keynote_Manager
{
    public class ReplaceAllHandler : IExternalEventHandler
    {
        private readonly MainWindow _window;

        public ReplaceAllHandler(MainWindow window)
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

            try
            {
                switch (vm.SelectedScope.Kind)
                {
                    case FindReplaceScopeKind.SheetNames:
                        ApplyAllSheetNameChanges(doc, vm);
                        break;

                    case FindReplaceScopeKind.Keynotes:
                        ApplyAllKeynoteChanges(doc, vm); // <-- instance method now
                        break;

                    case FindReplaceScopeKind.ViewTitles:
                        TaskDialog.Show("QPK Keynote Manager",
                            "Replace All for View Titles is not wired yet in this step.");
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

        private static void ApplyAllSheetNameChanges(Document doc, MainViewModel vm)
        {
            if (vm.SheetNameResults == null || vm.SheetNameResults.Count == 0)
            {
                TaskDialog.Show("QPK Keynote Manager", "No sheet name changes to apply. Run Preview first.");
                return;
            }

            int changed = 0;
            int failed = 0;
            int skipped = 0;

            using (var tx = new Transaction(doc, "QPK Find & Replace — Sheet Names (All)"))
            {
                tx.Start();

                foreach (var row in vm.SheetNameResults.ToList())
                {
                    if (row == null)
                    {
                        failed++;
                        continue;
                    }

                    if (row.IsApplied)
                    {
                        skipped++;
                        continue; // ✅ skip already-applied rows
                    }

                    if (row.SheetId == null || row.SheetId == ElementId.InvalidElementId)
                    {
                        failed++;
                        continue;
                    }

                    var sheet = doc.GetElement(row.SheetId) as ViewSheet;
                    if (sheet == null)
                    {
                        failed++;
                        continue;
                    }

                    var p = sheet.get_Parameter(BuiltInParameter.SHEET_NAME);
                    if (p == null || p.IsReadOnly)
                    {
                        failed++;
                        continue;
                    }

                    var newName = (row.ReplacedText ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(newName))
                    {
                        failed++;
                        continue;
                    }

                    if (string.Equals(sheet.Name ?? string.Empty, newName, StringComparison.Ordinal))
                    {
                        // already matches; treat as skipped
                        skipped++;
                        row.IsApplied = true; // optional: mark green since it's effectively applied
                        continue;
                    }

                    bool ok = p.Set(newName);
                    if (ok)
                    {
                        changed++;
                        row.IsApplied = true;  // ✅ mark row green
                    }
                    else
                    {
                        failed++;
                    }
                }

                if (changed > 0)
                    tx.Commit();
                else
                    tx.RollBack();
            }

            TaskDialog.Show("QPK Keynote Manager",
                $"Sheet Names — Replace All complete.\n\nChanged: {changed}\nSkipped: {skipped}\nFailed: {failed}");
        }

        // ✅ MUST be instance method (uses _window)
        private void ApplyAllKeynoteChanges(Document doc, MainViewModel vm)
        {
            if (vm.KeynoteResults == null || vm.KeynoteResults.Count == 0)
            {
                TaskDialog.Show("QPK Keynote Manager", "No keynote changes to apply. Run Preview first.");
                return;
            }

            bool isCaseSensitive = vm.IsCaseSensitive;

            int changed = 0;
            int failed = 0;
            int skipped = 0;

            using (var tx = new Transaction(doc, "Replace Type Comments (All)"))
            {
                tx.Start();

                foreach (var row in vm.KeynoteResults.ToList())
                {
                    if (row == null)
                    {
                        failed++;
                        continue;
                    }

                    if (row.IsApplied)
                    {
                        skipped++;
                        continue; // ✅ Don’t run twice
                    }

                    if (row.TypeId == null || row.TypeId == ElementId.InvalidElementId)
                    {
                        failed++;
                        continue;
                    }

                    var tElem = doc.GetElement(row.TypeId);
                    if (tElem == null)
                    {
                        failed++;
                        continue;
                    }

                    string newText = (row.FullNewComment ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(newText))
                    {
                        failed++;
                        continue;
                    }

                    bool ok = _window.SetTypeComment(tElem, newText, isCaseSensitive);

                    if (ok)
                    {
                        row.IsApplied = true; // ✅ mark green
                        changed++;
                    }
                    else
                    {
                        // Could be read-only OR missing param OR already equal
                        failed++;
                    }
                }

                if (changed > 0)
                    tx.Commit();
                else
                    tx.RollBack();
            }

            TaskDialog.Show("QPK Keynote Manager",
                $"Keynotes — Replace All complete.\n\nChanged: {changed}\nSkipped: {skipped}\nFailed/No-Op: {failed}");
        }

        public string GetName()
        {
            return "QPK Keynote Manager - Replace All Handler";
        }
    }
}
