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
                        ApplyAllKeynoteChanges(doc, vm);
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

            using (var tx = new Transaction(doc, "QPK Find & Replace — Sheet Names (All)"))
            {
                tx.Start();

                foreach (var row in vm.SheetNameResults.ToList())
                {
                    if (row.IsApplied)
                        continue;   // ✅ skip already-applied rows
                    if (row == null || row.SheetId == null || row.SheetId == ElementId.InvalidElementId)
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

                    var newName = row.ReplacedText?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(newName))
                    {
                        failed++;
                        continue;
                    }

                    if (string.Equals(sheet.Name ?? string.Empty, newName, StringComparison.Ordinal))
                        continue;

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

                tx.Commit();
            }

            TaskDialog.Show("QPK Keynote Manager",
                $"Sheet Names — Replace All complete.\n\nChanged: {changed}\nFailed/Skipped: {failed}");
        }

        private static void ApplyAllKeynoteChanges(Document doc, MainViewModel vm)
        {
            if (vm.KeynoteResults == null || vm.KeynoteResults.Count == 0)
            {
                TaskDialog.Show("QPK Keynote Manager", "No keynote changes to apply. Run Preview first.");
                return;
            }

            int changed = 0;
            int failed = 0;

            using (var tx = new Transaction(doc, "QPK Find & Replace — Keynotes (All)"))
            {
                tx.Start();

                foreach (var rowObj in vm.KeynoteResults.ToList())
                {
                    var row = rowObj as ReplaceResult;
                    if (row == null || row.TypeId == null || row.TypeId == ElementId.InvalidElementId)
                    {
                        failed++;
                        continue;
                    }

                    var typeElem = doc.GetElement(row.TypeId);
                    if (typeElem == null)
                    {
                        failed++;
                        continue;
                    }

                    var p = TryGetTypeCommentsParam(typeElem);
                    if (p == null || p.IsReadOnly)
                    {
                        failed++;
                        continue;
                    }

                    var newComment = (row.FullNewComment ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(newComment))
                    {
                        failed++;
                        continue;
                    }

                    bool ok = p.Set(newComment);
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

                tx.Commit();
            }

            TaskDialog.Show("QPK Keynote Manager",
                $"Keynotes — Replace All complete.\n\nChanged: {changed}\nFailed/Skipped: {failed}");
        }

        private static Parameter TryGetTypeCommentsParam(Element typeElem)
        {
            if (typeElem == null) return null;

            var p = typeElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null) return p;

            p = typeElem.LookupParameter("Type Comments")
                ?? typeElem.LookupParameter("Comments")
                ?? typeElem.LookupParameter("Comment");

            return p;
        }

        public string GetName()
        {
            return "QPK Keynote Manager - Replace All Handler";
        }
    }
}
