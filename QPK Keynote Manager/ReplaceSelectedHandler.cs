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
            UIDocument? uidoc = app?.ActiveUIDocument;
            Document? doc = uidoc?.Document;

            if (doc == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "No active Revit document found.");
                return;
            }

            MainViewModel vm = _window.VM;
            if (vm?.SelectedScope == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Select a scope first.");
                return;
            }

            IReplaceRow? selected = _window.GetSelectedRow();
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
                        if (selected is ViewTitleNameReplaceRow vrow)
                            ApplySelectedViewTitleChange(doc, vm, vrow);
                        else
                            TaskDialog.Show("QPK Keynote Manager", "Selected row is not a View Title/Name result.");
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

        private static void ApplySelectedViewTitleChange(Document doc, MainViewModel vm, ViewTitleNameReplaceRow row)
        {
            if (row == null) return;

            if (row.IsApplied)
                return;

            if (row.ViewId == null || row.ViewId == ElementId.InvalidElementId)
            {
                TaskDialog.Show("QPK Keynote Manager", "Selected row is missing a valid ViewId.");
                return;
            }

            View? view = doc.GetElement(row.ViewId) as View;
            if (view == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Could not find the view for the selected row.");
                return;
            }

            string proposed = (row.FullNewText ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(proposed))
            {
                TaskDialog.Show("QPK Keynote Manager", "Proposed value is empty. Skipping.");
                return;
            }

            using (Transaction tx = new Transaction(doc, "QPK Find & Replace — View Title/Name (Selected)"))
            {
                tx.Start();

                bool ok = false;

                if (string.Equals(row.Mode, "VT", StringComparison.OrdinalIgnoreCase))
                {
                    Parameter p = view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION);
                    if (p != null && !p.IsReadOnly)
                        ok = p.Set(proposed);
                }
                else if (string.Equals(row.Mode, "VN", StringComparison.OrdinalIgnoreCase))
                {
                    ok = TrySetViewNameUnique(doc, view, proposed, out string actual);
                    if (ok)
                        row.FullNewText = actual; // what we actually applied (in case (2), (3), etc.)
                }

                if (ok)
                {
                    tx.Commit();

                    // Mark ALL matching rows (same view + same mode) as applied
                    foreach (ViewTitleNameReplaceRow r in vm.ViewTitleResults)
                    {
                        if (r.ViewId == row.ViewId &&
                            string.Equals(r.Mode, row.Mode, StringComparison.OrdinalIgnoreCase))
                        {
                            r.IsApplied = true;
                        }
                    }

                    TaskDialog.Show("QPK Keynote Manager", "Updated view title/name for selected row.");
                }
                else
                {
                    tx.RollBack();
                    TaskDialog.Show("QPK Keynote Manager",
                        "Failed to update view title/name.\n\nPossible reasons:\n- Parameter is read-only\n- Proposed text is invalid\n- View Name conflict");
                }
            }
        }

        private static bool TrySetViewNameUnique(Document doc, View view, string proposed, out string actual)
        {
            actual = proposed;

            if (doc == null || view == null) return false;
            if (string.IsNullOrWhiteSpace(proposed)) return false;

            // Collect existing view names
            HashSet<string> taken = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Select(v => v.Name ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);

            // Allow the current name to be replaced
            taken.Remove(view.Name ?? string.Empty);

            // Make unique: "Name", "Name (2)", "Name (3)"...
            if (!taken.Contains(actual))
                return TrySetViewName(view, actual);

            int n = 2;
            while (true)
            {
                string cand = $"{proposed} ({n})";
                if (!taken.Contains(cand))
                {
                    actual = cand;
                    return TrySetViewName(view, actual);
                }
                n++;
            }
        }

        private static bool TrySetViewName(View view, string newName)
        {
            try
            {
                view.Name = newName;
                return true;
            }
            catch
            {
                return false;
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

            ViewSheet? sheet = doc.GetElement(row.SheetId) as ViewSheet;
            if (sheet == null)
            {
                TaskDialog.Show("QPK Keynote Manager", "Could not find the sheet element for the selected row.");
                return;
            }

            Parameter p = sheet.get_Parameter(BuiltInParameter.SHEET_NAME);
            if (p == null || p.IsReadOnly)
            {
                TaskDialog.Show("QPK Keynote Manager", "Sheet Name parameter is missing or read-only.");
                return;
            }

            string newName = row.ReplacedText?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(newName))
            {
                TaskDialog.Show("QPK Keynote Manager", "Proposed sheet name is empty. Skipping.");
                return;
            }

            using (Transaction tx = new Transaction(doc, "QPK Find & Replace — Sheet Name (Selected)"))
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

            Element tElem = doc.GetElement(row.TypeId);
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

            using (Transaction tx = new Transaction(doc, "Replace Type Comment (Selected)"))
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


        private static Parameter? TryGetTypeCommentsParam(Element typeElem)
        {
            if (typeElem == null) return null;

            // Best: BuiltInParameter
            Parameter p = typeElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (p != null) return p;

            // Fallbacks (in case of custom/shared params)
            p = typeElem.LookupParameter("Type Comments")
                ?? typeElem.LookupParameter("Comments")
                ?? typeElem.LookupParameter("Comment");

            return p;
        }

        public string GetName() => "QPK Keynote Manager - Replace Selected Handler";
    }
}
