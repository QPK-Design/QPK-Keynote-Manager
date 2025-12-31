using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using System.Linq;

namespace QPK_Keynote_Manager
{
    [Transaction(TransactionMode.Manual)]
    public class WorksetPreloader : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Define your predefined workset names
            List<string> worksetNames = new List<string>
            {
                "Architecture",
                "Linked Structure",
                "Linked MEP",
                "Site",
                "Rendering",
                "Hidden Elements",
                "Shared Levels and Grids"
            };

            try
            {
                // Check if worksharing is enabled
                if (!doc.IsWorkshared)
                {
                    // Enable worksharing WITHOUT wrapping in a transaction
                    // This method creates its own transaction internally
                    doc.EnableWorksharing("Shared Levels and Grids", "Workset1");
                }

                // Create the predefined worksets
                using (Transaction trans = new Transaction(doc, "Create Predefined Worksets"))
                {
                    trans.Start();

                    foreach (string worksetName in worksetNames)
                    {
                        // Check if workset already exists
                        if (!WorksetExists(doc, worksetName))
                        {
                            Workset.Create(doc, worksetName);
                        }
                    }

                    trans.Commit();
                }

                TaskDialog.Show("QPK Worksets", "Predefined worksets created successfully!");
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("QPK Worksets Error", ex.Message);
                return Result.Failed;
            }
        }

        private bool WorksetExists(Document doc, string worksetName)
        {
            FilteredWorksetCollector collector = new FilteredWorksetCollector(doc);
            foreach (Workset workset in collector)
            {
                if (workset.Name == worksetName && workset.Kind == WorksetKind.UserWorkset)
                {
                    return true;
                }
            }
            return false;
        }
    }
}