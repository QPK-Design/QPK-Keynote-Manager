#region Namespaces
using System;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace QPK_Keynote_Manager
{
    /// <summary>
    /// Command that opens the modeless QPK Keynote Manager window.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ShowMainWindowCommand : IExternalCommand
    {
        // Keep a static reference so we don't open multiple windows.
        private static MainWindow _mainWindow;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;

                // Create window if none exists (or was closed)
                if (_mainWindow == null)
                {
                    _mainWindow = new MainWindow(uidoc);  // <-- IMPORTANT
                    _mainWindow.Closed += (s, e) => _mainWindow = null;
                    _mainWindow.Show();                  // modeless window
                }
                else
                {
                    // Bring to front if already open
                    if (!_mainWindow.IsVisible)
                        _mainWindow.Show();

                    _mainWindow.Activate();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
