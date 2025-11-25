#region Namespaces
using System;
using System.Windows;                // For Window
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
        // Keep a static instance so we don't open multiple copies.
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

                // If window doesn't exist or was closed, create a new one
                if (_mainWindow == null)
                {
                    // If your MainWindow needs Revit data, you can pass uidoc here later:
                    // _mainWindow = new MainWindow(uidoc);
                    _mainWindow = new MainWindow();

                    // When user closes the window, clear the static reference
                    _mainWindow.Closed += (s, e) => _mainWindow = null;

                    _mainWindow.Show();
                }
                else
                {
                    // If window already exists, just bring it to front
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
