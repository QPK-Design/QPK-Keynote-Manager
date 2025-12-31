using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace QPK_Keynote_Manager
{
    public class ExternalApp : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            const string tabName = "QPK";
            const string panelName = "QPK Addins";
            try
            {
                // Create custom tab (ignore error if it already exists)
                try
                {
                    a.CreateRibbonTab(tabName);
                }
                catch
                {
                    // Tab already exists – ignore
                }

                // Find or create the panel
                RibbonPanel panel = null;
                foreach (var rp in a.GetRibbonPanels(tabName))
                {
                    if (rp.Name == panelName)
                    {
                        panel = rp;
                        break;
                    }
                }
                if (panel == null)
                {
                    panel = a.CreateRibbonPanel(tabName, panelName);
                }

                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string folder = Path.GetDirectoryName(assemblyPath) ?? string.Empty;

                // FIND & REPLACE BUTTON
                PushButtonData buttonData = new PushButtonData(
                    "QPK_MainWindow",
                    "QPK Keynote Manager",
                    assemblyPath,
                    "QPK_Keynote_Manager.ShowMainWindowCommand"
                );

                string img = Path.Combine(folder, "Resources", "QPKKNM_32.png"); // Find & Replace icon
                if (File.Exists(img))
                {
                    buttonData.LargeImage = LoadPng(img);
                }

                panel.AddItem(buttonData);

                // PRELOAD WORKSETS BUTTON
                PushButtonData worksetButtonData = new PushButtonData(
                    "QPK_CreateWorksets",
                    "Create\nWorksets",
                    assemblyPath,
                    "QPK_Keynote_Manager.WorksetPreloader"
                );

                worksetButtonData.ToolTip = "Create predefined worksets in the active model";
                worksetButtonData.LongDescription = "Automatically creates a standard set of worksets for the current project. If worksharing is not enabled, it will be enabled first.";

                // Optional: Add an icon for the worksets button
                string worksetImg = Path.Combine(folder, "Resources", "QPKKNM_32.png"); // Preload Worksets icon
                if (File.Exists(worksetImg))
                {
                    worksetButtonData.LargeImage = LoadPng(worksetImg);
                }

                panel.AddItem(worksetButtonData);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("QPK Keynote Manager - Startup Error", ex.ToString());
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }

        private BitmapImage LoadPng(string path)
        {
            BitmapImage img = new BitmapImage();
            img.BeginInit();
            img.UriSource = new Uri(path, UriKind.Absolute);
            img.EndInit();
            return img;
        }
    }
}