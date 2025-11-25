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

                PushButtonData buttonData = new PushButtonData(
                    "QPK_MainWindow",
                    "Open QPK Tool",
                    assemblyPath,
                    "QPK_Keynote_Manager.ShowMainWindowCommand"
                );

                // Optional image
                string img = Path.Combine(folder, "QPKIcon.png");
                if (File.Exists(img))
                {
                    buttonData.LargeImage = LoadPng(img);
                }

                panel.AddItem(buttonData);

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
