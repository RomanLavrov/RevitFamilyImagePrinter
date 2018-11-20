using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using MessageBox = System.Windows.MessageBox;
using View = Autodesk.Revit.DB.View;

namespace RevitFamilyImagePrinter.Commands
{
    [Transaction(TransactionMode.Manual)]
    class Print2D : IExternalCommand
    {
        public int UserScale { get; set; }
        public int UserImageSize { get; set; }
        // TODO - Get Folder from FBD
        string imagePath = "D:\\TypeImages\\";
        IList<ElementId> views = new List<ElementId>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
            viewCollector.OfClass(typeof(View));
            
            foreach (Element viewElement in viewCollector)
            {
                View view = (View)viewElement;

                if (view.Name.Equals("Level 1") && view.ViewType == ViewType.EngineeringPlan)
                {
                    views.Add(view.Id);
                    uidoc.ActiveView = view;
                }
            }

            IList<UIView> uiviews = uidoc.GetOpenUIViews();
            foreach (var item in uiviews)
            {
                //item.ZoomToFit();
                //item.Zoom(0.95);
                uidoc.RefreshActiveView();
            }

            ShowOptions();

            //TODO - Move transaction to separate method
            using (Transaction transaction = new Transaction(doc))
            {
                transaction.Start("SetView");
                doc.ActiveView.DetailLevel = ViewDetailLevel.Medium;
                doc.ActiveView.Scale = UserScale;
                transaction.Commit();
            }

            using (Transaction transaction = new Transaction(doc))
            {
                transaction.Start("Print");
                PrintImage(doc);
                transaction.Commit();
            }

            return Result.Succeeded;
        }

        private void PrintImage(Document doc)
        {
            int indexDot = doc.Title.IndexOf('.');
            var name = doc.Title.Substring(0, indexDot);
            //TODO - use Path.Combine
            var tempFile = imagePath + name + ".png";

            IList<ElementId> views = new List<ElementId>();
            views.Add(doc.ActiveView.Id);

            //TODO - Get User settings for this options
            var exportOptions = new ImageExportOptions
            {
                ViewName = "temp",
                FilePath = tempFile,
                FitDirection = FitDirectionType.Vertical,
                HLRandWFViewsFileType = ImageFileType.PNG,
                ImageResolution = ImageResolution.DPI_300,
                ShouldCreateWebSite = false,
                PixelSize = UserImageSize
                /*,
                ZoomType = ZoomFitType.FitToPage*/
            };

            if (views.Count > 0)
            {
                exportOptions.SetViewsAndSheets(views);
                exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
            }
            else
            {
                exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
            }

            exportOptions.ViewName = "temp";

            if (ImageExportOptions.IsValidFileName(tempFile))
            {
                doc.ExportImage(exportOptions);
            }
        }

        private void ShowOptions()
        {
            SinglePrintOptions options = new SinglePrintOptions();
            Window window = new Window
            {
                Height = 180,
                Width = 260,
                Title = "Image Print Settings",
                Content = options,
                Background = System.Windows.Media.Brushes.WhiteSmoke,
                WindowStyle = WindowStyle.ToolWindow,
                Name = "Options",
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen

            };

            window.ShowDialog();

            if (window.DialogResult == true)
            {
                UserScale = options.userScale;
                UserImageSize = options.userImageSize;
            }
        }
    }
}
