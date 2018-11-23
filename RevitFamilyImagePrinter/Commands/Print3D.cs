using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using View = System.Windows.Forms.View;

namespace RevitFamilyImagePrinter.Commands
{
    [Transaction(TransactionMode.Manual)]
    class Print3D : IExternalCommand
    {
        public int UserScale { get; set; }
        public int UserImageSize { get; set; }
        string imagePath = "D:\\";
        private Document doc;
        private UIApplication uiapp;
        private UIDocument uidoc;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;

            RevitCommandId commandId = RevitCommandId.LookupPostableCommandId(PostableCommand.Default3DView);

            if (commandData.Application.CanPostCommand(commandId))
            {
                commandData.Application.PostCommand(commandId);
            }

            commandData.Application.ViewActivated += SetViewParameters;

            return Result.Succeeded;
        }

        private void PrintImage(Document doc)
        {
            int indexDot = doc.Title.IndexOf('.');
            var name = doc.Title.Substring(0, indexDot);

            var tempFile = imagePath + name + ".png";
            IList<ElementId> views = new List<ElementId>();

            views.Add(uidoc.ActiveView.Id);

            var exportOptions = new ImageExportOptions
            {
                FilePath = tempFile,
               // FitDirection = FitDirectionType.Vertical,
                HLRandWFViewsFileType = ImageFileType.PNG,
                ImageResolution = ImageResolution.DPI_300,
                ShouldCreateWebSite = false,
                PixelSize = 64,// UserImageSize,
               // ZoomType = ZoomFitType.Zoom
            };

            if (views.Count > 0)
            {
                exportOptions.SetViewsAndSheets(views);
                exportOptions.ExportRange = ExportRange.SetOfViews;
            }
            else
            {
                exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
            }

            //exportOptions.ZoomType = ZoomFitType.FitToPage;
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
                UserScale = options.UserScale;
                UserImageSize = options.UserImageSize;
            }
        }

        private void SetViewParameters(object sender, ViewActivatedEventArgs args)
        {
            View3D view = null;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(View3D));
            foreach (View3D VARIABLE in collector)
            {
                if (VARIABLE != null)
                {
                    view = VARIABLE;

                    using (Transaction transaction = new Transaction(doc))
                    {
                        transaction.Start("SetView");
                        uidoc.ActiveView.DetailLevel = ViewDetailLevel.Fine;
                        uidoc.ActiveView.Scale = 10;
                        transaction.Commit();
                    }
                }
            }

            using (Transaction transaction = new Transaction(doc))
            {
                transaction.Start("Print");
                PrintImage(doc);
                transaction.Commit();
            }

        }
    }
}
