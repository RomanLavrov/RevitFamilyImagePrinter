using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using View = Autodesk.Revit.DB.View;

namespace RevitFamilyImagePrinter.Commands
{
    [Transaction(TransactionMode.Manual)]
    class Print2DFolder : IExternalCommand
    {
        private const int windowHeightOffset = 40;
        private const int windowWidthOffset = 40;
        public int UserScale { get; set; }
        public int UserImageSize { get; set; }
        //TODO Add FBD for images output
        string imagePath = "D:\\TypeImages\\";
        IList<ElementId> views = new List<ElementId>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;           

            ShowOptions();
            //TODO Add FBD for rvt projects as input
            var fileList = Directory.GetFiles("D:\\Test");
            foreach (var item in fileList)
            {
                if (!item.Contains("000"))
                {
                    try
                    {
                        Debug.Print(item);
                        uidoc = commandData.Application.OpenAndActivateDocument(item);
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
                        foreach (var uiView in uiviews)
                        {
                            uiView.ZoomToFit();
                            //TODO - set zoom from user settings
                            uiView.Zoom(0.95);
                            
                            uidoc.RefreshActiveView();
                        }

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

                        uidoc = commandData.Application.OpenAndActivateDocument("D:\\Empty.rvt");
                        doc.Close();
                    }
                    catch
                    {
                        Debug.Print(item);
                    }
                }
            }

            Rename2DImages();

            return Result.Succeeded;
        }

        private void PrintImage(Document doc)
        {
            int indexDot = doc.Title.IndexOf('.');
            var name = doc.Title.Substring(0, indexDot);
            var tempFile = imagePath + name + ".png";

            IList<ElementId> views = new List<ElementId>();
            views.Add(doc.ActiveView.Id);

            var exportOptions = new ImageExportOptions
            {
                ViewName = "temp",
                FilePath = tempFile,
                FitDirection = FitDirectionType.Vertical,
                HLRandWFViewsFileType = ImageFileType.PNG,
                ImageResolution = ImageResolution.DPI_300,
                ShouldCreateWebSite = false,
                PixelSize = UserImageSize,
                ZoomType = ZoomFitType.FitToPage
            };

            if (views.Count > 0)
            {
                exportOptions.SetViewsAndSheets(views);
                //exportOptions.ExportRange = ExportRange.SetOfViews;
                exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
            }
            else
            {
                exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
            }           

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
                Height = options.Height + windowHeightOffset,
                Width = options.Width + windowWidthOffset,
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

        private void Rename2DImages()
        {
            var picturesList = Directory.GetFiles(imagePath);
            foreach (var item in picturesList)
            {
                if (item.IndexOf("- Structural Plan - Level 1") > 0)
                {
                    try
                    {
                        var index = item.IndexOf("- Structural Plan - Level 1");
                        File.Move(item, item.Substring(0, index) + ".png");
                    }
                    catch { };
                }
            }
        }
    }
}
