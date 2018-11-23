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

namespace RevitFamilyImagePrinter.Commands
{
    [Transaction(TransactionMode.Manual)]
    class Print3DFolder : IExternalCommand
    {
        public int UserScale { get; set; }
        public int UserImageSize { get; set; }
        string imagePath = "D:\\TypeImages3D\\";
        private Document doc;
        private UIApplication uiapp;
        private UIDocument uidoc;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;
            
            var fileList = Directory.GetFiles("D:\\Test");
            foreach (var item in fileList)
            {
                try
                {
                    uidoc = commandData.Application.OpenAndActivateDocument(item);
                    var collectorF = new FilteredElementCollector(this.doc);
                    var viewFamilyType = collectorF.OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                        .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);
                    Document doc = uidoc.Document;
                    View3D view3d = doc.ActiveView as View3D;

                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("Add view");
                        view3d = View3D.CreateIsometric(doc, viewFamilyType.Id);
                        trans.Commit();
                    }
                    uidoc.ActiveView = view3d;
                    
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfClass(typeof(View3D));

                    IList<UIView> uIViews = uidoc.GetOpenUIViews();
                    foreach (var uiview in uIViews)
                    {
                        uiview.ZoomToFit();
                        uiview.Zoom(0.95);
                        uidoc.RefreshActiveView();
                    }
                    foreach (View3D VARIABLE in collector)
                    {
                        if (VARIABLE != null)
                        {
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

                    commandData.Application.ViewActivated += SetViewParameters;
                    uidoc = commandData.Application.OpenAndActivateDocument("D:\\Empty.rvt");
                    doc.Close();
                }
                catch { }              
            }

            var picturesList =  Directory.GetFiles(imagePath);
            foreach (var item in picturesList)
            {                    
                if (item.IndexOf("- 3D View - 3D View") > 0)
                {
                    try
                    {
                        var index = item.IndexOf("- 3D View - 3D View");                           
                        System.IO.File.Move(item, item.Substring(0, index) + ".png");
                    }
                    catch{};
                }
            }


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
                 FitDirection = FitDirectionType.Vertical,
                HLRandWFViewsFileType = ImageFileType.PNG,
                ImageResolution = ImageResolution.DPI_300,
                ShouldCreateWebSite = false,
                PixelSize = 460, // UserImageSize,
                ZoomType = ZoomFitType.Zoom
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

            exportOptions.ZoomType = ZoomFitType.FitToPage;
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
                    using (Transaction transaction = new Transaction(doc))
                    {
                        transaction.Start("SetView");
                        uidoc.ActiveView.DetailLevel = ViewDetailLevel.Fine;
                        uidoc.ActiveView.Scale = 50;
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
