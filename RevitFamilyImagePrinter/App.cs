#region Namespaces
using System;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace RevitFamilyImagePrinter
{
    class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            string tabName = "Image Printer";
            a.CreateRibbonTab(tabName);

            var assembly = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttontPrint2DSingle = new PushButtonData("Print 2D", " 2D Single ", assembly, "RevitFamilyImagePrinter.Commands.Print2D");
            buttontPrint2DSingle.ToolTip = "Create 2D family image from current view";

            PushButtonData buttontPrint2DMulti = new PushButtonData("Print 2D Folder", " 2D Folder ", assembly, "RevitFamilyImagePrinter.Commands.Print2DFolder");
            buttontPrint2DMulti.ToolTip = "Create 2D family images from selected folder";

            PushButtonData buttontPrint3DSingle = new PushButtonData("Print 3D", " 3D Single ", assembly, "RevitFamilyImagePrinter.Commands.Print3D");
            buttontPrint3DSingle.ToolTip = "Create 3D family image from current view";

            PushButtonData buttontPrint3DMulti = new PushButtonData("Print 3D Folder", " 3D Folder ", assembly, "RevitFamilyImagePrinter.Commands.Print3DFolder");
            buttontPrint3DMulti.ToolTip = "Create 3D family image from selected folder";

            RibbonPanel printPanel = a.CreateRibbonPanel(tabName, "Family Image Printer");
            printPanel.AddItem(buttontPrint2DSingle);
            printPanel.AddItem(buttontPrint2DMulti);
            printPanel.AddItem(buttontPrint3DSingle);
            printPanel.AddItem(buttontPrint3DMulti);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
