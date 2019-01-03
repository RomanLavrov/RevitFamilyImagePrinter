#region Namespaces
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitFamilyImagePrinter.Infrastructure;

#endregion

namespace RevitFamilyImagePrinter
{
    class App : IExternalApplication
    {
		public static string Version { get; private set; }
		public static string DefaultFolder => Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
		public static string DefaultProject => Path.Combine(DefaultFolder, "Empty.rvt");
	    public static Logger Logger = Logger.GetLogger();

	    public Result OnStartup(UIControlledApplication a)
		{
			Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
			Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture;

			ControlledApplication c = a.ControlledApplication;
			Version = c.VersionNumber;
			Logger.WriteLine($"Revit Version -> {Version}");

			string tabName = "Image Printer";
            a.CreateRibbonTab(tabName);

			var assembly = Assembly.GetExecutingAssembly().Location;

			PushButtonData buttonPrint2DSingle = new PushButtonData("Print 2D", " 2D Single ", assembly, "RevitFamilyImagePrinter.Commands.Print2D");
			buttonPrint2DSingle.ToolTip = "Create 2D family image from current view";

			PushButtonData buttonPrint2DMulti = new PushButtonData("Print 2D Folder", " 2D Folder ", assembly, "RevitFamilyImagePrinter.Commands.Print2DFolder");
			buttonPrint2DMulti.ToolTip = "Create 2D family images from selected folder";

			PushButtonData buttonPrint3DSingle = new PushButtonData("Print 3D", " 3D Single ", assembly, "RevitFamilyImagePrinter.Commands.Print3D");
			buttonPrint3DSingle.ToolTip = "Create 3D family image from current view";

			PushButtonData buttonPrint3DMulti = new PushButtonData("Print 3D Folder", " 3D Folder ", assembly, "RevitFamilyImagePrinter.Commands.Print3DFolder");
			buttonPrint3DMulti.ToolTip = "Create 3D family image from selected folder";

			PushButtonData buttonRemoveEmptyFamilies = new PushButtonData("Remove empty families", "Remove Families", assembly, "RevitFamilyImagePrinter.Commands.RemoveEmptyFamilies");
			buttonRemoveEmptyFamilies.ToolTip = "Remove all families without instances.";

			PushButtonData buttonProjectCreator = new PushButtonData("Project creator", "Project creator", assembly, "RevitFamilyImagePrinter.Commands.ProjectCreator");
			buttonProjectCreator.ToolTip = "Create .rvt files from .rfa files";

			PushButtonData buttonPrintView = new PushButtonData("Print current view", "Print view", assembly, "RevitFamilyImagePrinter.Commands.PrintView");
			buttonPrintView.ToolTip = "Create a screenshot of current view and save it as a picture";

			RibbonPanel printPanel = a.CreateRibbonPanel(tabName, "Family Image Printer");
			printPanel.AddItem(buttonPrint2DSingle);
			printPanel.AddItem(buttonPrint2DMulti);
			printPanel.AddItem(buttonPrint3DSingle);
			printPanel.AddItem(buttonPrint3DMulti);
			printPanel.AddSeparator();
			printPanel.AddItem(buttonPrintView);
			//printPanel.AddItem(buttonRemoveEmptyFamilies);
			//printPanel.AddItem(buttonProjectCreator);

			return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
			if(File.Exists(DefaultProject) && RevitPrintHelper.IsFileAccessible(DefaultProject))
				File.Delete(DefaultProject);
			Logger.EndLogSession();
			//to clean folder, from which files were printed
            return Result.Succeeded;
        }
    }
}
