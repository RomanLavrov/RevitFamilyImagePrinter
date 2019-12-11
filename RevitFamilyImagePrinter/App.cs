using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI;
using RevitFamilyImagePrinter.Infrastructure;
using RevitFamilyImagePrinter.Properties;
using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Media.Imaging;

namespace RevitFamilyImagePrinter
{
	public class App : IExternalApplication
	{
		public static string Version { get; private set; }
		public static string Language { get; set; }
		public static string DefaultFolder { get; private set; }
		public static string DefaultProject { get; private set; }
		public static Logger Logger { get; private set; }
		public static Translator Translator { get; private set; }

		public Result OnStartup(UIControlledApplication a)
		{
			Initialize(a.ControlledApplication);
			if (!IsCompatibleVersion(a.ControlledApplication.VersionNumber)) return Result.Failed;

			string tabName = Translator.GetValue(Translator.Keys.tabName);
			a.CreateRibbonTab(tabName);

			var assembly = Assembly.GetExecutingAssembly().Location;

			PushButtonData buttonPrint2DSingle = new PushButtonData("Print 2D",
				Translator.GetValue(Translator.Keys.buttonPrint2DSingle_Name), assembly,
				"RevitFamilyImagePrinter.Commands.Print2D")
			{
				ToolTip = Translator.GetValue(Translator.Keys.buttonPrint2DSingle_ToolTip),

				LargeImage = GetImage(Resources._2D_Single.GetHbitmap()),

			};
			ContextualHelp contextualHelp = new ContextualHelp(ContextualHelpType.Url, "https://help.building360.ch/ImagePrinter/");
			buttonPrint2DSingle.SetContextualHelp(contextualHelp);


			PushButtonData buttonPrint2DMulti = new PushButtonData("Print 2D Folder", Translator.GetValue(Translator.Keys.buttonPrint2DMulti_Name), assembly,
				"RevitFamilyImagePrinter.Commands.Print2DFolder")
			{
				ToolTip = Translator.GetValue(Translator.Keys.buttonPrint2DMulti_ToolTip),
				LargeImage = GetImage(Resources._2D_Folder.GetHbitmap())
			};
			buttonPrint2DMulti.SetContextualHelp(contextualHelp);

			PushButtonData buttonPrint3DSingle = new PushButtonData("Print 3D", Translator.GetValue(Translator.Keys.buttonPrint3DSingle_Name), assembly,
				"RevitFamilyImagePrinter.Commands.Print3D")
			{
				ToolTip = Translator.GetValue(Translator.Keys.buttonPrint3DSingle_ToolTip),
				LargeImage = GetImage(Resources._3D_Single.GetHbitmap())
			};
			buttonPrint3DSingle.SetContextualHelp(contextualHelp);

			PushButtonData buttonPrint3DMulti = new PushButtonData("Print 3D Folder", Translator.GetValue(Translator.Keys.buttonPrint3DMulti_Name), assembly,
				"RevitFamilyImagePrinter.Commands.Print3DFolder")
			{
				ToolTip = Translator.GetValue(Translator.Keys.buttonPrint3DMulti_ToolTip),
				LargeImage = GetImage(Resources._3D_Folder.GetHbitmap())
			};
			buttonPrint3DMulti.SetContextualHelp(contextualHelp);

			PushButtonData buttonPrintView = new PushButtonData("Print current view", Translator.GetValue(Translator.Keys.buttonPrintView_Name), assembly,
				"RevitFamilyImagePrinter.Commands.PrintView")
			{
				ToolTip = Translator.GetValue(Translator.Keys.buttonPrintView_ToolTip),
				LargeImage = GetImage(Resources.viewexport.GetHbitmap())
			};
			buttonPrintView.SetContextualHelp(contextualHelp);

			PushButtonData buttonLink = new PushButtonData("building360.ch", "building360.ch", assembly,
				"RevitFamilyImagePrinter.Commands.Link")
			{
				ToolTip = Translator.GetValue(Translator.Keys.buttonLink_ToolTip),
				LargeImage = GetImage(Resources.logo_small.GetHbitmap())
			};

			PushButtonData buttonRemoveEmptyFamilies = new PushButtonData("Remove empty families", "Remove Families", assembly, "RevitFamilyImagePrinter.Commands.RemoveEmptyFamilies");
			buttonRemoveEmptyFamilies.ToolTip = "Remove all families without instances.";

			PushButtonData buttonProjectCreator = new PushButtonData("Project creator", "Project creator", assembly, "RevitFamilyImagePrinter.Commands.ProjectCreator");
			buttonProjectCreator.ToolTip = "Create .rvt files from .rfa files";

			RibbonPanel printPanel = a.CreateRibbonPanel(tabName, Translator.GetValue(Translator.Keys.tabTitle));
			printPanel.AddItem(buttonPrint2DSingle);
			printPanel.AddItem(buttonPrint2DMulti);
			printPanel.AddItem(buttonPrint3DSingle);
			printPanel.AddItem(buttonPrint3DMulti);
			printPanel.AddSeparator();
			printPanel.AddItem(buttonPrintView);
			printPanel.AddSeparator();
			printPanel.AddItem(buttonLink);

			return Result.Succeeded;
		}

		private bool IsCompatibleVersion(string version)
		{
			int numVersion = -1;
			int.TryParse(version, out numVersion);
			if (numVersion < 2018)
			{
				new Autodesk.Revit.UI.TaskDialog(Translator.GetValue(Translator.Keys.warningMessageTitle))
				{
					TitleAutoPrefix = false,
					MainIcon = Autodesk.Revit.UI.TaskDialogIcon.TaskDialogIconWarning,
					MainContent = Translator.GetValue(Translator.Keys.warningMessageNotCompatible)
				}.Show();
				return false;
			}
			return true;
		}

		public Result OnShutdown(UIControlledApplication a)
		{
			if (File.Exists(DefaultProject) && PrintHelper.IsFileAccessible(DefaultProject))
				File.Delete(DefaultProject);
			Logger.EndLogSession();
			return Result.Succeeded;
		}

		private static BitmapSource GetImage(IntPtr bm)
		{
			BitmapSource bmSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
				bm,
				IntPtr.Zero,
				Int32Rect.Empty,
				BitmapSizeOptions.FromEmptyOptions());
			return bmSource;
		}

		private static void Initialize(ControlledApplication cApp)
		{
			Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
			Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture;

			string myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			DefaultFolder = Path.Combine(myDocuments, "Building360");
			if (!Directory.Exists(DefaultFolder))
				Directory.CreateDirectory(DefaultFolder);
			DefaultProject = Path.Combine(DefaultFolder, "Empty.rvt");
			Logger = Logger.GetLogger();

			Version = cApp.VersionNumber;
			Language = cApp.Language.ToString();

			string loggerCompInfo = string.Empty;
			loggerCompInfo += $"\nImagePrinter version: {Assembly.GetExecutingAssembly().GetName().Version.ToString(3)}";
			loggerCompInfo += $"\nRevit Version: {cApp.VersionNumber}";
			loggerCompInfo += $"\nRevit Language: {Language}";

			Logger.WriteLine(loggerCompInfo, false);
			Translator = new Translator(Language);
		}
	}
}
