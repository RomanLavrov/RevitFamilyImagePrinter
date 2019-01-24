using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using RevitFamilyImagePrinter.Infrastructure;
using System.Linq;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print2DFolder : IExternalCommand
	{
		#region Properties
		public UserImageValues UserValues { get; set; } = new UserImageValues();
		public DirectoryInfo UserFolderFrom { get; set; }
		public DirectoryInfo UserFolderTo { get; set; }
		#endregion

		#region Variables
		private UIDocument _uiDoc;
		private readonly Logger _logger = App.Logger;
		#endregion

		#region Constants
		private const int windowHeightOffset = 40;
		private const int windowWidthOffset = 10;
		private const int maxSizeLength = 2097152;
		#endregion

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			PrintProgressHelper progressHelper = null;
			try
			{
				UIApplication uiApp = commandData.Application;
				_uiDoc = uiApp.ActiveUIDocument;
				var initProjectPath = _uiDoc.Document.PathName;
				PrintHelper.CreateEmptyProject(uiApp.Application);

				DirectoryInfo familiesFolder =
					PrintHelper.SelectFolderDialog($"{App.Translator.GetValue(Translator.Keys.folderDialogFromTitle)}");
				if (familiesFolder == null)
					return Result.Cancelled;

				UserFolderTo = 
					PrintHelper.SelectFolderDialog($"{App.Translator.GetValue(Translator.Keys.folderDialogToTitle)}");
				if (UserFolderTo == null)
					return Result.Cancelled;

				UserValues =
					PrintHelper.ShowOptionsDialog(_uiDoc, windowHeightOffset, windowWidthOffset, false, false);
				if (UserValues == null)
					return Result.Failed;

				var families = GetFamilyFilesFromFolder(familiesFolder);
				if (families == null)
					return Result.Failed;

				progressHelper = new PrintProgressHelper(familiesFolder,
					$"{App.Translator.GetValue(Translator.Keys.textBlockProcessCreatingProjects)}");
				progressHelper.Show();
				progressHelper.SubscribeOnLoadedFamily(uiApp);
				progressHelper.SetProgressBarMaximum(families.Count);

				UserFolderFrom = new DirectoryInfo(Path.Combine(familiesFolder.FullName,
					App.Translator.GetValue(Translator.Keys.folderProjectsName)));
				foreach (var i in families)
				{
					PathData pathData = new PathData()
					{
						FamilyPath = i.FullName,
						ProjectsPath = UserFolderFrom.FullName,
						ImagesPath = UserFolderTo.FullName
					};
					ProjectHelper.CreateProjectsFromFamily(_uiDoc, pathData, UserValues);
				}

				if (!string.IsNullOrEmpty(initProjectPath) && File.Exists(initProjectPath))
					_uiDoc = PrintHelper.OpenDocument(_uiDoc, initProjectPath);
				else
					_uiDoc = PrintHelper.OpenDocument(_uiDoc, App.DefaultProject);
			}
			catch (Exception exc)
			{
				PrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessage2dFolderPrinting)}", _logger);
				return Result.Failed;
			}
			finally
			{
				progressHelper?.Close();
			}
			return Result.Succeeded;
		}

		private List<FileInfo> GetFamilyFilesFromFolder(DirectoryInfo familiesFolder)
		{
			try
			{
				return ProjectHelper.GetFamilyFilesFromFolder(familiesFolder);
			}
			catch (Exception exc)
			{
				PrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageFamiliesRetrieving)}", App.Logger);
				return null;
			}
		}
	}
}
