using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using RevitFamilyImagePrinter.Infrastructure;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print3DFolder : IExternalCommand
	{
		#region Properties

		public UserImageValues UserValues { get; set; } = new UserImageValues();
		public DirectoryInfo UserFolderFrom { get; set; } = new DirectoryInfo(@"D:\WebTypes\TestTypes");
		public DirectoryInfo UserFolderTo { get; set; } = new DirectoryInfo(@"D:\TypeImages");

		#endregion

		#region Variables

		private UIApplication _uiApp;
		private UIDocument _uiDoc;
		private readonly Logger _logger = App.Logger;

		#endregion

		#region Constants

		private const int windowHeightOffset = 40;
		private const int windowWidthOffset = 10;
		private const int maxSizeLength = 2097152;
		private readonly string endl = Environment.NewLine;

		#endregion

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			PrintProgressHelper progressHelper = null;
			try
			{
				_uiApp = commandData.Application;
				_uiDoc = _uiApp.ActiveUIDocument;
				var initProjectPath = _uiDoc.Document.PathName;
				RevitPrintHelper.CreateEmptyProject(commandData.Application.Application);

				DirectoryInfo familiesFolder =
					RevitPrintHelper.SelectFolderDialog($"{App.Translator.GetValue(Translator.Keys.folderDialogFromTitle)}");
				if (familiesFolder == null)
					return Result.Cancelled;
				UserFolderTo = RevitPrintHelper.SelectFolderDialog($"{App.Translator.GetValue(Translator.Keys.folderDialogToTitle)}");
				if (UserFolderTo == null)
					return Result.Cancelled;

				UserValues =
					RevitPrintHelper.ShowOptionsDialog(_uiDoc, windowHeightOffset, windowWidthOffset, false, false);
				if (UserValues == null)
					return Result.Failed;

				progressHelper = new PrintProgressHelper(familiesFolder,
					$"{App.Translator.GetValue(Translator.Keys.textBlockProcessCreatingProjects)}");
				progressHelper.Show(true);
				progressHelper.SubscribeOnLoadedFamily(_uiApp);
				progressHelper.SetProgressBarMaximum(familiesFolder.GetFiles().Count(x => x.Extension.Equals(".rfa")));

				if (!CreateProjects(commandData, elements, familiesFolder))
					return Result.Failed;

				var fileList = Directory.GetFiles(UserFolderFrom.FullName);
				progressHelper.SetProgressText($"{App.Translator.GetValue(Translator.Keys.textBlockProcessPreparingPrinting)}");
				progressHelper.SetProgressBarMaximum(UserFolderFrom.GetFiles().Count(x => x.Extension.Equals(".rvt")));
				progressHelper.SubscribeOnViewActivated(_uiApp, true);

				//int createdImages = 0;
				
				foreach (var item in fileList)
				{
					try
					{
						FileInfo fileInfo = new FileInfo(item);
						if (!fileInfo.Extension.Equals(".rvt"))
							continue;
						RevitPrintHelper.OpenDocument(_uiDoc, App.DefaultProject);
						_uiDoc = commandData.Application.OpenAndActivateDocument(item);
						if (fileInfo.Length > maxSizeLength)
							RevitPrintHelper.RemoveEmptyFamilies(_uiDoc);
					    RevitPrintHelper.SetActive3DView(_uiDoc);
						ViewChangesCommit();
						PrintCommit(_uiDoc.Document);
					}
					catch (CorruptModelException exc)
					{
						RevitPrintHelper.ProcessError(exc,
							$"{exc.Message}{Environment.NewLine}{new FileInfo(item).Name}", _logger, false);
					}
					catch (Exception exc)
					{
						RevitPrintHelper.ProcessError(exc,
							$"{App.Translator.GetValue(Translator.Keys.errorMessage3dFolderPrintingCycle)}", _logger, false);
					}
				}

				//RevitPrintHelper.CheckImagesAmount(UserFolderFrom, createdImages);

				if (!string.IsNullOrEmpty(initProjectPath) && File.Exists(initProjectPath))
					_uiDoc = RevitPrintHelper.OpenDocument(_uiDoc, initProjectPath);
				else
					_uiDoc = RevitPrintHelper.OpenDocument(_uiDoc, App.DefaultProject);
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessage3dFolderPrinting)}", _logger);

				return Result.Failed;
			}
			finally
			{
				progressHelper?.Close();
			}
			return Result.Succeeded;
		}

	    private bool CreateProjects(ExternalCommandData commandData, ElementSet elements, DirectoryInfo familiesFolder)
		{
			ProjectCreator creator = new ProjectCreator()
			{
				FamiliesFolder = familiesFolder,
				UserValues = this.UserValues
			};
			string tmp = string.Empty;
			var result = creator.Execute(commandData, ref tmp, elements);
			if (result != Result.Succeeded)
				return false;
			UserFolderFrom = creator.ProjectsFolder;
			return true;
		}

		public void ViewChangesCommit()
		{
			try
			{
				RevitPrintHelper.View3DChangesCommit(_uiDoc, UserValues);
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageViewCorrecting)}", _logger);
			}
		}

		private void PrintCommit(Document doc)
		{
			try
			{
				string initialName = RevitPrintHelper.GetFileName(doc);
				string filePath = Path.Combine(UserFolderTo.FullName, initialName);
				RevitPrintHelper.PrintImageTransaction(_uiDoc, UserValues, filePath, true);
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageViewPrinting)}", _logger);
			}
		}
	}
}
