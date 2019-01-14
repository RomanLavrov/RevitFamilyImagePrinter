using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using RevitFamilyImagePrinter.Infrastructure;
using View = Autodesk.Revit.DB.View;
using System.Linq;
using Autodesk.Revit.Exceptions;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print2DFolder : IExternalCommand
	{

		#region Properties
		public UserImageValues UserValues { get; set; } = new UserImageValues();
		public DirectoryInfo UserFolderFrom { get; set; } = new DirectoryInfo(@"D:\WebTypes\TestTypes");
		public DirectoryInfo UserFolderTo { get; set; } = new DirectoryInfo(@"D:\TypeImages");
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
				RevitPrintHelper.CreateEmptyProject(uiApp.Application);

				DirectoryInfo familiesFolder =
					RevitPrintHelper.SelectFolderDialog($"{App.Translator.GetValue(Translator.Keys.folderDialogFromTitle)}");
				if (familiesFolder == null)
					return Result.Cancelled;

				UserFolderTo = 
					RevitPrintHelper.SelectFolderDialog($"{App.Translator.GetValue(Translator.Keys.folderDialogToTitle)}");
				if (UserFolderTo == null)
					return Result.Cancelled;

				UserValues =
					RevitPrintHelper.ShowOptionsDialog(_uiDoc, windowHeightOffset, windowWidthOffset, false, false);
				if (UserValues == null)
					return Result.Failed;

				progressHelper = new PrintProgressHelper(familiesFolder,
					$"{App.Translator.GetValue(Translator.Keys.textBlockProcessCreatingProjects)}");
				progressHelper.Show();
				progressHelper.SubscribeOnLoadedFamily(uiApp);
				progressHelper.SetProgressBarMaximum(familiesFolder.GetFiles().Count(x => x.Extension.Equals(".rfa")));

				if (!CreateProjects(commandData, elements, familiesFolder))
					return Result.Failed;

				var fileList = Directory.GetFiles(UserFolderFrom.FullName);
				progressHelper.SetProgressText($"{App.Translator.GetValue(Translator.Keys.textBlockProcessPreparingPrinting)}");
				progressHelper.SetProgressBarMaximum(UserFolderFrom.GetFiles().Count(x => x.Extension.Equals(".rvt")));
				progressHelper.SubscribeOnViewActivated(uiApp);

				foreach (var item in fileList)
				{
					try
					{
						FileInfo info = new FileInfo(item);
						if (!info.Extension.Equals(".rvt"))
							continue;
						RevitPrintHelper.OpenDocument(_uiDoc, App.DefaultProject);
						_uiDoc = uiApp.OpenAndActivateDocument(item);

						if (info.Length > maxSizeLength)
						{
							RevitPrintHelper.RemoveEmptyFamilies(_uiDoc);
						}

						RevitPrintHelper.SetActive2DView(_uiDoc);
						ViewChangesCommit();
						PrintCommit(_uiDoc.Document);
					}
					catch (CorruptModelException exc)
					{
						RevitPrintHelper.ProcessError(exc,
							$"{exc.Message}{Environment.NewLine}{new FileInfo(item).Name}", _logger);
					}
					catch (Exception exc)
					{
						RevitPrintHelper.ProcessError(exc,
							$"{App.Translator.GetValue(Translator.Keys.errorMessage2dFolderPrintingCycle)}", _logger);
					}
				}

				if (!string.IsNullOrEmpty(initProjectPath) && File.Exists(initProjectPath))
					_uiDoc = RevitPrintHelper.OpenDocument(_uiDoc, initProjectPath);
				else
					_uiDoc = RevitPrintHelper.OpenDocument(_uiDoc, App.DefaultProject);
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessage2dFolderPrinting)}", _logger);
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
				RevitPrintHelper.View2DChangesCommit(_uiDoc, UserValues);
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageViewCorrecting)}", _logger);
			}
		}

		private void PrintCommit(Document _doc)
		{
			try
			{
				string initialName = RevitPrintHelper.GetFileName(_doc);
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
