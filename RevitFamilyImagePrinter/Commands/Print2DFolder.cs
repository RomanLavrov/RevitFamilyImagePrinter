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
					RevitPrintHelper.SelectFolderDialog("Select folder with needed families to be printed");
				if (familiesFolder == null)
					return Result.Cancelled;

				UserFolderTo = RevitPrintHelper.SelectFolderDialog("Select folder where to save printed files");
				if (UserFolderTo == null)
					return Result.Cancelled;

				UserValues =
					RevitPrintHelper.ShowOptionsDialog(_uiDoc, windowHeightOffset, windowWidthOffset, false, false);
				if (UserValues == null)
					return Result.Failed;

				progressHelper = new PrintProgressHelper(familiesFolder,
					"Creating .rvt projects from .rfa files, this may take a time...");
				progressHelper.Show();
				progressHelper.SubscribeOnLoadedFamily(uiApp);
				progressHelper.SetProgressBarMaximum(familiesFolder.GetFiles().Count(x => x.Extension.Equals(".rfa")));

				if (!CreateProjects(commandData, elements, familiesFolder))
					return Result.Failed;

				var fileList = Directory.GetFiles(UserFolderFrom.FullName);
				progressHelper.SetProgressText("Preparation for printing...");
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
						if (!exc.Message.Equals("The model was saved by a later version of Revit.")) throw;
						string errorMessage = $"{exc.Message}{Environment.NewLine}{new FileInfo(item).Name}";
						new TaskDialog("Error")
						{
							TitleAutoPrefix = false,
							MainIcon = TaskDialogIcon.TaskDialogIconError,
							MainContent = errorMessage
						}.Show();
						_logger.WriteLine($"### ERROR ### - {errorMessage}\n{exc.Message}\n{exc.StackTrace}");
					}
					catch (Exception exc)
					{
						string errorMessage =
							"Error occured in cycle during Print2DFolder command execution";
						new TaskDialog("Error")
						{
							TitleAutoPrefix = false,
							MainIcon = TaskDialogIcon.TaskDialogIconError,
							MainContent = errorMessage
						}.Show();
						_logger.WriteLine($"### ERROR ### - {errorMessage}\n{exc.Message}\n{exc.StackTrace}");
					}
				}

				if (!string.IsNullOrEmpty(initProjectPath) && File.Exists(initProjectPath))
					_uiDoc = RevitPrintHelper.OpenDocument(_uiDoc, initProjectPath);
				else
					_uiDoc = RevitPrintHelper.OpenDocument(_uiDoc, App.DefaultProject);
			}
			catch (Exception exc)
			{
				string errorMessage = "Error occured during Print2DFolder command execution";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainIcon = TaskDialogIcon.TaskDialogIconError,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ### - {errorMessage}\n{exc.Message}\n{exc.StackTrace}");

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
				string errorMessage = "Error occured during current view correction";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainIcon = TaskDialogIcon.TaskDialogIconError,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ### - {errorMessage}\n{exc.Message}\n{exc.StackTrace}");
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
				string errorMessage = "Error occured during printing current view";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainIcon = TaskDialogIcon.TaskDialogIconError,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ### - {errorMessage}\n{exc.Message}\n{exc.StackTrace}");
			}
		}
	}
}
