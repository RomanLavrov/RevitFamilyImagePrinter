using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using RevitFamilyImagePrinter.Infrastructure;
using Application = Autodesk.Revit.Creation.Application;

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

		#region TMP

		PrintProgress progressWindow = new PrintProgress();

		private int totalFiles;

		private bool isCancelled;

		private TextBlock ProcessTextBlock;

		private ProgressBar PrintProgressBar;

		private Button CancelButton;

		private Label Counter;

		private Window window;

		private void CancelButton_Click(object sender, RoutedEventArgs e)
		{
			isCancelled = false;
			CancelButton.IsEnabled = false;
		}

		private void ApplicationOnViewActivated(object sender, ViewActivatedEventArgs e)
		{
			UIApplication uiApp = sender as UIApplication;
			string viewName = uiApp.ActiveUIDocument.ActiveView.Document.Title;
			if (viewName.ToLower().Equals("empty") ||
			    ProcessTextBlock.Text.Contains(viewName)) return;
			ProcessTextBlock.Text = $"{viewName} view has been activated";
			PrintProgressBar.Value++;
		}

		private void ApplicationOnDocumentSavedAs(object sender, DocumentSavedAsEventArgs e)
		{
			Document docSender = sender as Document;
			PrintProgressBar.Value++;
			ProcessTextBlock.Text = $"{docSender.Title} has been created";
			Counter.Content = PrintProgressBar.Value;
		}

		#endregion

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			_uiApp = commandData.Application;
			_uiDoc = _uiApp.ActiveUIDocument;
			var initProjectPath = _uiDoc.Document.PathName;
			RevitPrintHelper.CreateEmptyProject(commandData.Application.Application);

			DirectoryInfo familiesFolder = RevitPrintHelper.SelectFolderDialog("Select folder with needed families to be printed");
			if (familiesFolder == null)
				return Result.Cancelled;
			UserFolderTo = RevitPrintHelper.SelectFolderDialog("Select folder where to save printed files");
			if (UserFolderTo == null)
				return Result.Cancelled;

			UserValues =
				RevitPrintHelper.ShowOptionsDialog(_uiDoc, windowHeightOffset, windowWidthOffset, false, false);
			if (UserValues == null)
				return Result.Failed;

			UserFolderFrom = new DirectoryInfo(Path.Combine(familiesFolder.FullName, "Projects"));

			progressWindow.CancelButton.Click += CancelButton_Click;
			ProcessTextBlock = progressWindow.ProcessTextBlock;
			PrintProgressBar = progressWindow.PrintProgressBar;
			CancelButton = progressWindow.CancelButton;
			Counter = progressWindow.LabelCounter;
			window = new Window
			{
				Height = 200,
				Width = 400,
				Title = "Printing",
				WindowStyle = WindowStyle.ToolWindow,
				Name = "Printing",
				ResizeMode = ResizeMode.NoResize,
				WindowStartupLocation = WindowStartupLocation.CenterScreen,
				Content = progressWindow,
				Topmost = true
			};

			//commandData.Application.Application.DocumentSavedAs += ApplicationOnDocumentSavedAs;
			commandData.Application.ActiveUIDocument.Document.DocumentSavedAs += ApplicationOnDocumentSavedAs;

			PrintProgressBar.Value = 0;
			PrintProgressBar.Maximum = familiesFolder.GetFiles().Length;
			//ProcessTextBlock.Text = "Projects are being created...";

			window.Show();

			if (!CreateProjects(commandData, elements, familiesFolder))
				return Result.Failed;

			var fileList = Directory.GetFiles(UserFolderFrom.FullName);

			PrintProgressBar.Value = 0;
			PrintProgressBar.Maximum = fileList.Length;

			commandData.Application.ViewActivated += ApplicationOnViewActivated;

			foreach (var item in fileList)
			{
				if (isCancelled) break;
				RevitPrintHelper.OpenDocument(_uiDoc, App.DefaultProject);
				FileInfo info = new FileInfo(item);
				if (!info.Extension.Equals(".rvt"))
					continue;
				_uiDoc = commandData.Application.OpenAndActivateDocument(item);
				if (info.Length > maxSizeLength)
					RevitPrintHelper.RemoveEmptyFamilies(_uiDoc);
				RevitPrintHelper.SetActive3DView(_uiDoc);
				ViewChangesCommit();
				PrintCommit(_uiDoc.Document);
			}

			if (!string.IsNullOrEmpty(initProjectPath) && File.Exists(initProjectPath))
				_uiDoc = RevitPrintHelper.OpenDocument(_uiDoc, initProjectPath);
			else
				_uiDoc = RevitPrintHelper.OpenDocument(_uiDoc, App.DefaultProject);

			ProcessTextBlock.Text = "Done!";
			window.Close();
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
				string errorMessage = "### ERROR ### - Error occured during current view correction";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"{errorMessage}{endl}{exc.Message}{endl}{exc.StackTrace}");
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
				string errorMessage = "### ERROR ### - Error occured during printing of current view";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"{errorMessage}{endl}{exc.Message}{endl}{exc.StackTrace}");
			}
		}
	}
}
