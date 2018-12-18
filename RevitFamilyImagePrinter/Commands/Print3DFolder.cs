﻿using System;
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

		private Document _doc;
		private UIDocument _uiDoc;
		private readonly Logger _logger = Logger.GetLogger();

		#endregion

		#region Constants

		private const int windowHeightOffset = 40;
		private const int windowWidthOffset = 10;
		private const int maxSizeLength = 2097152;
		private readonly string endl = Environment.NewLine;

		#endregion

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			_uiDoc = commandData.Application.ActiveUIDocument;
			_doc = _uiDoc.Document;

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

			CreateProjects(commandData, elements, familiesFolder);

			var fileList = Directory.GetFiles(UserFolderFrom.FullName);
			foreach (var item in fileList)
			{
				FileInfo info = new FileInfo(item);
				if (!info.Extension.Equals(".rvt"))
					continue;
				_uiDoc = commandData.Application.OpenAndActivateDocument(item);
				using (_doc = _uiDoc.Document)
				{
					RevitPrintHelper.SetActive3DView(_uiDoc);
					ViewChangesCommit();
					PrintCommit();

					//commandData.Application.ViewActivated += SetViewParameters;
				}

			}
			return Result.Succeeded;
		}

		private void CreateProjects(ExternalCommandData commandData, ElementSet elements, DirectoryInfo familiesFolder)
		{
			ProjectCreator creator = new ProjectCreator()
			{
				FamiliesFolder = familiesFolder,
				UserValues = this.UserValues
			};
			string tmp = string.Empty;
			creator.Execute(commandData, ref tmp, elements);
			UserFolderFrom = creator.ProjectsFolder;
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
				_logger.WriteLine($"{errorMessage}{endl}{exc.Message}");
			}
		}

		private void PrintCommit()
		{
			try
			{
				string initialName = RevitPrintHelper.GetFileName(_doc);
				string filePath = Path.Combine(UserFolderTo.FullName,
					$"{initialName}{UserValues.UserExtension}");
				RevitPrintHelper.PrintImageTransaction(_doc, UserValues, filePath, true);
			}
			catch (Exception exc)
			{
				string errorMessage = "### ERROR ### - Error occured during printing of current view";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"{errorMessage}{endl}{exc.Message}");
			}
		}
	}
}
