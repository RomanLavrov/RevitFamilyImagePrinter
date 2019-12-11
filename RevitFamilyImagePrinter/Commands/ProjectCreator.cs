using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RevitFamilyImagePrinter.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class ProjectCreator : IExternalCommand
	{
		#region Properties

		public DirectoryInfo FamiliesFolder { get; set; }
		public DirectoryInfo ProjectsFolder => new DirectoryInfo(Path.Combine(FamiliesFolder.FullName, 
			App.Translator.GetValue(Translator.Keys.folderProjectsName)));

		#endregion

		#region Variables
		private UIDocument _uiDoc;
		private Document _doc;
		private readonly List<string> _allSymbols = new List<string>();
		private readonly Logger _logger = App.Logger;
		#endregion

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			UIApplication uiApplication = commandData.Application;
			_uiDoc = uiApplication.ActiveUIDocument;
			_doc = _uiDoc.Document;

			var familyFiles = ProjectHelper.GetFamilyFilesFromFolder(FamiliesFolder);

			if (familyFiles == null || familyFiles.Count == 0)
				return Result.Failed;

			if (!Directory.Exists(ProjectsFolder.FullName))
				Directory.CreateDirectory(ProjectsFolder.FullName);

			if (!ProcessProjects(familyFiles))
				return Result.Failed;

			CheckProjectsAmount();
			return Result.Succeeded;
		}

		private void CheckProjectsAmount()
		{
			if (!ProjectsFolder.Exists)
				throw new Exception($"{App.Translator.GetValue(Translator.Keys.errorMessageNoProjectsFolder)}");
			var projectsCreated = ProjectsFolder.GetFiles().Where(x => x.Extension.Equals(".rvt")).ToList();
			try
			{
				Assert.AreEqual(_allSymbols.Count, projectsCreated.Count);
			}
			catch (AssertFailedException exc)
			{
				new TaskDialog($"{App.Translator.GetValue(Translator.Keys.warningMessageTitle)}")
				{
					TitleAutoPrefix = false,
					MainIcon = TaskDialogIcon.TaskDialogIconWarning,
					MainContent =
						$"{App.Translator.GetValue(Translator.Keys.warningMessageProjectsAmount)}"
				}.Show();
				var differences = _allSymbols.Except(projectsCreated.Select(x => x.Name.ToString()));
				string output = string.Empty;
				foreach (var i in differences)
				{
					output += $"{i}\n";
				}
				_logger.WriteLine(
					$"### ERROR ### - The amount of projects created is not equal to the amount of types in families." +
					$"\n{exc.Message}\nMismatch projects:\n{output}");
			}
		}

		private bool ProcessProjects(IEnumerable<FileInfo> fileList)
		{
			try
			{
				foreach (var info in fileList)
				{
					PathData pathData = new PathData
					{
						FamilyPath = info.FullName,
						ProjectsPath = ProjectsFolder.FullName
					};
					_allSymbols.AddRange(ProjectHelper.CreateProjectsFromFamily(_uiDoc, pathData));
				}
				return true;
			}
			catch (Exception exc)
			{
				PrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageProjectProcessing)}", _logger);
				return false;
			}
		}

		private void RemoveExcessFamilies()
		{
			try
			{
				ProjectHelper.RemoveExcessFamilies(_doc);
			}
			catch (Exception exc)
			{
				PrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageFamilyRemoving)}", _logger);
			}
		}
	}
}
