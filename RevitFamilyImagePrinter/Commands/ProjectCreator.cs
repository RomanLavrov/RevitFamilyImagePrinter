using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RevitFamilyImagePrinter;
using RevitFamilyImagePrinter.Infrastructure;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Xml.Serialization;
using Application = Autodesk.Revit.ApplicationServices.Application;
using InvalidOperationException = Autodesk.Revit.Exceptions.InvalidOperationException;
using IOException = Autodesk.Revit.Exceptions.IOException;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class ProjectCreator : IExternalCommand
	{
		#region Properties

		public DirectoryInfo FamiliesFolder { get; set; }
		public DirectoryInfo ProjectsFolder => new DirectoryInfo(Path.Combine(FamiliesFolder.FullName, "Projects"));
		public UserImageValues UserValues { get; set; }

		#endregion

		#region Variables
		private ExternalCommandData _commandData;
		private ElementSet _elements;
		private UIDocument _uiDoc;
		private Document _doc;
		private readonly List<string> _allSymbols = new List<string>();
		private readonly Logger _logger = App.Logger;
		#endregion

		#region Constants 

		private const int windowWidthOffset = 20;
		private const int windowHeightOffset = 40;

		#endregion

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			this._commandData = commandData;
			this._elements = elements;
			UIApplication uiApplication = commandData.Application;
			_uiDoc = uiApplication.ActiveUIDocument;
			_doc = _uiDoc.Document;

			var familyFiles = GetFamilyFilesFromFolder();

			if (familyFiles == null || familyFiles.Count == 0)
				return Result.Failed;

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

		#region GetFamilyFilesFromFolder
		private List<FileInfo> GetFamilyFilesFromFolder()
		{
			try
			{
				if (FamiliesFolder == null) return null;

				
				var familyFilesList = FamiliesFolder.GetFiles().Where(x => x.Extension.Equals(".rfa")).ToList();
				if (!familyFilesList.Any())
				{
					new TaskDialog("Fail")
					{
						TitleAutoPrefix = false,
						MainContent = ".rfa files have not been found in specified folder."
					}.Show();
					//.Show("Fail", ".rfa files have not been found in specified folder.");
					return null;
				}

				if (UserValues == null)
					return null;

				if (!Directory.Exists(ProjectsFolder.FullName))
					Directory.CreateDirectory(ProjectsFolder.FullName);
				return familyFilesList;
			}
			catch(Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageFamiliesRetrieving)}", _logger);
				return null;
			}
		}

		#endregion

		#region ProcessProjects
		private bool ProcessProjects(IEnumerable<FileInfo> fileList)
		{
			try
			{
				RemoveExcessFamilies();
				foreach (var info in fileList)
				{
					FamilyData data = new FamilyData()
					{
						FamilyName = info.Name.Remove(info.Name.LastIndexOf('.'), 4), // removing extension
						FamilyPath = info.FullName
					};
					Family family = LoadFamily(info);
					if (family == null) continue;

					ISet<ElementId> familySymbolId = family.GetFamilySymbolIds();
					foreach (ElementId id in familySymbolId)
					{
						Element element = family.Document.GetElement(id);
						data.FamilySymbols.Add(element as FamilySymbol);
					}

					foreach (var symbol in data.FamilySymbols)
					{
						string nameProject = $"{data.FamilyName}&{symbol.Name}";
						_allSymbols.Add(nameProject);

						string pathProject = Path.Combine(ProjectsFolder.FullName, $"{nameProject}.rvt");
						//string pathImage = Path.Combine(ImagesFolder.FullName, $"{nameProject}{UserValues.UserExtension}");

						RemoveExistingInstances(symbol.Id);
						InsertInstanceIntoProject(symbol);
						if (File.Exists(pathProject) && RevitPrintHelper.IsFileAccessible(pathProject))
							File.Delete(pathProject);
						try
						{
							_doc.SaveAs(pathProject);
						}
						catch (InvalidOperationException exc)
						{
							string errorMessage = $"File {pathProject} already exists!";
							_logger.WriteLine($"### ERROR ### - {errorMessage}\n{exc.Message}\n{exc.StackTrace}");
						}
						RemoveTypeFromProject(symbol);
					}
					RemoveFamily(family);
				}
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageProjectProcessing)}", _logger);
				return false;
			}
			return true;
		}

		private void RemoveExcessFamilies()
		{
			try
			{
				FilteredElementCollector famCollector
					= new FilteredElementCollector(_doc);
				famCollector.OfClass(typeof(Family));

				var familiesList = famCollector.ToElements();

				for (int i = 0; i < familiesList.Count; i++)
				{
					DeleteCommit(familiesList[i]);
				}
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageFamilyRemoving)}", _logger);
			}
		}

		private void DeleteCommit(Element element)
		{
			using (Transaction transaction = new Transaction(_doc))
			{
				transaction.Start("Delete");
				_doc.Delete(element.Id);
				transaction.Commit();
			}
		}

		private Family LoadFamily(FileInfo file)
		{
			Family family = null;
			try
			{
				bool success = false;
				using (Transaction transaction = new Transaction(_doc))
				{
					transaction.Start("Load Family");
					success = _doc.LoadFamily(file.FullName, out family);
					transaction.Commit();
				}
				if (!success)
					FindExistingFamily(ref family, file);
			}
			catch (Exception exc)
			{	
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageFamilyLoading)}", _logger);
			}
			return family;
		}

		private bool FindExistingFamily(ref Family family, FileInfo file)
		{
			string familyName = file.Name.Replace(file.Extension, string.Empty);
			var families = new FilteredElementCollector(_uiDoc.Document)
				  .OfClass(typeof(Family))
				  .ToElements();
			foreach(var i in families)
			{
				if (i.Name != familyName) continue;
				family = i as Family;
				return true;

			}
			return false;
		}

		#region SaveFamilySymbol

		private void RemoveExistingInstances(ElementId id)
		{
			try
			{
				var instances = new FilteredElementCollector(_doc)
				  .OfClass(typeof(FamilyInstance))
				  .ToElements();
				foreach (var i in instances)
				{
					if (i.Id != id)
					{
						DeleteCommit(i);
					}
				}
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageInstanceRemoving)}", _logger);
			}
		}

		private void InsertInstanceIntoProject(FamilySymbol symbol)
		{
			View view = _uiDoc.ActiveView;
			if (symbol == null) return;

			try
			{
				using (var transact = new Transaction(_doc, "Insert Symbol"))
				{
					transact.Start();
					symbol.Activate();
					XYZ point = new XYZ(0, 0, 0);
					Level level = view.GenLevel;
					Element host = level as Element;
					_doc.Create.NewFamilyInstance(point, symbol, host, StructuralType.NonStructural);
					transact.Commit();
				}
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageInstanceInserting)}", _logger);
			}
		}

		#endregion

		private void RemoveTypeFromProject(FamilySymbol symbol)
		{
			try
			{
				if (symbol == null) return;

				DeleteCommit(symbol);
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageTypeRemoving)}", _logger);
			}
		}

		private void RemoveFamily(Family family)
		{
			try
			{
				DeleteCommit(family);
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageFamilyRemoving)}", _logger);
			}
		}
		#endregion
	}
}
