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
		private readonly Logger _logger;
		#endregion

		#region Constants 

		private const int windowWidthOffset = 20;
		private const int windowHeightOffset = 40;

		#endregion

		public ProjectCreator()
		{
			_logger = Logger.GetLogger();
		}

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

			return !CheckProjectsAmount() ? Result.Failed : Result.Succeeded;
		}

		private bool CheckProjectsAmount()
		{
			if (!ProjectsFolder.Exists)
				return false;
			var projectsCreated = Directory.GetFiles(ProjectsFolder.FullName).ToList();
			_logger.NewLine();
			try
			{
				Assert.AreEqual(_allSymbols.Count, projectsCreated.Count);
			}
			catch(Exception exc)
			{
				new TaskDialog("Warning!")
				{
					TitleAutoPrefix = false,
					MainContent = "Attention! The amount of projects created is not equal to amount of types in families. Check log file."
				}.Show();
				_logger.WriteLine($"### ERROR ###\tThe amount of projects created is not equal to amount of types in families.\n{exc.Message}\nMissed projects:");
				var differences = _allSymbols.Except(projectsCreated);
				foreach(var i in differences)
				{
					_logger.WriteLine(i);
				}
				return false;
			}
			return true;
		}

		#region GetFamilyFilesFromFolder
		private List<string> GetFamilyFilesFromFolder()
		{
			try
			{
				//FamiliesFolder = RevitPrintHelper.SelectFolderDialog("Select folder with families to be printed");
				if (FamiliesFolder == null) return null;

				var allFilesList = Directory.GetFiles(FamiliesFolder.FullName);
				var familyFilesList = allFilesList.Where(x => x.EndsWith(".rfa")).ToList();
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

				//UserValues = RevitPrintHelper.ShowOptionsDialog(_uiDoc, windowHeightOffset, windowWidthOffset, false, false);
				if (UserValues == null)
					return null;

				if (!Directory.Exists(ProjectsFolder.FullName))
					Directory.CreateDirectory(ProjectsFolder.FullName);
				return familyFilesList;
			}
			catch(Exception exc)
			{
				string errorMessage = "Error during getting families from folder.";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ###\t{errorMessage}\n{exc.Message}");
				return null;
			}
		}

		#endregion

		#region ProcessProjects
		private bool ProcessProjects(IEnumerable<string> fileList)
		{
			try
			{
				RemoveExcessFamilies();
				foreach (var i in fileList)
				{
					FileInfo info = new FileInfo(i);
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
						if (!File.Exists(pathProject))
							_doc.SaveAs(pathProject);

						RemoveTypeFromProject(symbol);
					}
					RemoveFamily(family);
				}
			}
			catch (Exception exc)
			{
				string errorMessage = "Error during processing of project";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ###\t{errorMessage}\n{exc.Message}");
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
				string errorMessage = "Error during deleting element from project.";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ###\t{errorMessage}\n{exc.Message}");
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
				string errorMessage = "Error during loading of family from file.";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ###\t{errorMessage}\n{exc.Message}");
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
				string errorMessage = "Error during removing of instances.";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ###\t{errorMessage}\n{exc.Message}");
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
				string errorMessage = "Error during inserting of instances";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ###\t{errorMessage}\n{exc.Message}");
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
				string errorMessage = "Error during removing of type from project";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ###\t{errorMessage}\n{exc.Message}");
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
				string errorMessage = "Error during removing of family from project";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"### ERROR ###\t{errorMessage}\n{exc.Message}");
			}
		}
		#endregion
	}
}
