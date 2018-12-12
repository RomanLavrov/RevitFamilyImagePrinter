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
		public DirectoryInfo FamiliesFolder { get; private set; }
		public DirectoryInfo ImagesFolder
		{
			get
			{
				return new DirectoryInfo(Path.Combine(FamiliesFolder.FullName, "Images"));
			}
		}
		public DirectoryInfo ProjectsFolder
		{
			get
			{
				return new DirectoryInfo(Path.Combine(FamiliesFolder.FullName, "Projects"));
			}
		}
		#endregion

		#region Variables
		private ExternalCommandData commandData;
		private ElementSet elements;
		private UIDocument uiDoc;
		private Document doc;
		private UserImageValues userValues;
		private List<string> allSymbols = new List<string>();
		private readonly Logger logger;
		#endregion

		public ProjectCreator()
		{
			logger = Logger.GetLogger();
		}

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			this.commandData = commandData;
			this.elements = elements;
			UIApplication uiapp = commandData.Application;
			Application app = uiapp.Application;
			uiDoc = uiapp.ActiveUIDocument;
			doc = uiDoc.Document;

			//FamiliesFolder = RevitPrintHelper.SelectFolderDialog("Select folder with families to be printed");
			//var familyFiles = Directory.GetFiles(FamiliesFolder.FullName).ToList();
			//var amount = TMPGetAmountOfFiles(familyFiles);
			//TaskDialog.Show("AMOUNT OF SYMBOLS", $"{amount} symbols in all families");
			var familyFiles = GetFamilyFilesFromFolder();

			//TaskDialog.Show("RER", $"{Environment.CurrentDirectory}");
			if (familyFiles == null || familyFiles.Count == 0)
				return Result.Failed;

			if (!ProcessProjects(familyFiles))
				return Result.Failed;

			if (!CheckProjectsAmount())
				return Result.Failed;

			return Result.Succeeded;
		}

		private bool CheckProjectsAmount()
		{
			if (!ProjectsFolder.Exists)
				return false;
			var projectsCreated = Directory.GetFiles(ProjectsFolder.FullName).ToList();
			logger.NewLine();
			try
			{
				Assert.AreEqual(allSymbols.Count, projectsCreated.Count);
			}
			catch(Exception exc)
			{
				TaskDialog.Show("Warning!", "Attention! The amount of projects created is not equal to amount of types in families. Check log file.");
				logger.WriteLine($"### ERROR ###\tThe amount of projects created is not equal to amount of types in families.\n{exc.Message}\nMissed projects:");
				var differences = allSymbols.Except(projectsCreated);
				foreach(var i in differences)
				{
					logger.WriteLine(i);
				}
				return false;
			}
			logger.EndLogSession();
			return true;
		}

		#region GetFamilyFilesFromFolder
		private List<string> GetFamilyFilesFromFolder()
		{
			try
			{
				FamiliesFolder = RevitPrintHelper.SelectFolderDialog("Select folder with families to be printed");
				userValues = RevitPrintHelper.ShowOptionsDialog(uiDoc, 40, 10, false);
				if (FamiliesFolder == null || userValues == null) return null;

				CreateFolders();

				var fileList = Directory.GetFiles(FamiliesFolder.FullName).ToList();
				return fileList;
			}
			catch(Exception exc)
			{
				TaskDialog.Show("Error", "Error during getting families from folder.");
				logger.WriteLine($"### ERROR ###\tError during getting families from folder.\n{exc.Message}");
				return null;
			}
		}

		private void CreateFolders()
		{
			if (!Directory.Exists(ProjectsFolder.FullName))
				Directory.CreateDirectory(ProjectsFolder.FullName);
			if (!Directory.Exists(ImagesFolder.FullName))
				Directory.CreateDirectory(ImagesFolder.FullName);
		}
		#endregion

		#region ProcessProjects
		private bool ProcessProjects(IEnumerable<string> fileList)
		{
			try
			{
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
						//string path = $@"D:\TypeImages\Families\{data.FamilyName}&{symbol.Name}.rvt";
						string nameProject = $"{data.FamilyName}&{symbol.Name}";
						allSymbols.Add(nameProject);
						string pathProject = Path.Combine(ProjectsFolder.FullName, $"{nameProject}.rvt");
						string pathImage = Path.Combine(ImagesFolder.FullName, $"{nameProject}{userValues.UserExtension}");
						if (File.Exists(pathProject)) continue;

						SaveFamilySymbol(symbol, pathProject);
						PrintImage(pathImage);
						RemoveTypeFromProject(symbol);
					}
					RemoveFamily(family);
				}
			}
			catch (Exception exc)
			{
				TaskDialog.Show("Error", "Error during processing of project");
				logger.WriteLine($"### ERROR ###\tError during processing of projects.\n{exc.Message}");
				return false;
			}
			return true;
		}

		private Family LoadFamily(FileInfo file)
		{
			Family family = null;
			try
			{
				bool success = false;
				using (Transaction transaction = new Transaction(doc))
				{
					transaction.Start("Load Family");
					success = doc.LoadFamily(file.FullName, out family);
					transaction.Commit();
				}
				if (!success)
					FindExistingFamily(ref family, file);
			}
			catch (Exception exc)
			{
				TaskDialog.Show("Error", "Error during loading of family from file.");
				logger.WriteLine($"### ERROR ###\tError during loading of family from file.\n{exc.Message}");
			}
			return family;
		}

		private bool FindExistingFamily(ref Family family, FileInfo file)
		{
			string familyName = file.Name.Replace(file.Extension, string.Empty);
			var families = new FilteredElementCollector(uiDoc.Document)
				  .OfClass(typeof(Family))
				  .ToElements();
			foreach(var i in families)
			{
				if (i.Name == familyName)
				{
					family = i as Family;
					return true;
				}

			}
			return false;
		}

		#region SaveFamilySymbol
		private void SaveFamilySymbol(FamilySymbol symbol, string pathProject)
		{
			RemoveExistingInstances(symbol.Id);
			InsertInstanceIntoProject(symbol);
			SaveProjectAs(symbol, pathProject);
		}

		private void RemoveExistingInstances(ElementId id)
		{
			try
			{
				var instances = new FilteredElementCollector(doc)
				  .OfClass(typeof(FamilyInstance))
				  .ToElementIds();
				foreach (var i in instances)
				{
					if (i != id)
					{
						using (Transaction transaction = new Transaction(doc, "Remove Unnecessary Instance"))
						{
							transaction.Start();
							doc.Delete(i);
							transaction.Commit();
						}
					}
				}
			}
			catch (Exception exc)
			{
				TaskDialog.Show("Error", "Error during removing of instances");
				logger.WriteLine($"### ERROR ###\tError during removing of instances.\n{exc.Message}");
			}
		}

		private void InsertInstanceIntoProject(FamilySymbol symbol)
		{
			View view = uiDoc.ActiveView;
			FamilyInstance instance = null;
			if (symbol == null) return;

			try
			{
				using (var transact = new Transaction(doc, "Insert Symbol"))
				{
					transact.Start();
					symbol.Activate();
					XYZ point = new XYZ(0, 0, 0);
					Level level = view.GenLevel;
					Element host = level as Element;
					instance = doc.Create.NewFamilyInstance(point, symbol, host, StructuralType.NonStructural);
					transact.Commit();
				}
			}
			catch (Exception exc)
			{
				TaskDialog.Show("Error", "Error during inserting of instances");
				logger.WriteLine($"### ERROR ###\tError during inserting of instances.\n{exc.Message}");
			}
		}

		private void SaveProjectAs(FamilySymbol symbol, string path)
		{
			try
			{
				//uidoc.GetOpenUIViews().FirstOrDefault().ZoomToFit(); // Remove in production
				/*var syms = new FilteredElementCollector(uidoc.Document)
				  .OfClass(typeof(FamilySymbol))
				  .ToElementIds();
				Debug.WriteLine($"SymbolName = {symbol.Name}, symbolId = {symbol.Id} / Amount = {syms.Count}");*/
				doc.SaveAs(path);
				// закрыть текущий, открыть пустой и поудалять временные
				// ОБЯЗАТЕЛЬНО ФАЙЛ-ЗАКЛАДКА
			}
			catch (Exception exc)
			{
				TaskDialog.Show("Error", "Error during saving of project");
				logger.WriteLine($"### ERROR ###\tError during saving of project.\n{exc.Message}");
			}
		}
		#endregion

		private void PrintImage(string path)
		{
			try
			{
				string msg = "FamilyPrint";
				Print2D print2d = new Print2D()
				{
					UserValues = userValues,
					UserImagePath = path,
					IsAuto = true,
					UIDoc = uiDoc,
				};
				print2d.Execute(commandData, ref msg, elements);
			}
			catch (Exception exc)
			{
				TaskDialog.Show("Error", "Error during printing of type");
				logger.WriteLine($"### ERROR ###\tError during printing of type.\n{exc.Message}");
			}
		}

		private void RemoveTypeFromProject(FamilySymbol symbol)
		{
			try
			{
				View view = uiDoc.ActiveView;
				if (symbol == null) return;

				int deletedSymbol = symbol.Id.IntegerValue;
				using (var transact = new Transaction(doc, "Delete Symbol"))
				{
					transact.Start();
					doc.Delete(symbol.Id);
					transact.Commit();
				}
			}
			catch (Exception exc)
			{
				TaskDialog.Show("Error", "Error during removing of type from project");
				logger.WriteLine($"### ERROR ###\tError during removing of symbol from project.\n{exc.Message}");
			}
		}

		private void RemoveFamily(Family family)
		{
			try
			{
				string deletedFamily = family.Name;
				using (Transaction transaction = new Transaction(doc))
				{
					transaction.Start("Delete Family");
					doc.Delete(family.Id);
					transaction.Commit();
				}
			}
			catch (Exception exc)
			{
				TaskDialog.Show("Error", "Error during removing of family from project");
				logger.WriteLine($"### ERROR ###\tError during removing of family from project.\n{exc.Message}");
			}
		}
		#endregion
	}
}
