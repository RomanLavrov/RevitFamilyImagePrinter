using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace RevitFamilyImagePrinter.Infrastructure
{
	public static class ProjectHelper
	{
		/// <summary>
		/// Creates .rvt project from the specified .rfa file. Can print each .rvt file as an image.
		/// </summary>
		/// <param name="uiDoc">Autodesk.Revit.UI.UIDocument to manipulate in nested functions.</param>
		/// <param name="pathData">Paths to .rfa, .rvt and image files directories.</param>
		/// <param name="userValues">User print values that define printed image parameters. If null, projects won't be printed.</param>
		/// <param name="is3D">Defines whether images will be printed in isometric view</param>
		/// <returns>Returns a collection of symbols created during family processing.</returns>
		public static IEnumerable<string> CreateProjectsFromFamily(UIDocument uiDoc, PathData pathData, UserImageValues userValues = null, bool is3D = false)
		{
			Document doc = uiDoc.Document;
			RemoveExcessFamilies(doc);
			List<string> allSymbols = new List<string>();
			FileInfo familyInfo = new FileInfo(pathData.FamilyPath);
			FamilyData data = new FamilyData()
			{
				FamilyName = familyInfo.Name.Remove(familyInfo.Name.LastIndexOf('.'), 4),
				FamilyPath = familyInfo.FullName
			};
			Family family = LoadFamily(uiDoc, familyInfo);
			if (family == null) return null;

			if (!Directory.Exists(pathData.ProjectsPath))
				Directory.CreateDirectory(pathData.ProjectsPath);

			ISet<ElementId> familySymbolId = family.GetFamilySymbolIds();
			foreach (ElementId id in familySymbolId)
			{
				Element element = family.Document.GetElement(id);
				data.FamilySymbols.Add(element as FamilySymbol);
			}

			foreach (var symbol in data.FamilySymbols)
			{
				string nameProject = $"{data.FamilyName}&{symbol.Name}";
				allSymbols.Add(nameProject);

				string pathProject = Path.Combine(pathData.ProjectsPath, $"{nameProject}.rvt");

				RemoveExistingInstances(doc, symbol.Id);
				try
				{
					InsertInstanceIntoProject(uiDoc, symbol);
					CreateProjectFromFamilySymbol(uiDoc, symbol, pathProject);
					if (userValues != null)
					{
						PrintProject(uiDoc, userValues, pathData.ImagesPath, is3D);
					}
				}
				catch (Autodesk.Revit.Exceptions.InvalidOperationException exc)
				{
					string errorMessage = $"File {pathProject} already exists!";
					PrintHelper.ProcessError(exc, errorMessage, App.Logger, false);
				}

				DeleteElementCommit(doc, symbol);
			}

			DeleteElementCommit(doc, family);
			return allSymbols;
		}

		private static void PrintProject(UIDocument uiDoc, UserImageValues userValues, string imagesFolder, bool is3D)
		{
			string initialName = PrintHelper.GetFileName(uiDoc.Document);
			string filePath = Path.Combine(imagesFolder, initialName);

			if (is3D)
			{
				PrintHelper.SetActive3DView(uiDoc);
				PrintHelper.View3DChangesCommit(uiDoc, userValues);
			}
			else
			{
				PrintHelper.SetActive2DView(uiDoc);
				PrintHelper.View2DChangesCommit(uiDoc, userValues);
			}
			PrintHelper.PrintImageTransaction(uiDoc, userValues, filePath, true);
		}

		public static void CreateProjectFromFamilySymbol(UIDocument uiDoc, FamilySymbol symbol, string pathProject)
		{
			if (File.Exists(pathProject) && PrintHelper.IsFileAccessible(pathProject))
				File.Delete(pathProject);
			uiDoc.Document.SaveAs(pathProject);
		}

		public static void RemoveExcessFamilies(Document doc)
		{
			FilteredElementCollector famCollector
				= new FilteredElementCollector(doc);
			famCollector.OfClass(typeof(Family));

			var familiesList = famCollector.ToElements();

			for (int i = 0; i < familiesList.Count; i++)
			{
				DeleteElementCommit(doc, familiesList[i]);
			}
		}

		public static Family LoadFamily(UIDocument uiDoc, FileInfo file)
		{
			Document doc = uiDoc.Document;
			Family family = null;
			bool success = false;
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Load Family");
				success = doc.LoadFamily(file.FullName, out family);
				transaction.Commit();
			}
			if (!success)
				IsFamilyExists(uiDoc, ref family, file);
			return family;
		}

		public static bool IsFamilyExists(UIDocument uiDoc, ref Family family, FileInfo file)
		{
			string familyName = file.Name.Replace(file.Extension, string.Empty);
			var families = new FilteredElementCollector(uiDoc.Document)
				.OfClass(typeof(Family))
				.ToElements();
			foreach (var i in families)
			{
				if (i.Name != familyName) continue;
				family = i as Family;
				return true;
			}
			return false;
		}

		public static void InsertInstanceIntoProject(UIDocument uiDoc, FamilySymbol symbol)
		{
			Document doc = uiDoc.Document;
			View view = uiDoc.ActiveView;
			using (var transaction = new Transaction(doc, "Insert Symbol"))
			{
				transaction.Start();
				symbol.Activate();
				XYZ point = new XYZ(0, 0, 0);
				Level level = view.GenLevel;
				Element host = level as Element;
				doc.Create.NewFamilyInstance(point, symbol, host, StructuralType.NonStructural);
				transaction.Commit();
			}
		}

		public static void RemoveExistingInstances(Document doc, ElementId id)
		{
			var instances = new FilteredElementCollector(doc)
				.OfClass(typeof(FamilyInstance))
				.ToElements();
			foreach (var i in instances)
			{
				if (i.Id != id)
				{
					DeleteElementCommit(doc, i);
				}
			}
		}

		public static void DeleteElementCommit(Document doc, Element element)
		{
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Delete");
				doc.Delete(element.Id);
				transaction.Commit();
			}
		}

		public static List<FileInfo> GetFamilyFilesFromFolder(DirectoryInfo familiesFolder)
		{
			if (familiesFolder == null) return null;

			var familyFilesList = familiesFolder.GetFiles().Where(x => x.Extension.Equals(".rfa")).ToList();
			if (!familyFilesList.Any())
			{
				new TaskDialog("Fail")
				{
					TitleAutoPrefix = false,
					MainContent = ".rfa files have not been found in specified folder."
				}.Show();
				return null;
			}
			return familyFilesList;
		}
	}
}
