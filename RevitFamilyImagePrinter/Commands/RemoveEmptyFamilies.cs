using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class RemoveEmptyFamilies : IExternalCommand
	{
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			TaskDialogResult res = TaskDialog.Show("Confirm families removal", "This action will delete all families without instances. Are you sure?",
				TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
			if (res == TaskDialogResult.No)
				return Result.Cancelled;

			UIApplication uiapp = commandData.Application;
			UIDocument uidoc = uiapp.ActiveUIDocument;

			DeleteOperation(uidoc);

			return Result.Succeeded;
		}

		public void DeleteOperation(UIDocument uidoc)
		{
			using (Document doc = uidoc.Document)
			{
				FilteredElementCollector famCollector
				  = new FilteredElementCollector(doc);
				famCollector.OfClass(typeof(Family));

				FilteredElementCollector instCollector
					= new FilteredElementCollector(doc);
				instCollector.OfClass(typeof(FamilyInstance));


				List<ElementType> elementsType = new List<ElementType>();
				foreach (FamilyInstance fi in instCollector)
				{
					ElementId typeId = fi.GetTypeId();
					elementsType.Add(doc.GetElement(typeId) as ElementType);
				}

				List<Element> elems = famCollector
												.Where(x => !x.Name.Equals(elementsType.FirstOrDefault().FamilyName))
												.Select(x => x)
												.ToList();

				for (int i = 0; i < elems.Count(); i++)
				{
					DeleteCommit(doc, elems[i]);
				}
				doc.Save();
			}
		}

		private void DeleteCommit(Document doc, Element element)
		{
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Delete");
				doc.Delete(element.Id);
				transaction.Commit();
			}
		}
	}
}
