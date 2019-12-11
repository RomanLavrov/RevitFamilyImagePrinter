using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace RevitFamilyImagePrinter.Infrastructure
{
	public class FamilyData
	{
		public FamilyData()
		{
			FamilySymbols = new List<FamilySymbol>();
		}
		public string FamilyName { get; set; }
		public string FamilyPath { get; set; }
		public List<FamilySymbol> FamilySymbols { get; set; }
	}
}
