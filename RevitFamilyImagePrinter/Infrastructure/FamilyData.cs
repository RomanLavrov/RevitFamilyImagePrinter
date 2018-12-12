using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
