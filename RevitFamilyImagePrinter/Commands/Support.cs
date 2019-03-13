using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitFamilyImagePrinter.Commands
{
    [Transaction(TransactionMode.Manual)]
    class Support : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Process.Start("mailto:support@building360.ch?subject=ImagePrinter");
            return Result.Succeeded;
        }
    }
}
