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
    class UserHelp : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Process.Start("http://building360.ch/ImagePrinter");
            return Result.Succeeded;
        }
    }
}
