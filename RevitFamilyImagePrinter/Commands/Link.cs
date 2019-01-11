using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

/// <summary>
/// Command to open building360.ch webpage in default user browser.
/// </summary>

namespace RevitFamilyImagePrinter.Commands
{
    [Transaction(TransactionMode.Manual)]

    class Link : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            System.Diagnostics.Process.Start("http://building360.ch");
            return Result.Succeeded;
        }
    }
}
