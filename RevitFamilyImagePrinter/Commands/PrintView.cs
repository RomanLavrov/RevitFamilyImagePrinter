using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitFamilyImagePrinter.Infrastructure;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class PrintView : IExternalCommand
	{
		#region User Values
		public UserImageValues UserValues { get; set; } = new UserImageValues();
		public string UserImagePath { get; set; }
		#endregion

		#region Variables

		private Document doc => UIDoc?.Document;
		private readonly Logger _logger = Logger.GetLogger();
		public UIDocument UIDoc;

		#endregion

		#region Constants

		private const int windowHeightOffset = 40;
		private const int windowWidthOffset = 20;

		#endregion

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			UIApplication uiapp = commandData.Application;
			UIDoc = uiapp.ActiveUIDocument;

			UserImageValues userInputValues = RevitPrintHelper.ShowOptionsDialog(UIDoc, windowHeightOffset, windowWidthOffset, false, false, false);
			if (userInputValues == null)
				return Result.Cancelled;
			this.UserValues = userInputValues;

			string initialName = RevitPrintHelper.GetFileName(doc);
			string filePath = RevitPrintHelper.SelectFileNameDialog(initialName);
			if (filePath == initialName) return Result.Failed;

			IList<ElementId> views = new List<ElementId>();
			views.Add(doc.ActiveView.Id);

			//FileInfo imageFile = new FileInfo($"{filePath}{UserValues.UserExtension}");

			using (Transaction transaction = new Transaction(doc, "PrintView"))
			{
				transaction.Start();
				var exportOptions = new ImageExportOptions
				{
					ViewName = "temporary",
					FilePath = filePath,
					FitDirection = FitDirectionType.Vertical,
					HLRandWFViewsFileType = RevitPrintHelper.GetImageFileType(UserValues.UserExtension),
					ImageResolution = UserValues.UserImageResolution,
					ShouldCreateWebSite = false,
					PixelSize = UserValues.UserImageHeight
				};

				if (views.Count > 0)
				{
					exportOptions.SetViewsAndSheets(views);
				}

				exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;

				if (ImageExportOptions.IsValidFileName(filePath))
				{
					doc.ExportImage(exportOptions);
				}

				transaction.Commit();
			}
			return Result.Succeeded;
		}
	}
}
