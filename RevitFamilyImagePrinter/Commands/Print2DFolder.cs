using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using View = Autodesk.Revit.DB.View;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print2DFolder : IExternalCommand
	{
		#region Properties
		public UserImageValues userValues { get; set; } = new UserImageValues();
		public DirectoryInfo userFolderFrom { get; set; } = new DirectoryInfo(@"D:\WebTypes\TestTypes");
		public DirectoryInfo userFolderTo { get; set; } = new DirectoryInfo(@"D:\TypeImages");
		#endregion

		#region Variables
		//TODO Add FBD for images output
		private string imagePath = "D:\\TypeImages\\";
		private IList<ElementId> views = new List<ElementId>();
		private Document doc;
		private UIDocument uiDoc;
		#endregion

		#region Constants
		private const int windowHeightOffset = 40;
		private const int windowWidthOffset = 10;
		private const int maxSizeLength = 2097152;
		#endregion

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			try
			{
				UIApplication uiapp = commandData.Application;
				uiDoc = uiapp.ActiveUIDocument;
				using (doc = uiDoc.Document)
				{
					UserImageValues userInputValues = RevitPrintHelper.ShowOptionsDialog(doc, uiDoc, 40, 10, false);
					if (userInputValues == null)
						return Result.Failed;
					this.userValues = userInputValues;
				}

				userFolderFrom = RevitPrintHelper.SelectFolderDialog("Select folder with needed files to be printed");
				userFolderTo = RevitPrintHelper.SelectFolderDialog("Select folder where to save printed files");
				if (userFolderFrom == null || userFolderTo == null)
					return Result.Failed;
				var fileList = Directory.GetFiles(userFolderFrom.FullName);
				foreach (var item in fileList)
				{
					FileInfo info = new FileInfo(item);
					if (!info.Extension.Equals(".rvt")) // item.Contains("000") ||
						continue;
					using (uiDoc = commandData.Application.OpenAndActivateDocument(item))
					{
						if (info.Length > maxSizeLength)
							RevitPrintHelper.RemoveEmptyFamilies(uiDoc);
						using (doc = uiDoc.Document)
						{
							doc = uiDoc.Document;

							FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
							viewCollector.OfClass(typeof(View));

							foreach (Element viewElement in viewCollector)
							{
								View view = (View)viewElement;
								if (view.Name.Equals("Level 1") && view.ViewType == ViewType.EngineeringPlan)
								{
									views.Add(view.Id);
									uiDoc.ActiveView = view;
								}
							}

							ViewChangesCommit();
							PrintCommit();

							//uiDoc = commandData.Application.OpenAndActivateDocument("D:\\Empty.rvt");
						}
					}
				}

				Rename2DImages();
			}
			catch (Exception exc)
			{
				TaskDialog.Show($"ERROR", $"{exc.Message} // {exc.Source} // {exc.StackTrace}");
				Debug.Print($"### EXCEPTION ### - {exc.Message} // {exc.Source} // {exc.StackTrace}");
			}
			return Result.Succeeded;
		}

		public void ViewChangesCommit()
		{
			IList<UIView> uiviews = uiDoc.GetOpenUIViews();
			foreach (var item in uiviews)
			{
				item.ZoomToFit();
				item.Zoom(userValues.UserZoomValue);
				uiDoc.RefreshActiveView();
			}

			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("SetView");
				doc.ActiveView.DetailLevel = userValues.UserDetailLevel;
				doc.ActiveView.Scale = userValues.UserScale;
				transaction.RollBack();
			}
		}

		private void PrintCommit()
		{
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Print");
				PrintImage();
				transaction.RollBack();
			}
		}

		private void PrintImage()
		{
			string initialName = RevitPrintHelper.GetFileName(doc);

			string filePath = Path.Combine(userFolderTo.FullName,
				initialName + userValues.UserExtension);

			IList<ElementId> views = new List<ElementId>();
			views.Add(doc.ActiveView.Id);

			var exportOptions = new ImageExportOptions
			{
				ViewName = "temp",
				FilePath = filePath,
				FitDirection = FitDirectionType.Vertical,
				HLRandWFViewsFileType = RevitPrintHelper.GetImageFileType(userValues.UserExtension),
				ImageResolution = userValues.UserImageResolution,
				ShouldCreateWebSite = false,
				PixelSize = userValues.UserImageSize,
				ZoomType = ZoomFitType.FitToPage
			};

			if (views.Count > 0)
			{
				exportOptions.SetViewsAndSheets(views);
				//exportOptions.ExportRange = ExportRange.SetOfViews;
				exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
			}
			else
			{
				exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
			}

			if (ImageExportOptions.IsValidFileName(filePath))
			{
				doc.ExportImage(exportOptions);
			}
		}

		private void Rename2DImages()
		{
			var picturesList = Directory.GetFiles(imagePath);
			foreach (var item in picturesList)
			{
				int index = item.IndexOf("- Structural Plan - Level 1");
				if (index > 0)
				{
					//try
					//{
					File.Move(item, item.Substring(0, index) + ".png");
					//}
					//catch { };
				}
			}
		}
	}
}
