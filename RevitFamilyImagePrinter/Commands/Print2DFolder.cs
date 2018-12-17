using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using RevitFamilyImagePrinter.Infrastructure;
using View = Autodesk.Revit.DB.View;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print2DFolder : IExternalCommand
	{
		#region Properties
		public UserImageValues UserValues { get; set; } = new UserImageValues();
		public DirectoryInfo UserFolderFrom { get; set; } = new DirectoryInfo(@"D:\WebTypes\TestTypes");
		public DirectoryInfo UserFolderTo { get; set; } = new DirectoryInfo(@"D:\TypeImages");
		#endregion

		#region Variables
		//private IList<ElementId> views = new List<ElementId>();
		private Document _doc;
		private UIDocument _uiDoc;
		private readonly Logger _logger = Logger.GetLogger();
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
				_uiDoc = uiapp.ActiveUIDocument;
				_doc = _uiDoc.Document;
				
				UserFolderFrom = RevitPrintHelper.SelectFolderDialog("Select folder with needed files to be printed");
				if (UserFolderFrom == null)
					return Result.Cancelled;
				UserFolderTo = RevitPrintHelper.SelectFolderDialog("Select folder where to save printed files");
				if (UserFolderTo == null)
					return Result.Cancelled;

				UserValues = RevitPrintHelper.ShowOptionsDialog(_uiDoc, windowHeightOffset, windowWidthOffset, false);
				if (UserValues == null)
					return Result.Failed;

				var fileList = Directory.GetFiles(UserFolderFrom.FullName);
				foreach (var item in fileList)
				{
					FileInfo info = new FileInfo(item);
					if (!info.Extension.Equals(".rvt"))
						continue;
					using (_uiDoc = commandData.Application.OpenAndActivateDocument(item))
					{
						if (info.Length > maxSizeLength)
							RevitPrintHelper.RemoveEmptyFamilies(_uiDoc);
						using (_doc = _uiDoc.Document)
						{
							_doc = _uiDoc.Document;

							RevitPrintHelper.SetActive2DView(_uiDoc);

							ViewChangesCommit();
							PrintCommit();

							//uiDoc = commandData.Application.OpenAndActivateDocument("D:\\Empty.rvt");
						}
					}
				}

				//Rename2DImages();
			}
			catch (Exception exc)
			{
				string errorMessage = $"Error occured during command execution";
				_logger.WriteLine($"### EXCEPTION ### - {errorMessage}\n{exc.Message} ");
				return Result.Failed;
			}
			return Result.Succeeded;
		}

		public void ViewChangesCommit()
		{
			IList<UIView> uiviews = _uiDoc.GetOpenUIViews();
			foreach (var item in uiviews)
			{
				item.ZoomToFit();
				item.Zoom(UserValues.UserZoomValue);
				_uiDoc.RefreshActiveView();
			}

			using (Transaction transaction = new Transaction(_doc))
			{
				transaction.Start("SetView");
				_doc.ActiveView.DetailLevel = UserValues.UserDetailLevel;
				_doc.ActiveView.Scale = UserValues.UserScale;
				transaction.Commit();
			}
		}

		private void PrintCommit()
		{
			string initialName = RevitPrintHelper.GetFileName(_doc);
			string filePath = Path.Combine(UserFolderTo.FullName,
				$"{initialName}{UserValues.UserExtension}");

			using (Transaction transaction = new Transaction(_doc))
			{
				transaction.Start("Print");
				RevitPrintHelper.PrintImage(_doc, UserValues, filePath, true);
				transaction.Commit();
			}
		}

		//private void PrintImage()
		//{
		//	string initialName = RevitPrintHelper.GetFileName(doc);

		//	string filePath = Path.Combine(UserFolderTo.FullName,
		//		initialName + UserValues.UserExtension);

		//	IList<ElementId> views = new List<ElementId>();
		//	views.Add(doc.ActiveView.Id);

		//	var exportOptions = new ImageExportOptions
		//	{
		//		ViewName = "temp",
		//		FilePath = filePath,
		//		FitDirection = FitDirectionType.Vertical,
		//		HLRandWFViewsFileType = RevitPrintHelper.GetImageFileType(UserValues.UserExtension),
		//		ImageResolution = UserValues.UserImageResolution,
		//		ShouldCreateWebSite = false,
		//		PixelSize = UserValues.UserImageSize,
		//		ZoomType = ZoomFitType.FitToPage
		//	};

		//	if (views.Count > 0)
		//	{
		//		exportOptions.SetViewsAndSheets(views);
		//		//exportOptions.ExportRange = ExportRange.SetOfViews;
		//		exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
		//	}
		//	else
		//	{
		//		exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
		//	}

		//	if (ImageExportOptions.IsValidFileName(filePath))
		//	{
		//		doc.ExportImage(exportOptions);
		//	}
		//}

		//private void Rename2DImages()
		//{
		//	var picturesList = Directory.GetFiles(UserFolderTo.FullName);
		//	foreach (var item in picturesList)
		//	{
		//		int index = item.IndexOf("- Structural Plan - Level 1");
		//		if (index > 0)
		//		{
		//			//try
		//			//{
		//			File.Move(item, item.Substring(0, index) + ".png");
		//			//}
		//			//catch { };
		//		}
		//	}
		//}
	}
}
