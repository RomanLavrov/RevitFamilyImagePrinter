using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using MessageBox = System.Windows.MessageBox;
using View = Autodesk.Revit.DB.View;
using Ookii.Dialogs.Wpf;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print2D : IExternalCommand
	{
		#region User Values
		public UserImageValues userValues { get; set; } = new UserImageValues();
		#endregion
		
		#region Variables
		private IList<ElementId> views = new List<ElementId>();
		private Document doc;
		private UIDocument uiDoc;
		#endregion

		#region Constants
		private const int windowHeightOffset = 40;
		private const int windowWidthOffset = 10;
		#endregion	

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			UIApplication uiapp = commandData.Application;
			uiDoc = uiapp.ActiveUIDocument;
			using (doc = uiDoc.Document)
			{
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

				UserImageValues userInputValues = RevitPrintHelper.ShowOptionsDialog(doc, uiDoc, 40, 10);
				if (userInputValues == null)
					return Result.Failed;
				this.userValues = userInputValues;

				ViewChangesCommit();
				PrintCommit();
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
				transaction.Commit();
			}
		}

		private void PrintCommit()
		{
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Print");
				PrintImage();
				transaction.Commit();
			}
		}

		private void PrintImage()
		{
			string initialName = RevitPrintHelper.GetFileName(doc);
			string userImagePath = RevitPrintHelper.SelectFileNameDialog(initialName);
			if (userImagePath == initialName) return;

			IList<ElementId> views = new List<ElementId>();
			views.Add(doc.ActiveView.Id);

			var exportOptions = new ImageExportOptions
			{
				ViewName = "temp",
				FilePath = userImagePath,
				FitDirection = FitDirectionType.Vertical,
				HLRandWFViewsFileType = RevitPrintHelper.GetImageFileType(userValues.UserExtension),
				ImageResolution = userValues.UserImageResolution,
				ShouldCreateWebSite = false,
				PixelSize = userValues.UserImageSize
			};

			if (views.Count > 0)
			{
				exportOptions.SetViewsAndSheets(views);
				exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
			}
			else
			{
				exportOptions.ExportRange = ExportRange.VisibleRegionOfCurrentView;
			}

			if (ImageExportOptions.IsValidFileName(userImagePath))
			{
				doc.ExportImage(exportOptions);
			}
		}

		private string SelectFileNameDialog(string name)
		{
			string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			VistaFileDialog dialog = new VistaSaveFileDialog()
			{
				InitialDirectory = path,
				RestoreDirectory = true,
				Title = "Choose a directory",
				FilterIndex = 0,
				Filter = "Image Files (*.PNG, *.JPG, *.BMP) | *.png;*.jpg;*.bmp",
				FileName = name
			};

			if ((bool)dialog.ShowDialog())
			{
				path = dialog.FileName;
				return path;
			}
			return name;
		}

		private ImageFileType GetImageFileType(string userImagePath)
		{
			switch (Path.GetExtension(userImagePath).ToLower())
			{
				case ".png": return ImageFileType.PNG;
				case ".jpeg": return ImageFileType.JPEGLossless;
				case ".bmp": return ImageFileType.BMP;
				case ".tiff": return ImageFileType.TIFF;
				case ".targa": return ImageFileType.TARGA;
				default: throw new Exception("Unknown Image Format");
			}
		}

		//private bool ShowOptionsDialog()
		//{
		//	SinglePrintOptions options = new SinglePrintOptions()
		//	{
		//		Doc = this.doc,
		//		UIDoc = this.uiDoc
		//	};
		//	Window window = new Window
		//	{
		//		Height = options.Height + windowHeightOffset,
		//		Width = options.Width + windowWidthOffset,
		//		Title = "Image Print Settings",
		//		Content = options,
		//		Background = System.Windows.Media.Brushes.WhiteSmoke,
		//		WindowStyle = WindowStyle.ToolWindow,
		//		Name = "Options",
		//		ResizeMode = ResizeMode.NoResize,
		//		WindowStartupLocation = WindowStartupLocation.CenterScreen
		//	};

		//	window.ShowDialog();

		//	if (window.DialogResult != true)
		//		return false;
		//	/(options);
		//	return true;
		//}

		//private void InitializeVariables(SinglePrintOptions options)
		//{
		//	userValues.UserScale = options.UserScale;
		//	userValues.UserImageSize = options.UserImageSize;
		//	userValues.UserImageResolution = options.UserImageResolution;
		//	userValues.UserZoomValue = options.UserZoomValue;
		//	userValues.UserExtension = options.UserExtension;
		//	userValues.UserDetailLevel = options.UserDetailLevel;
		//}
	}
}
