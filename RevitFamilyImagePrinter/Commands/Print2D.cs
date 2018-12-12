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
using Application = Autodesk.Revit.ApplicationServices.Application;
using View = Autodesk.Revit.DB.View;
using Ookii.Dialogs.Wpf;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print2D : IExternalCommand
	{
		#region User Values
		public UserImageValues UserValues { get; set; } = new UserImageValues();
		public string UserImagePath { get; set; }
		#endregion
		
		#region Variables
		private IList<ElementId> views = new List<ElementId>();
		private Document doc
		{
			get
			{
				if (UIDoc != null)
					return UIDoc.Document;
				return null;
			}
		}
		public UIDocument UIDoc;
		public bool IsAuto = false;
		#endregion

		#region Constants
		private const int windowHeightOffset = 40;
		private const int windowWidthOffset = 10;
		#endregion	

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			UIApplication uiapp = commandData.Application;
			UIDoc = uiapp.ActiveUIDocument;

			FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
			viewCollector.OfClass(typeof(View));

			foreach (Element viewElement in viewCollector)
			{
				View view = (View)viewElement;

				if (view.Name.Equals("Level 1") && view.ViewType == ViewType.EngineeringPlan)
				{
					views.Add(view.Id);
					UIDoc.ActiveView = view;
				}
			}

			if(!message.Equals("FamilyPrint"))
			{
				UserImageValues userInputValues = RevitPrintHelper.ShowOptionsDialog(UIDoc, 40, 10);
				if (userInputValues == null)
					return Result.Failed;
				this.UserValues = userInputValues;
			}

			ViewChangesCommit();
			PrintCommit();

			return Result.Succeeded;
		}

		public void ViewChangesCommit()
		{
			IList<UIView> uiviews = UIDoc.GetOpenUIViews();
			foreach (var item in uiviews)
			{
				item.ZoomToFit();
				item.Zoom(UserValues.UserZoomValue);
				UIDoc.RefreshActiveView();
			}
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("SetView");
				doc.ActiveView.DetailLevel = UserValues.UserDetailLevel;
				doc.ActiveView.Scale = UserValues.UserScale;
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
			if (!IsAuto)
				UserImagePath = RevitPrintHelper.SelectFileNameDialog(initialName);
			if (UserImagePath == initialName) return;

			IList<ElementId> views = new List<ElementId>();
			views.Add(doc.ActiveView.Id);

			var exportOptions = new ImageExportOptions
			{
				ViewName = "temp",
				FilePath = UserImagePath,
				FitDirection = FitDirectionType.Vertical,
				HLRandWFViewsFileType = RevitPrintHelper.GetImageFileType(UserValues.UserExtension),
				ImageResolution = UserValues.UserImageResolution,
				ShouldCreateWebSite = false,
				PixelSize = UserValues.UserImageSize
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

			if (ImageExportOptions.IsValidFileName(UserImagePath))
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
