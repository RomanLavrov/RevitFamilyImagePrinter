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
		#region Variables
		private IList<ElementId> views = new List<ElementId>();
		private Document doc;
		private UIDocument uidoc;
		#endregion

		#region Constants
		private const int WindowHeightOffset = 40;
		private const int WindowWidthOffset = 10;
		#endregion

		#region User Values
		public int UserScale { get; set; }
		public int UserImageSize { get; set; }
		public ImageResolution UserImageResolution { get; set; }
		public string UserExtension { get; set; }
		public double UserZoomValue { get; set; }
		public ViewDetailLevel UserDetailLevel { get; set; }
		#endregion



		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			UIApplication uiapp = commandData.Application;
			uidoc = uiapp.ActiveUIDocument;
			doc = uidoc.Document;

			FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
			viewCollector.OfClass(typeof(View));

			foreach (Element viewElement in viewCollector)
			{
				View view = (View)viewElement;

				if (view.Name.Equals("Level 1") && view.ViewType == ViewType.EngineeringPlan)
				{
					views.Add(view.Id);
					uidoc.ActiveView = view;
				}
			}

			if (!ShowOptionsDialog())
				return Result.Failed;
			ViewChangesCommit();
			PrintCommit();

			return Result.Succeeded;
		}

		public void ViewChangesCommit()
		{
			IList<UIView> uiviews = uidoc.GetOpenUIViews();
			foreach (var item in uiviews)
			{
				item.ZoomToFit();
				item.Zoom(UserZoomValue);
				uidoc.RefreshActiveView();
			}

			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("SetView");
				doc.ActiveView.DetailLevel = UserDetailLevel;
				doc.ActiveView.Scale = UserScale;
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
			string initialName = GetFileName();
			string userImagePath = SelectFileNameDialog(initialName);
			if (userImagePath == initialName) return;

			IList<ElementId> views = new List<ElementId>();
			views.Add(doc.ActiveView.Id);

			var exportOptions = new ImageExportOptions
			{
				ViewName = "temp",
				FilePath = userImagePath,
				FitDirection = FitDirectionType.Vertical,
				HLRandWFViewsFileType = GetImageFileType(UserExtension),
				ImageResolution = UserImageResolution,
				ShouldCreateWebSite = false,
				PixelSize = UserImageSize
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

			exportOptions.ViewName = "temp";

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

		private string GetFileName()
		{
			switch (App.Version)
			{
				case "2018":
					{
						int indexDot = doc.Title.IndexOf('.');
						var name = doc.Title.Substring(0, indexDot);
						return name;
					}
				case "2019":
					{
						return doc.Title;
					}
				default: throw new Exception("Unknown Revit Version");
			}
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

		private bool ShowOptionsDialog()
		{
			SinglePrintOptions options = new SinglePrintOptions()
			{
				Doc = this.doc,
				UIDoc = this.uidoc
			};
			Window window = new Window
			{
				Height = options.Height + WindowHeightOffset,
				Width = options.Width + WindowWidthOffset,
				Title = "Image Print Settings",
				Content = options,
				Background = System.Windows.Media.Brushes.WhiteSmoke,
				WindowStyle = WindowStyle.ToolWindow,
				Name = "Options",
				ResizeMode = ResizeMode.NoResize,
				WindowStartupLocation = WindowStartupLocation.CenterScreen
			};

			window.ShowDialog();

			if (window.DialogResult != true)
				return false;
			InitializeVariables(options);
			return true;
		}

		private void InitializeVariables(SinglePrintOptions options)
		{
			UserScale = options.UserScale;
			UserImageSize = options.UserImageSize;
			UserImageResolution = options.UserImageResolution;
			UserZoomValue = options.UserZoomValue;
			UserExtension = options.UserExtension;
			UserDetailLevel = options.UserDetailLevel;
		}
	}
}
