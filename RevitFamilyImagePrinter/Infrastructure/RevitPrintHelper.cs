using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Ookii.Dialogs.Wpf;

namespace RevitFamilyImagePrinter.Infrastructure
{
	public static class RevitPrintHelper
	{
		public static void PrintImage(Document doc, UserImageValues userValues, string filePath, bool isAuto = false)
		{
			string initialName = GetFileName(doc);
			if (!isAuto)
				filePath = SelectFileNameDialog(initialName);
			if (filePath == initialName) return;

			IList<ElementId> views = new List<ElementId>();
			views.Add(doc.ActiveView.Id);

			var exportOptions = new ImageExportOptions
			{
				ViewName = "temporary",
				FilePath = filePath,
				FitDirection = FitDirectionType.Vertical,
				HLRandWFViewsFileType = GetImageFileType(userValues.UserExtension),
				ImageResolution = userValues.UserImageResolution,
				ShouldCreateWebSite = false,
				PixelSize = userValues.UserImageSize
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
		}

		public static void SetActive2DView(UIDocument uiDoc)
		{
			Document doc = uiDoc.Document;
			FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
			viewCollector.OfClass(typeof(View));

			foreach (Element viewElement in viewCollector)
			{
				View view = (View)viewElement;

				if (view.Name.Equals("Level 1") && view.ViewType == ViewType.EngineeringPlan)
				{
					//views.Add(view.Id);
					uiDoc.ActiveView = view;
				}
			}
		}

		public static UserImageValues ShowOptionsDialog(UIDocument uiDoc, int windowHeightOffset = 40, int windowWidthOffset = 20, bool isApplyButtonVisible = true)
		{
			Window window = null;
			SinglePrintOptions options = null;
			using (Document doc = uiDoc.Document)
			{
				options = new SinglePrintOptions()
				{
					Doc = doc,
					UIDoc = uiDoc,
					IsPreview = isApplyButtonVisible
				};
				window = new Window
				{
					Height = options.Height + windowHeightOffset,
					Width = options.Width + windowWidthOffset,
					Title = "Image Print Settings",
					Content = options,
					Background = System.Windows.Media.Brushes.WhiteSmoke,
					WindowStyle = WindowStyle.ToolWindow,
					Name = "Options",
					ResizeMode = ResizeMode.NoResize,
					WindowStartupLocation = WindowStartupLocation.CenterScreen
				};
				window.Closing += Window_Closing;

				window.ShowDialog();
			}


			if (window.DialogResult != true)
				return null;
			return InitializeVariables(options);
		}

		private static UserImageValues InitializeVariables(SinglePrintOptions options)
		{
			return new UserImageValues()
			{
				UserScale = options.UserScale,
				UserImageSize = options.UserImageSize,
				UserImageResolution = options.UserImageResolution,
				UserZoomValue = options.UserZoomValue,
				UserExtension = options.UserExtension,
				UserDetailLevel = options.UserDetailLevel
			};
		}

		public static string SelectFileNameDialog(string name)
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

		public static DirectoryInfo SelectFolderDialog(string description, DirectoryInfo directory = null)
		{
			VistaFolderBrowserDialog dialog = new VistaFolderBrowserDialog()
			{
				Description = description,
				UseDescriptionForTitle = true,
				RootFolder = Environment.SpecialFolder.MyDocuments,
				ShowNewFolderButton = true
			};

			if (dialog.ShowDialog() != true)
			{
				return null;
			}
			return new DirectoryInfo(dialog.SelectedPath);

		}

		public static string GetFileName(Document doc)
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

		public static ImageFileType GetImageFileType(string userImagePath)
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

		private static void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			Window wnd = sender as Window;
			SinglePrintOptions options = wnd.Content as SinglePrintOptions;
			if(!options.IsPreview && !options.IsCancelled)
				options.SaveConfig();
		}

		public static void RemoveEmptyFamilies(UIDocument uiDoc)
		{
			using (Document doc = uiDoc.Document)
			{
				FilteredElementCollector famCollector
				  = new FilteredElementCollector(doc);
				famCollector.OfClass(typeof(Family));

				FilteredElementCollector instCollector
					= new FilteredElementCollector(doc);
				instCollector.OfClass(typeof(FamilyInstance));


				List<ElementType> elementsType = new List<ElementType>();
				foreach (FamilyInstance fi in instCollector)
				{
					ElementId typeId = fi.GetTypeId();
					elementsType.Add(doc.GetElement(typeId) as ElementType);
				}

				List<Element> elems = famCollector
												.Where(x => !x.Name.Equals(elementsType.FirstOrDefault().FamilyName))
												.Select(x => x)
												.ToList();

				for (int i = 0; i < elems.Count(); i++)
				{
					DeleteCommit(doc, elems[i]);
				}
				doc.Save();
			}
		}

		private static void DeleteCommit(Document doc, Element element)
		{
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Delete");
				doc.Delete(element.Id);
				transaction.Commit();
			}
		}
	}
}
