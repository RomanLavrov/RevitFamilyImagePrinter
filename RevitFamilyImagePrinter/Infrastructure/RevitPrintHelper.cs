using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Ookii.Dialogs.Wpf;
using RevitFamilyImagePrinter.Windows;
using Image = System.Drawing.Image;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace RevitFamilyImagePrinter.Infrastructure
{
	public static class RevitPrintHelper
	{
		#region Private

		private static void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			Window wnd = sender as Window;
			PrintOptions options = wnd.Content as PrintOptions;
			if (!options.IsPreview && !options.IsCancelled)
				options.SaveConfig();
		}

		private static UserImageValues InitializeVariables(PrintOptions options)
		{
			return new UserImageValues()
			{
				UserScale = options.UserScale,
				UserImageHeight = options.UserImageHeight,
				UserImageResolution = options.UserImageResolution,
				UserZoomValue = options.UserZoomValue,
				UserExtension = options.UserExtension,
				UserDetailLevel = options.UserDetailLevel,
				UserAspectRatio = options.UserAspectRatio
			};
		}

		private static void DeleteTransaction(Document doc, Element element)
		{
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Delete");
				doc.Delete(element.Id);
				transaction.Commit();
			}
		}

		private static void CropImageRectangle(UserImageValues userValues, FileInfo imageFile, FileInfo tmpFile)
		{

			int imgHeight = userValues.UserImageHeight;
			int imgWidth = 0;
			switch (userValues.UserAspectRatio)
			{
				case ImageAspectRatio.Ratio_16to9:
					imgWidth = imgHeight * 16 / 9;
					break;
				case ImageAspectRatio.Ratio_1to1: imgWidth = imgHeight;
					break;
				case ImageAspectRatio.Ratio_4to3:
					imgWidth = imgHeight * 4 / 3;
					break;
			}

			using (Bitmap image = Image.FromFile(tmpFile.FullName) as Bitmap)
			{

				if (image == null) return;

				//System.Drawing.Rectangle cropRectangle = new System.Drawing.Rectangle
				//{
				//	X = (int)(Math.Floor((image.Width - imgSize) / 2d)),
				//	Y = (image.Height - imgSize) / 2,
				//	Height = imgSize,
				//	Width = imgSize
				//};

				System.Drawing.Rectangle cropRectangle = new System.Drawing.Rectangle
				{
					Height = imgHeight,
					Width = imgWidth,
					X = (image.Width - imgWidth) / 2,
					Y = (image.Height - imgHeight) / 2
				};

				var result = image.Clone(cropRectangle, image.PixelFormat);
				result.Save($"{imageFile.FullName}");
				result.Dispose();
			}
			File.Delete(tmpFile.FullName);
		}

		private static void ZoomOpenUIViews(UIDocument uiDoc, double zoomValue, bool isToFit = true)
		{
			IList<UIView> uiViews = uiDoc.GetOpenUIViews();
			foreach (var item in uiViews)
			{
				if (isToFit)
				{
					item.ZoomToFit();
					uiDoc.RefreshActiveView();
				}
				item.Zoom(zoomValue);
				uiDoc.RefreshActiveView();
			}
		}

		private static void ActiveViewChangeTransaction(Document doc, UserImageValues userValues, bool is3D = false)
		{
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("SetView");
				doc.ActiveView.DetailLevel = is3D ? ViewDetailLevel.Fine : userValues.UserDetailLevel;
				doc.ActiveView.Scale = userValues.UserScale;
				transaction.Commit();
			}
		}

		private static double? GetScaleFromElement(UIDocument uiDoc)
		{
			Document doc = uiDoc.Document;
			var viewType = uiDoc.ActiveView.ViewType;
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(FamilyInstance));
			double scaleFactor = 1;
			foreach (var item in collector)
			{
				var box = item.get_BoundingBox(doc.ActiveView);
				if (box == null)
					return null;
				if (viewType == ViewType.ThreeD)
				{
					Scale3DCalculation(box, ref scaleFactor);
				}
				else
				{
					Scale2DCalculation(box, ref scaleFactor);
				}
			}

			if (viewType == ViewType.ThreeD)
			{
				var coefficient = 1.4;
				var finalScaleFactor = coefficient * scaleFactor;
				if (finalScaleFactor < 1)
					return finalScaleFactor;
			}
			scaleFactor *= 0.95;
			return scaleFactor;
		}

		private static void Scale2DCalculation(BoundingBoxXYZ box, ref double scaleFactor)
		{
			var width = box.Max.Y - box.Min.Y;
			var height = box.Max.X - box.Min.X;
			if (width / height < 1)
			{
				scaleFactor = (width / height);
			}
		}

		private static void Scale3DCalculation(BoundingBoxXYZ box, ref double scaleFactor)
		{
			var width = box.Max.Y - box.Min.Y;
			var height = box.Max.X - box.Min.X;
			var depth = box.Max.Z - box.Min.Z;

			var widthTotal = Math.Abs(Math.Cos(Math.PI / 6) * width) + Math.Abs(Math.Cos(Math.PI / 6) * depth);
			var heightTotal = Math.Abs(Math.Sin(Math.PI / 6) * width) + Math.Abs(Math.Sin(Math.PI / 6) * depth) + Math.Abs(height);

			if (widthTotal / heightTotal < 1)
			{
				scaleFactor = (widthTotal / heightTotal);
			}
		}

		private static bool IsDocumentActive(UIDocument uiDoc)
		{
			if (uiDoc.ActiveView == null)
				return false;
			IList<UIView> uiViews = uiDoc.GetOpenUIViews();
			foreach (var uiView in uiViews)
			{
				if (uiView.ViewId == uiDoc.ActiveView.Id)
					return true;
			}
			return false;
		}

		private static void R2019_HotFix()
		{
			if (App.Version != "2019") return;
			//TODO - Rewrite with updated Revit 2019 Documentation!
			var window = new Window()
			{
				Width = 10,
				Height = 10,
				ShowInTaskbar = false,
				ShowActivated = false,
				WindowStyle = WindowStyle.None
			};
			window.Show();
			window.Close();
		}

		private static void R2019_3DViewFix(UIDocument uiDoc)
		{
			if (App.Version != "2019") return;
			FilteredElementCollector collector = new FilteredElementCollector(uiDoc.Document);
			collector.OfClass(typeof(Level));
			var levelsToHide = new List<ElementId>();
			foreach (Element level in collector)
			{
				if (level.CanBeHidden(uiDoc.ActiveView))
				{
					levelsToHide.Add(level.Id);
				}
			}
			if (levelsToHide.Count < 1) return;

			using (Transaction transaction = new Transaction(uiDoc.Document, "Level Isolating"))
			{
				transaction.Start();
				uiDoc.ActiveView.HideElements(levelsToHide);
				transaction.Commit();
			}
		}

		private static void CorrectFileName(ref string fileName)
		{
			string item = fileName;
			if (fileName.Contains("Ø"))
			{
				fileName = fileName.Replace("Ø", "D");
			}

			if (fileName.Contains("Ä"))
			{
				fileName = fileName.Replace("Ä", "AE");
			}

			if (fileName.Contains("Ö"))
			{
				fileName = fileName.Replace("Ö", "OE");
			}

			if (fileName.Contains("ö"))
			{
				fileName = fileName.Replace("ö", "oe");
			}

			if (fileName.Contains("Ü"))
			{
				fileName = fileName.Replace("Ü", "UE");
			}

			if (fileName.Contains("ä"))
			{
				fileName = fileName.Replace("ä", "ae");
			}

			if (fileName.Contains("ß"))
			{
				fileName = fileName.Replace("ß", "ss");
			}

			if (fileName.Contains("ü"))
			{
				fileName = fileName.Replace("ü", "ue");
			}

			if (item.Contains("von oben nach unten") ||
				item.Contains("von unten nach oben") ||
				item.Contains("von oben") ||
				item.Contains("nach oben") ||
				item.Contains("von unten") ||
				item.Contains("nach unten") ||
				item.Contains("von ob nach un") ||
				item.Contains("von un nach ob") ||
				item.Contains("nach unten von oben") ||
				item.Contains("SCHWENKBAR MIT MOTORZOOM"))
			{
				fileName = fileName.Replace(' ', '_');
			}
		}

		#endregion

		#region Public

		#region Dialogs

		public static UserImageValues ShowOptionsDialog(UIDocument uiDoc, int windowHeightOffset = 40,
			int windowWidthOffset = 20, bool is3D = false, bool isApplyButtonVisible = true, bool isUpdateView = true)
		{
			Window window = null;
			PrintOptions options = null;
			using (Document doc = uiDoc.Document)
			{
				options = new PrintOptions()
				{
					Doc = doc,
					UIDoc = uiDoc,
					IsPreview = isApplyButtonVisible,
					Is3D = is3D,
					IsUpdateView = isUpdateView
				};
				window = new Window
				{
					Height = options.Height + windowHeightOffset,
					Width = options.Width + windowWidthOffset,
					Title = $"{App.Translator.GetValue(Translator.Keys.windowPrintSettingsTitle)}",
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

		public static string SelectFileNameDialog(string name)
		{
			string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			VistaFileDialog dialog = new VistaSaveFileDialog()
			{
				InitialDirectory = path,
				RestoreDirectory = true,
				Title = "Choose a directory",
				FilterIndex = 0,
				Filter = "Image Files (*.PNG, *.jpg, *.BMP) | *.png;*.jpg;*.bmp",
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

		#endregion

		public static void PrintImageTransaction(UIDocument uiDoc, UserImageValues userValues, string filePath, bool isAuto = false)
		{
			Document doc = uiDoc.Document;
			using (Transaction transaction = new Transaction(doc, "Print"))
			{
				transaction.Start();

				string initialName = GetFileName(doc);
				if (!isAuto)
					filePath = SelectFileNameDialog(initialName);
				if (filePath == initialName) return;
				IList<ElementId> views = new List<ElementId>();
				views.Add(doc.ActiveView.Id);

				//CorrectFileName(ref filePath);
				FileInfo imageFile = new FileInfo($"{filePath}{userValues.UserExtension}");
				var tmpFilePath = Path.Combine(imageFile.DirectoryName,
					$"{Guid.NewGuid().ToString()}{imageFile.Extension}");
				FileInfo tmpFile = new FileInfo(tmpFilePath);

				var exportOptions = new ImageExportOptions
				{
					ExportRange = ExportRange.VisibleRegionOfCurrentView,
					FilePath = tmpFilePath,
					FitDirection = FitDirectionType.Vertical,
					HLRandWFViewsFileType = GetImageFileType(userValues.UserExtension),
					ImageResolution = userValues.UserImageResolution,
					PixelSize = userValues.UserImageHeight,
					ShouldCreateWebSite = false,
					ShadowViewsFileType = GetImageFileType(userValues.UserExtension),
					ViewName = "temporary",
					ZoomType = ZoomFitType.FitToPage
				};

				ZoomOpenUIViews(uiDoc, userValues.UserZoomValue);

				if (views.Count > 0)
				{
					exportOptions.SetViewsAndSheets(views);
				}

				var scale = GetScaleFromElement(uiDoc);
				if (scale == null)
					return;

				ZoomOpenUIViews(uiDoc, (double)scale, false);

				if (ImageExportOptions.IsValidFileName(filePath))
				{
					doc.ExportImage(exportOptions);
				}
				transaction.Commit();

				CropImageRectangle(userValues, imageFile, tmpFile);
			}
			doc.Dispose();
		}

		public static void SetActive2DView(UIDocument uiDoc)
		{
			using (Document doc = uiDoc.Document)
			{
				FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
				viewCollector.OfClass(typeof(View));

				foreach (Element viewElement in viewCollector)
				{
					View view = (View)viewElement;

					if (view.Name.Equals("Level 1") && view.ViewType == ViewType.EngineeringPlan)
					{
						uiDoc.ActiveView = view;
					}
				}
			}
		}

		public static void SetActive3DView(UIDocument uiDoc)
		{
			using (Document doc = uiDoc.Document)
			{
				View3D view3D = null;
				var collector = new FilteredElementCollector(doc).OfClass(typeof(View3D));
				foreach (View3D view in collector)
				{
					if (view == null || view.IsTemplate) continue;
					view3D = view;
				}

				if (view3D == null)
				{
					var viewFamilyType = new FilteredElementCollector(doc)
						.OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
						.FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);
					using (Transaction trans = new Transaction(doc))
					{
						trans.Start("Add view");
						view3D = View3D.CreateIsometric(doc, viewFamilyType.Id);
						trans.Commit();
					}
				}
				uiDoc.ActiveView = view3D;
			}
		}

		public static void View2DChangesCommit(UIDocument uiDoc, UserImageValues userValues)
		{
			using (Document doc = uiDoc.Document)
			{
				ZoomOpenUIViews(uiDoc, userValues.UserZoomValue);
				ActiveViewChangeTransaction(doc, userValues);
				R2019_HotFix();
			}
		}

		public static void View3DChangesCommit(UIDocument uiDoc, UserImageValues userValues)
		{
			using (Document doc = uiDoc.Document)
			{
				ZoomOpenUIViews(uiDoc, userValues.UserZoomValue);

				FilteredElementCollector collector = new FilteredElementCollector(doc);
				collector.OfClass(typeof(View3D));
				foreach (View3D view3D in collector)
				{
					if (view3D == null) continue;
					ActiveViewChangeTransaction(doc, userValues, true);
				}
			}
			R2019_3DViewFix(uiDoc);
			R2019_HotFix();
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
			switch (Path.GetExtension(userImagePath)?.ToLower())
			{
				case ".png": return ImageFileType.PNG;
				case ".jpg": return ImageFileType.JPEGLossless;
				case ".bmp": return ImageFileType.BMP;
				case ".tiff": return ImageFileType.TIFF;
				case ".targa": return ImageFileType.TARGA;
				default: throw new Exception("Unknown Image Format");
			}
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
					.Where(x => !x.Name.Equals(elementsType.FirstOrDefault()?.FamilyName))
					.Select(x => x)
					.ToList();

				for (int i = 0; i < elems.Count(); i++)
				{
					DeleteTransaction(doc, elems[i]);
				}
			}
		}

		public static void CreateEmptyProject(Autodesk.Revit.ApplicationServices.Application application)
		{
			if (File.Exists(App.DefaultProject))
			{
				if (IsFileAccessible(App.DefaultProject))
					File.Delete(App.DefaultProject);
				else
					return;
			}
			Document emptyDoc = application.NewProjectDocument(UnitSystem.Metric);
			emptyDoc.SaveAs(App.DefaultProject);
		}

		public static UIDocument OpenDocument(UIDocument uiDoc, string newDocPath)
		{
			//FileStream stream = null;
			//try
			//{
			//	stream = File.Open(newDocPath, FileMode.Open);
			//}
			//catch (IOException)
			//{
			//	//the file is unavailable because it is:
			//	//still being written to
			//	//or being processed by another thread
			//	//or does not exist (has already been processed)
			//	return null;
			//}
			//finally
			//{
			//	stream?.Close();
			//}
			if (newDocPath.Equals(uiDoc.Application.ActiveUIDocument.Document?.PathName)) return uiDoc;
			UIDocument result = uiDoc.Application.OpenAndActivateDocument(newDocPath);
			if (!IsDocumentActive(uiDoc))
				uiDoc.Document.Close(false);
			return result;
		}

		public static bool IsFileAccessible(string path)
		{
			FileInfo file = new FileInfo(path);
			FileStream stream = null;
			try
			{
				stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
			}
			catch (IOException)
			{
				return false;
			}
			finally
			{
				stream?.Close();
			}
			return true;
		}

		public static void ProcessError(Exception exc, string errorMessage, Logger logger)
		{
			new TaskDialog($"{App.Translator.GetValue(Translator.Keys.errorMessageTitle)}")
			{
				TitleAutoPrefix = false,
				MainIcon = Autodesk.Revit.UI.TaskDialogIcon.TaskDialogIconError,
				MainContent = errorMessage
			}.Show();
			logger.WriteLine($"### ERROR ### - {errorMessage}\n{exc.Message}\n{exc.StackTrace}");
		}

		#endregion
	}
}
