using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Ookii.Dialogs.Wpf;
using RevitFamilyImagePrinter.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows;
using Image = System.Drawing.Image;

namespace RevitFamilyImagePrinter.Infrastructure
{
	internal partial class XYZPoint
	{
		public double X { get; set; }
		public double Y { get; set; }
		public double Z { get; set; }
	}
	public static class PrintHelper
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

			int resultHeight = userValues.UserImageHeight;
			int resultWidth = 0;
			switch (userValues.UserAspectRatio)
			{
				case ImageAspectRatio.Ratio_16to9:
					resultWidth = resultHeight * 16 / 9;
					break;
				case ImageAspectRatio.Ratio_1to1:
					resultWidth = resultHeight;
					break;
				case ImageAspectRatio.Ratio_4to3:
					resultWidth = resultHeight * 4 / 3;
					break;
			}

			Bitmap image = Image.FromFile(tmpFile.FullName) as Bitmap;
			if (image == null) return;

			System.Drawing.Rectangle cropRectangle = new System.Drawing.Rectangle
			{
				Width = resultWidth,
				Height = resultHeight,
				X = (image.Width - resultWidth) / 2,
				Y = (image.Height - resultHeight) / 2
			};

			if (image.Width - resultWidth < 0)
			{
				cropRectangle.Width = image.Width;
				cropRectangle.X = 0;
				using (var resultBitmap = image.Clone(cropRectangle, image.PixelFormat))
				{
					Image resultImage = ResizeImage(resultBitmap, resultWidth, resultHeight);
					resultImage.Save($"{imageFile.FullName}");
					resultImage.Dispose();
				}
			}
			else
			{
				using (var resultBitmap = image.Clone(cropRectangle, image.PixelFormat))
				{
					resultBitmap.Save($"{imageFile.FullName}");
				}
			}
			image.Dispose();
			File.Delete(tmpFile.FullName);
		}

		private static Image ResizeImage(Image sourceImage, int width, int height)
		{
			int sourceWidth = sourceImage.Width;
			int sourceHeight = sourceImage.Height;
			int sourceX = 0, sourceY = 0;
			int destX = 0, destY = 0;

			float coeff = 0;
			float coeffWidth = 0;
			float coeffHeight = 0;

			coeffWidth = ((float)width / (float)sourceWidth);
			coeffHeight = ((float)height / (float)sourceHeight);

			if (coeffHeight < coeffWidth)
			{
				coeff = coeffHeight;
				destX = System.Convert.ToInt32((width - (sourceWidth * coeff)) / 2);
			}
			else
			{
				coeff = coeffWidth;
				destY = System.Convert.ToInt32((height - (sourceHeight * coeff)) / 2);
			}

			int destWidth = (int)(sourceWidth * coeff);
			int destHeight = (int)(sourceHeight * coeff);

			Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
			bitmap.SetResolution(sourceImage.HorizontalResolution, sourceImage.VerticalResolution);

			using (Graphics graphics = Graphics.FromImage(bitmap))
			{
				graphics.Clear(System.Drawing.Color.White);
				graphics.InterpolationMode =
					InterpolationMode.HighQualityBicubic;

				graphics.DrawImage(sourceImage,
					new System.Drawing.Rectangle(destX, destY, destWidth, destHeight),
					new System.Drawing.Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
					GraphicsUnit.Pixel);
			}
			return bitmap;
		}

		private static void ZoomOpenUIViews(UIDocument uiDoc, double zoomValue, bool isToFit = true)
		{
			//IList<UIView> uiViews = uiDoc.GetOpenUIViews();
			var activeView = uiDoc
				.GetOpenUIViews()
				.FirstOrDefault(x => x.ViewId == uiDoc.ActiveView.Id);
			if (isToFit)
			{
				activeView.ZoomToFit();
			}
			activeView.Zoom(zoomValue);
			uiDoc.RefreshActiveView();
			//foreach (var item in uiViews)
			//{
			//	if(item.ViewId != uiDoc.ActiveView.Id) continue;
			//	if (isToFit)
			//	{
			//		item.ZoomToFit();
			//		uiDoc.RefreshActiveView();
			//	}
			//	item.Zoom(zoomValue);
			//	uiDoc.RefreshActiveView();
			//}
		}

		private static double GetScaleFromElement(UIDocument uiDoc)
		{

			var minPoint = new XYZPoint();
			var maxPoint = new XYZPoint();
			var viewType = uiDoc.ActiveView.ViewType;
			double scaleFactor = 1;

			GetExtremePoints(uiDoc.Document, ref minPoint, ref maxPoint);
			var rect = uiDoc.GetOpenUIViews()[0]?.GetWindowRectangle();

			if (viewType == ViewType.ThreeD)
			{
				Scale3DCalculation(rect, minPoint, maxPoint, ref scaleFactor);
				var coefficient = 1.4;
				var finalScaleFactor = coefficient * scaleFactor;
				if (finalScaleFactor < 1)
					return finalScaleFactor;
			}

			Scale2DCalculation(rect, minPoint, maxPoint, ref scaleFactor);

			scaleFactor *= 0.95;
			return scaleFactor;
		}

		private static void GetExtremePoints(Document doc, ref XYZPoint minPoint, ref XYZPoint maxPoint)
		{
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(FamilyInstance));
			BoundingBoxXYZ initialBox = collector.ToElements()[0]?.get_BoundingBox(doc.ActiveView);
			if (initialBox == null) return;
			minPoint.X = initialBox.Min.X;
			minPoint.Y = initialBox.Min.Y;
			minPoint.Z = initialBox.Min.Z;
			maxPoint.X = initialBox.Max.X;
			maxPoint.Y = initialBox.Max.Y;
			maxPoint.Z = initialBox.Max.Z;
			foreach (var item in collector)
			{
				var box = item.get_BoundingBox(doc.ActiveView);
				if (box == null)
					continue;

				if (box.Max.X > maxPoint.X)
					maxPoint.X = box.Max.X;
				if (box.Max.Y > maxPoint.Y)
					maxPoint.Y = box.Max.Y;
				if (box.Max.Z > maxPoint.Z)
					maxPoint.Z = box.Max.Z;

				if (box.Min.X < minPoint.X)
					minPoint.X = box.Min.X;
				if (box.Min.Y < minPoint.Y)
					minPoint.Y = box.Min.Y;
				if (box.Min.Z < minPoint.Z)
					minPoint.Z = box.Min.Z;

			}
		}

		private static double GetScaleFromView(Autodesk.Revit.DB.Rectangle rectangle)
		{
			var viewWidth = rectangle.Right - rectangle.Left;
			var viewHeight = rectangle.Bottom - rectangle.Top;
			if (viewWidth > viewHeight)
				return (double)viewHeight / viewWidth;
			return 1d;
		}

		private static void Scale2DCalculation(Autodesk.Revit.DB.Rectangle rectangle, 
			XYZPoint min, XYZPoint max, ref double scaleFactor)
		{
			double height = Math.Round(max.Y - min.Y, 3, MidpointRounding.AwayFromZero);
			double width = Math.Round(max.X - min.X, 3, MidpointRounding.AwayFromZero);
			if (height / width < 1)
			{
				scaleFactor = GetScaleFromView(rectangle);
			}
		}

		private static void Scale3DCalculation(Autodesk.Revit.DB.Rectangle rectangle, 
			XYZPoint min, XYZPoint max, ref double scaleFactor)
		{
			double height = Math.Round(max.Y - min.Y, 3, MidpointRounding.AwayFromZero);
			double width = Math.Round(max.X - min.X, 3, MidpointRounding.AwayFromZero);
			double depth = Math.Round(max.Z - min.Z, 3, MidpointRounding.AwayFromZero);

			var heightTotal  = Math.Abs(Math.Cos(Math.PI / 6) * width) + Math.Abs(Math.Cos(Math.PI / 6) * depth);
			var widthTotal = Math.Abs(Math.Sin(Math.PI / 6) * width) + Math.Abs(Math.Sin(Math.PI / 6) * depth) + Math.Abs(height);

			if (heightTotal / widthTotal < 1) // (widthTotal / heightTotal < 1)
			{
				scaleFactor = GetScaleFromView(rectangle); // (widthTotal / heightTotal)
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
			if (App.Version.Contains("2018")) return;
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
			HideElementsCommit(uiDoc, levelsToHide);
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
				Title = App.Translator.GetValue(Translator.Keys.fileDialogTitle),
				FilterIndex = 0,
				Filter = $"{App.Translator.GetValue(Translator.Keys.fileDialogFilter)} (*.PNG, *.JPG, *.BMP) | *.png;*.jpg;*.bmp",
				FileName = name
			};

			if (dialog.ShowDialog() == true)
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
			try
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

					string umlautName = new FileInfo(filePath).Name;
					string normalizedName = CorrectFileName(umlautName);
					filePath = filePath.Replace(umlautName, normalizedName);
					
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

					ZoomOpenUIViews(uiDoc, scale, false);

					R2019_HotFix();
					if (ImageExportOptions.IsValidFileName(filePath))
					{
						doc.ExportImage(exportOptions);
					}

					transaction.Commit();

					CropImageRectangle(userValues, imageFile, tmpFile);
				}
				doc.Dispose();
			}
			catch (Exception exc)
			{
				ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageViewPrinting)}", App.Logger, !isAuto);
			}
		}

		private static void ActiveViewChangeTransaction(UIDocument uiDoc, UserImageValues userValues, bool is3D = false)
		{
			using (Transaction transaction = new Transaction(uiDoc.Document))
			{
				transaction.Start("Set View");
				uiDoc.ActiveView.DetailLevel = is3D ? ViewDetailLevel.Fine : userValues.UserDetailLevel;
				uiDoc.ActiveView.Scale = userValues.UserScale;
				transaction.Commit();
			}
		}

		public static void SetActive2DView(UIDocument uiDoc)
		{
			using (Document doc = uiDoc.Document)
			{
				View view = null;
				FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
				viewCollector.OfClass(typeof(View));

				foreach (Element viewElement in viewCollector)
				{
					View tmpView = (View)viewElement;

					Debug.WriteLine($"Name: {tmpView.Name} ### ViewType: {tmpView.ViewType}");
					if (tmpView.Name.Equals($"{App.Translator.GetValue(Translator.Keys.level1Name)}")
						&& tmpView.ViewType == ViewType.EngineeringPlan)
					{
						view = tmpView;
					}
				}

				if (view == null)
					view = viewCollector.Cast<View>().FirstOrDefault(x => x.ViewType == ViewType.EngineeringPlan);
				if (view == null)
					view = ProjectHelper.CreateStructuralPlan(doc);
				uiDoc.ActiveView = view;
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
						trans.Start("Add View");
						view3D = View3D.CreateIsometric(doc, viewFamilyType.Id);
						trans.Commit();
					}
				}
				uiDoc.ActiveView = view3D;
			}
		}

		public static void View2DChangesCommit(UIDocument uiDoc, UserImageValues userValues)
		{
			try
			{
				using (Document doc = uiDoc.Document)
				{
					ZoomOpenUIViews(uiDoc, userValues.UserZoomValue);
					ActiveViewChangeTransaction(uiDoc, userValues);
					R2019_HotFix();
				}
			}
			catch (Exception exc)
			{
				ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageViewCorrecting)}", App.Logger, false);
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
					ActiveViewChangeTransaction(uiDoc, userValues, true);
				}
			}
			R2019_3DViewFix(uiDoc);
			R2019_HotFix();
		}

		public static string GetFileName(Document doc)
		{
			if (App.Version.Contains("2018"))
			{
				int indexDot = doc.Title.IndexOf('.');
				var name = doc.Title.Substring(0, indexDot);
				return name;
			}

			return doc.Title;
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
			if (newDocPath.Equals(uiDoc.Application.ActiveUIDocument.Document?.PathName) ||
				!File.Exists(newDocPath)) return uiDoc;
			UIDocument result = uiDoc.Application.OpenAndActivateDocument(newDocPath);
			if (!IsDocumentActive(uiDoc))
				uiDoc.Document.Close(false);
			return result;
		}

		public static void HideElementsCommit(UIDocument uiDoc, ICollection<ElementId> elements)
		{
			using (Transaction transaction = new Transaction(uiDoc.Document, "Level Isolating"))
			{
				transaction.Start();
				uiDoc.ActiveView.HideElements(elements);
				transaction.Commit();
			}
		}

		public static string CorrectFileName(string fileName)
		{
			fileName = fileName.Replace("Ø", "D");
			fileName = fileName.Replace("Ä", "AE");
			fileName = fileName.Replace("ä", "ae");
			fileName = fileName.Replace("Ö", "OE");
			fileName = fileName.Replace("ö", "oe");
			fileName = fileName.Replace("Ü", "UE");
			fileName = fileName.Replace("ü", "ue");
			fileName = fileName.Replace("ß", "ss");
			fileName = fileName.Replace("°", "");
			fileName = fileName.Replace(' ', '_');
			fileName = fileName.Replace('/', '_');
			return fileName;
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

		// public static void CheckImagesAmount(DirectoryInfo projectsDir, int imagesCreated)
		// {
		//     var projectsCreated = projectsDir.GetFiles().Where(x => x.Extension.Equals(".rvt")).ToList();
		//        try
		//     {
		//         Assert.AreEqual(projectsCreated.Count, imagesCreated);
		//     }
		//     catch (AssertFailedException exc)
		//     {
		//      uint difference = (uint) (projectsCreated.Count - imagesCreated);
		//      string errorMsg = App.Translator.GetValue(Translator.Keys.errorMessageDuringPrinting)
		//       .Replace("%amount%", difference.ToString());
		//ProcessError(exc, $"{errorMsg}", App.Logger);
		//     }
		//    }

		public static void ProcessError(Exception exc, string errorMessage, Logger logger, bool isDialog = true)
		{
			if (isDialog)
			{
				new Autodesk.Revit.UI.TaskDialog($"{App.Translator.GetValue(Translator.Keys.errorMessageTitle)}")
				{
					TitleAutoPrefix = false,
					MainIcon = Autodesk.Revit.UI.TaskDialogIcon.TaskDialogIconError,
					MainContent = errorMessage
				}.Show();
			}
			logger.WriteLine($"### ERROR ### - {errorMessage}\n{exc.Message}\n{exc.StackTrace}");
		}

		#endregion
	}
}
