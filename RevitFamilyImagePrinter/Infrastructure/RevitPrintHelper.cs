using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Ookii.Dialogs.Wpf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace RevitFamilyImagePrinter
{
	public static class RevitPrintHelper
	{
		public static UserImageValues ShowOptionsDialog(UIDocument uiDoc, int windowHeightOffset = 40, int windowWidthOffset = 10, bool isApplyButtonVisible = true)
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
			if(!options.IsPreview)
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
