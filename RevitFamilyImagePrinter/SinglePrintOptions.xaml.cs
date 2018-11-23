using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace RevitFamilyImagePrinter
{
	/// <summary>
	/// Interaction logic for SinglePrintOptions.xaml
	/// </summary>
	public partial class SinglePrintOptions : UserControl
	{
		#region Variables
		public int UserImageSize { get; set; }
		public int UserScale { get; set; }
		public ImageResolution UserImageResolution { get; set; }
		public double UserZoomValue { get; set; }
		public string UserExtension { get; set; } = ".png";
		public ViewDetailLevel UserDetailLevel { get; set; }

		private Window parentWindow;
		public Document Doc { get; set; }
		public UIDocument UIDoc { get; set; }
		#endregion

		public SinglePrintOptions()
		{
			InitializeComponent();
		}

		#region Events
		private void Button_Click_Apply(object sender, RoutedEventArgs e)
		{
			ApplyUserValues();
			btnApply.IsEnabled = false;
			UpdateView();
		}

		private void Button_Click_Print(object sender, RoutedEventArgs e)
		{
			Print();
		}

		private void TextBox_KeyDown(object sender, KeyEventArgs e)
		{
			btnApply.IsEnabled = true;
			if (e.Key == Key.Enter)
				Print();
		}

		private void RadioButton_Checked(object sender, RoutedEventArgs e)
		{
			UserExtension = (sender as RadioButton).Tag.ToString();
		}

		private void UserControl_Loaded(object sender, RoutedEventArgs e)
		{
			parentWindow = Window.GetWindow(this);
			SizeValue.Focus();
		}

		private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if ((sender as System.Windows.Controls.ComboBox) != ResolutionValue && btnApply != null)
				btnApply.IsEnabled = true;
		}
		#endregion

		#region Methods

		private void ApplyUserValues()
		{
			try
			{
				string strResolution = (this.ResolutionValue.SelectedItem as FrameworkElement).Tag.ToString();
				string strDetailLevel = (this.DetailLevelValue.SelectedItem as FrameworkElement).Tag.ToString();
				UserImageSize = int.Parse(this.SizeValue.Text);
				UserScale = int.Parse(this.ScaleValue.Text);
				UserZoomValue = double.Parse(this.ZoomValue.Text);
				UserImageResolution = GetImageResolution(strResolution);
				UserDetailLevel = GetUserDetailLevel(strDetailLevel);
				CorrectValues();
			}
			catch
			{
				TaskDialog.Show("Error", "Invalid input values. Please, try again.");
			}

		}

		private void FitUserScale()
		{
			if (UserScale > 512)
			{
				UserScale = 1;
				return;
			}
			else if (UserScale > 256)
			{
				UserScale = 10;
				return;
			}
			else if (UserScale > 32)
			{
				UserScale = 25;
				return;
			}
			else
				UserScale = 50;
		}

		private void FitUserZoom()
		{
			UserZoomValue = Math.Round(UserZoomValue) / 100;
		}

		private ViewDetailLevel GetUserDetailLevel(string strDetailLevel)
		{
			switch (strDetailLevel)
			{
				case "Coarse": return ViewDetailLevel.Coarse;
				case "Medium": return ViewDetailLevel.Medium;
				case "Fine": return ViewDetailLevel.Fine;
				default: return ViewDetailLevel.Undefined;
			}
		}

		private ImageResolution GetImageResolution(string resolution)
		{
			switch (resolution)
			{
				case "72": return ImageResolution.DPI_72;
				case "150": return ImageResolution.DPI_150;
				case "300": return ImageResolution.DPI_300;
				case "600": return ImageResolution.DPI_600;
				default: throw new Exception("Unknown Resolution Type");
			}
		}

		private void CorrectValues()
		{
			if (UserImageSize > 2048)
				UserImageSize = 2048;
			if (UserImageSize < 32)
				UserImageSize = 32;
			if (UserZoomValue > 100)
				UserZoomValue = 100;
			if (UserScale < 1 || UserImageSize < 1 || UserZoomValue <= 0)
				throw new InvalidCastException("The value cannot be zero or less than zero.");
			FitUserZoom();
			FitUserScale();
		}

		private void UpdateView()
		{
			IList<UIView> uiviews = UIDoc.GetOpenUIViews();
			foreach (var item in uiviews)
			{
				item.ZoomToFit();
				item.Zoom(UserZoomValue);
				UIDoc.RefreshActiveView();
			}

			using (Transaction transaction = new Transaction(Doc))
			{
				transaction.Start("SetView");
				Doc.ActiveView.DetailLevel = UserDetailLevel;
				Doc.ActiveView.Scale = UserScale;
				transaction.Commit();
			}
		}

		private void Print()
		{
			ApplyUserValues();
			parentWindow.DialogResult = true;
		}
		#endregion
	}
}
