using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
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
using Newtonsoft.Json;
using System.IO;

namespace RevitFamilyImagePrinter.Windows
{
	/// <summary>
	/// Interaction logic for PrintOptions.xaml
	/// </summary>
	public partial class PrintOptions : UserControl
	{
		#region Properties
		public int UserImageSize { get; set; }
		public int UserScale { get; set; }
		public ImageResolution UserImageResolution { get; set; }
		public double UserZoomValue { get; set; }
		public string UserExtension { get; set; }
		public ViewDetailLevel UserDetailLevel { get; set; }

		public Document Doc { get; set; }
		public UIDocument UIDoc { get; set; }
		public bool IsPreview { get; set; }
		public bool IsCancelled { get; set; }
		public bool IsUpdateView { get; set; }
		public bool Is3D { get; set; }
		#endregion

		#region Constants
		private const string configName = "config.json";
		private readonly string _configPath = System.IO.Path.Combine(App.DefaultFolder, configName);
		#endregion

		#region Variables
		private Window _parentWindow;
		#endregion


		public PrintOptions()
		{
			InitializeComponent();
		}

		#region Events
		private void Button_Click_Apply(object sender, RoutedEventArgs e)
		{
			Apply();
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
			UserExtension = (sender as RadioButton)?.Tag.ToString();
		}

		private void UserControl_Loaded(object sender, RoutedEventArgs e)
		{
			_parentWindow = Window.GetWindow(this);
			btnApply.Visibility = IsPreview ? System.Windows.Visibility.Visible
											: System.Windows.Visibility.Hidden;
			LoadConfig();
			CorrectValues();
			SizeValue.Focus();
			SizeValue.SelectAll();
		}

		private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if ((sender as System.Windows.Controls.ComboBox) != ResolutionValue
					&& btnApply != null)
			{
				btnApply.IsEnabled = true;
			}
		}
		#endregion

		#region Methods

		private bool Apply()
		{
			UserImageValues userValues = GetRoughValuesFromFields();
			if (userValues == null)
			{
				IsCancelled = true;
				return false;
			}
			SaveConfig();
			InitializeUserFields(userValues);
			CorrectValues();
			btnApply.IsEnabled = false;
			if (IsUpdateView)
				UpdateView();
			return true;
		}

		private UserImageValues GetRoughValuesFromFields()
		{
			try
			{
				string strResolution = (this.ResolutionValue.SelectedItem as FrameworkElement)?.Tag.ToString();
				string strDetailLevel = (this.DetailLevelValue.SelectedItem as FrameworkElement)?.Tag.ToString();
				return new UserImageValues()
				{
					UserImageSize = int.Parse(this.SizeValue.Text),
					UserScale = int.Parse(this.ScaleValue.Text),
					UserZoomValue = double.Parse(this.ZoomValue.Text),
					UserImageResolution = GetImageResolution(strResolution),
					UserDetailLevel = GetUserDetailLevel(strDetailLevel),
					UserExtension = this.UserExtension
				};
			}
			catch (Exception exc)
			{
				string errorMessage = $"Error occured during values correction.\n{exc.Message}";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainIcon = TaskDialogIcon.TaskDialogIconError,
					MainContent = errorMessage
				}.Show();
				return null;
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
			try
			{
				if (UserImageSize > 2048)
					UserImageSize = 2048;
				if (UserImageSize < 32)
					UserImageSize = 32;
				if (UserZoomValue > 100)
					UserZoomValue = 100;
				if (UserScale < 1 || UserImageSize < 1 || UserZoomValue <= 0)
					throw new InvalidCastException("The value cannot be zero or less than zero.");
				UserZoomValue = Math.Round(UserZoomValue) / 100;
				FitUserScale();
			}
			catch (Exception exc)
			{
				string errorMessage = $"Error occured during values correction.\n{exc.Message}";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainIcon = TaskDialogIcon.TaskDialogIconError,
					MainContent = errorMessage
				}.Show();
			}
		}

		private void UpdateView()
		{
			try
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
					Doc.ActiveView.DetailLevel = Is3D ? ViewDetailLevel.Fine : UserDetailLevel;
					Doc.ActiveView.Scale = UserScale;
					transaction.Commit();
				}
			}
			catch (Exception exc)
			{
				string errorMessage = $"### ERROR ### - Error occured during view update.\n{exc.Message}";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainIcon = TaskDialogIcon.TaskDialogIconError,
					MainContent = errorMessage
				}.Show();
			}
		}

		private void Print()
		{
			_parentWindow.DialogResult = Apply();
		}

		public void SaveConfig()
		{
			try
			{
				UserImageValues userRoughValues = GetRoughValuesFromFields();
				string jsonStr = JsonConvert.SerializeObject(userRoughValues);
				File.WriteAllText(_configPath, jsonStr);
			}
			catch (Exception exc)
			{
				string errorMessage = $"Error during saving user input values: {exc.Message}";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainIcon = TaskDialogIcon.TaskDialogIconError,
					MainContent = errorMessage
				}.Show();
			}
		}

		private void LoadConfig()
		{
			try
			{
				if (!File.Exists(_configPath))
				{
					InitializeUserFields(new UserImageValues()
					{
						UserDetailLevel = ViewDetailLevel.Medium,
						UserImageSize = 64,
						UserExtension = ".png",
						UserScale = 50,
						UserImageResolution = ImageResolution.DPI_150,
						UserZoomValue = 90
					});
					return;
				}
				string jsonStr = File.ReadAllText(_configPath);
				UserImageValues userRoughValues = JsonConvert.DeserializeObject<UserImageValues>(jsonStr);
				InitializeUserFields(userRoughValues);
				SetInitialFieldValues();
			}
			catch (Exception exc)
			{
				string errorMessage = $"Error during loading user input values: {exc.Message}";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainIcon = TaskDialogIcon.TaskDialogIconError,
					MainContent = errorMessage
				}.Show();
			}
		}

		private void InitializeUserFields(UserImageValues userValues)
		{
			this.UserImageResolution = userValues.UserImageResolution;
			this.UserImageSize = userValues.UserImageSize;
			this.UserScale = userValues.UserScale;
			this.UserDetailLevel = userValues.UserDetailLevel;
			this.UserExtension = userValues.UserExtension;
			this.UserZoomValue = userValues.UserZoomValue;
		}

		private void SetInitialFieldValues()
		{
			SizeValue.Text = UserImageSize.ToString();
			ScaleValue.Text = UserScale.ToString();
			ZoomValue.Text = UserZoomValue.ToString(CultureInfo.InvariantCulture);
			ResolutionValue.SelectedIndex = GetResolutionItemIndex();
			DetailLevelValue.SelectedIndex = GetDetailingItemIndex();
			SetRadioButtonChecked();
		}

		private void SetRadioButtonChecked()
		{
			RadioButton btn = null;
			switch (UserExtension)
			{
				case ".png": btn = RadioButtonPng; break;
				case ".jpg": btn = RadioButtonJpg; break;
				case ".bmp": btn = RadioButtonBmp; break;
				default: throw new Exception("Unknown extension");
			}
			btn.IsChecked = true;
		}

		private int GetResolutionItemIndex()
		{
			switch (UserImageResolution)
			{
				case ImageResolution.DPI_72: return 0;
				case ImageResolution.DPI_150: return 1;
				case ImageResolution.DPI_300: return 2;
				case ImageResolution.DPI_600: return 3;
				default: throw new Exception("Unknown resolution type");
			}
		}

		private int GetDetailingItemIndex()
		{
			switch (UserDetailLevel)
			{
				case ViewDetailLevel.Coarse: return 0;
				case ViewDetailLevel.Medium: return 1;
				case ViewDetailLevel.Fine: return 2;
				default: throw new Exception("Unknown detailing type");
			}
		}
		#endregion
	}
}
