using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.CompilerServices;
using RevitFamilyImagePrinter.Infrastructure;

namespace RevitFamilyImagePrinter.Windows
{
	/// <summary>
	/// Interaction logic for PrintOptions.xaml
	/// </summary>
	public partial class PrintOptions : UserControl, INotifyPropertyChanged
	{
		#region Properties
		public int UserImageHeight { get; set; }
		public int UserScale { get; set; }
		public ImageResolution UserImageResolution { get; set; }
		public double UserZoomValue { get; set; }
		public string UserExtension { get; set; }
		public ViewDetailLevel UserDetailLevel { get; set; }
		public ImageAspectRatio UserAspectRatio { get; set; }

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
		private Logger _logger = App.Logger;
		#endregion

		#region Labels

		public string labelSize_Text => App.Translator.GetValue(Translator.Keys.labelSize_Text);
		public string labelScale_Text => App.Translator.GetValue(Translator.Keys.labelScale_Text);
		public string labelZoom_Text => App.Translator.GetValue(Translator.Keys.labelZoom_Text);
		public string labelResolution_Text => App.Translator.GetValue(Translator.Keys.labelResolution_Text);
		public string labelDetailLevel_Text => App.Translator.GetValue(Translator.Keys.labelDetailLevel_Text);
		public string labelFormat_Text => App.Translator.GetValue(Translator.Keys.labelFormat_Text);
		public string labelAspectRatio_Text => App.Translator.GetValue(Translator.Keys.labelAspectRatio_Text);
		public string labelParameters_Text => App.Translator.GetValue(Translator.Keys.labelParameters_Text);

		public string textBlockResolutionWebLow_Text => App.Translator.GetValue(Translator.Keys.textBlockResolutionWebLow_Text);
		public string textBlockResolutionWebHigh_Text => App.Translator.GetValue(Translator.Keys.textBlockResolutionWebHigh_Text);
		public string textBlockResolutionPrintLow_Text => App.Translator.GetValue(Translator.Keys.textBlockResolutionPrintLow_Text);
		public string textBlockResolutionPrintHigh_Text => App.Translator.GetValue(Translator.Keys.textBlockResolutionPrintHigh_Text);

		public string textBlockDetailLevelCoarse_Text => App.Translator.GetValue(Translator.Keys.textBlockDetailLevelCoarse_Text);
		public string textBlockDetailLevelMedium_Text => App.Translator.GetValue(Translator.Keys.textBlockDetailLevelMedium_Text);
		public string textBlockDetailLevelFine_Text => App.Translator.GetValue(Translator.Keys.textBlockDetailLevelFine_Text);

		public string buttonApply_Text => App.Translator.GetValue(Translator.Keys.buttonApply_Text);
		public string buttonPrint_Text => App.Translator.GetValue(Translator.Keys.buttonPrint_Text);

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
			buttonApply.IsEnabled = true;
			if (e.Key == Key.Enter)
				Print();
		}

		private void RadioButtonExtension_Checked(object sender, RoutedEventArgs e)
		{
			UserExtension = (sender as RadioButton)?.Tag.ToString();
		}

		private void RadioButtonRatio_Checked(object sender, RoutedEventArgs e)
		{
			UserAspectRatio = GetAspectRatio((sender as RadioButton)?.Tag.ToString());
		}

		private void UserControl_Loaded(object sender, RoutedEventArgs e)
		{
			_parentWindow = Window.GetWindow(this);
			buttonApply.Visibility = IsPreview ? System.Windows.Visibility.Visible
											: System.Windows.Visibility.Hidden;
			LoadConfig();
			CorrectValues();
			SizeValue.Focus();
			SizeValue.SelectAll();
		}

		private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if ((sender as System.Windows.Controls.ComboBox) != comboBoxResolutionValue
					&& buttonApply != null)
			{
				buttonApply.IsEnabled = true;
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
			buttonApply.IsEnabled = false;
			if (IsUpdateView)
				UpdateView();
			return true;
		}

		private UserImageValues GetRoughValuesFromFields()
		{
			try
			{
				string strResolution = (this.comboBoxResolutionValue.SelectedItem as FrameworkElement)?.Tag.ToString();
				string strDetailLevel = (this.comboBoxDetailLevelValue.SelectedItem as FrameworkElement)?.Tag.ToString();
				return new UserImageValues()
				{
					UserImageHeight = int.Parse(this.SizeValue.Text),
					UserScale = int.Parse(this.ScaleValue.Text),
					UserZoomValue = double.Parse(this.ZoomValue.Text),
					UserImageResolution = GetImageResolution(strResolution),
					UserDetailLevel = GetUserDetailLevel(strDetailLevel),
					UserExtension = this.UserExtension,
					UserAspectRatio = this.UserAspectRatio
				};
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageValuesRetrieving)}", _logger);
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

			if (UserScale > 256)
			{
				UserScale = 10;
				return;
			}

			if (UserScale > 32)
			{
				UserScale = 25;
				return;
			}

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

		private ImageAspectRatio GetAspectRatio(string stringRatio)
		{
			switch (stringRatio)
			{
				case "16:9": return ImageAspectRatio.Ratio_16to9;
				case "4:3": return ImageAspectRatio.Ratio_4to3;
				default: case "1:1": return ImageAspectRatio.Ratio_1to1;
			}
		}

		private ImageResolution GetImageResolution(string resolution)
		{
			switch (resolution)
			{
				default: case "72": return ImageResolution.DPI_72;
				case "150": return ImageResolution.DPI_150;
				case "300": return ImageResolution.DPI_300;
				case "600": return ImageResolution.DPI_600;
			}
		}

		private void CorrectValues()
		{
			try
			{
				if (UserImageHeight > 2048)
					UserImageHeight = 2048;
				if (UserImageHeight < 32)
					UserImageHeight = 32;
				if (UserZoomValue > 100)
					UserZoomValue = 100;
				if (UserScale < 1 || UserImageHeight < 1 || UserZoomValue <= 0)
					throw new InvalidCastException("The value cannot be zero or less than zero.");
				UserZoomValue = Math.Round(UserZoomValue) / 100;
				FitUserScale();
			}
			catch (Exception exc)
			{
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageValuesCorrection)}", _logger);
			}
		}

		private void UpdateView()
		{
			try
			{
				//RevitPrintHelper.View2DChangesCommit(UIDoc, UserImageValues);
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
				RevitPrintHelper.ProcessError(exc, 
					$"{App.Translator.GetValue(Translator.Keys.errorMessageViewUpdating)}", _logger);
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
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageValuesSaving)}", _logger);
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
						UserImageHeight = 64,
						UserExtension = ".png",
						UserScale = 50,
						UserImageResolution = ImageResolution.DPI_150,
						UserZoomValue = 90,
						UserAspectRatio = ImageAspectRatio.Ratio_1to1
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
				RevitPrintHelper.ProcessError(exc,
					$"{App.Translator.GetValue(Translator.Keys.errorMessageValuesLoading)}", _logger);
			}
		}

		private void InitializeUserFields(UserImageValues userValues)
		{
			this.UserImageResolution = userValues.UserImageResolution;
			this.UserImageHeight = userValues.UserImageHeight;
			this.UserScale = userValues.UserScale;
			this.UserDetailLevel = userValues.UserDetailLevel;
			this.UserExtension = userValues.UserExtension;
			this.UserZoomValue = userValues.UserZoomValue;
			this.UserAspectRatio = userValues.UserAspectRatio;
		}

		private void SetInitialFieldValues()
		{
			SizeValue.Text = UserImageHeight.ToString();
			ScaleValue.Text = UserScale.ToString();
			ZoomValue.Text = UserZoomValue.ToString(CultureInfo.InvariantCulture);
			comboBoxResolutionValue.SelectedIndex = GetResolutionItemIndex();
			comboBoxDetailLevelValue.SelectedIndex = GetDetailingItemIndex();
			SetRadioButtonExtensionChecked();
			SetRadioButtonRatioChecked();
		}

		private void SetRadioButtonExtensionChecked()
		{
			RadioButton btn = null;
			switch (UserExtension)
			{
				default: case ".png": btn = RadioButtonPng; break;
				case ".jpg": btn = RadioButtonJpg; break;
				case ".bmp": btn = RadioButtonBmp; break;
			}
			btn.IsChecked = true;
		}

		private void SetRadioButtonRatioChecked()
		{
			RadioButton btn = null;
			switch (UserAspectRatio)
			{
				default: case ImageAspectRatio.Ratio_1to1: btn = RadioButton1to1; break;
				case ImageAspectRatio.Ratio_16to9: btn = RadioButton16to9; break;
				case ImageAspectRatio.Ratio_4to3: btn = RadioButton4to3; break;
			}
			btn.IsChecked = true;
		}

		private int GetResolutionItemIndex()
		{
			switch (UserImageResolution)
			{
				default: case ImageResolution.DPI_72: return 0;
				case ImageResolution.DPI_150: return 1;
				case ImageResolution.DPI_300: return 2;
				case ImageResolution.DPI_600: return 3;
			}
		}

		private int GetDetailingItemIndex()
		{
			switch (UserDetailLevel)
			{
				case ViewDetailLevel.Coarse: return 0;
				default: case ViewDetailLevel.Medium: return 1;
				case ViewDetailLevel.Fine: return 2;
			}
		}
		#endregion

		public event PropertyChangedEventHandler PropertyChanged;

		protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}
	}
}
