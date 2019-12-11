using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using RevitFamilyImagePrinter.Infrastructure;

namespace RevitFamilyImagePrinter
{
	/// <summary>
	/// Interaction logic for PrintProgress.xaml
	/// </summary>
	public partial class PrintProgress : UserControl, INotifyPropertyChanged
	{
		#region Labels

		public string buttonCancel_Text => App.Translator.GetValue(Translator.Keys.buttonCancel_Text);
		public string textBlockProcess_Text => App.Translator.GetValue(Translator.Keys.textBlockProcess_Text);

		#endregion

		public PrintProgress()
		{
			InitializeComponent();
		}

		public event PropertyChangedEventHandler PropertyChanged;

		protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}
	}
}
