using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
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
