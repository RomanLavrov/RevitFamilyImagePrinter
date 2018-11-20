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
        public int userImageSize { get; set; }
        public int userScale { get; set; }

        public SinglePrintOptions()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var parentWindow = Window.GetWindow(this);
            userImageSize = Int32.Parse(this.SizeValue.Text);
            userScale = Int32.Parse(this.ScaleValue.Text);
            parentWindow.DialogResult = true;
        }
    }
}
