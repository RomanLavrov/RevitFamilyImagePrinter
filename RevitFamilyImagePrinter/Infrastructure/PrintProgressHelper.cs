using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace RevitFamilyImagePrinter.Infrastructure
{
    internal class PrintProgressHelper
    {
	    public PrintProgressHelper(DirectoryInfo familiesFolder, string processText)
	    {
		    _printProgress = new PrintProgress();
		    _processTextBlock = _printProgress.ProcessTextBlock;
		    _printProgressBar = _printProgress.PrintProgressBar;
		    _processTextBlock.Text = processText;
		    _familiesFolder = familiesFolder;
		    _familiesAmount = familiesFolder.GetFiles().Count(x => x.Extension.Equals(".rfa"));
		    PreviousViewName = string.Empty;
	    }

		//  ~PrintProgressHelper()
	  //  {
			//this.Close();
	  //  }

		public string PreviousViewName { get; set; }

		#region Methods

		public Window Show(bool is3D = false, int windowHeightOffset = 40, int windowWidthOffset = 20)
	    {
		    _progressWindow = new Window
		    {
			    Height = 100 + windowHeightOffset,
			    Width = 400 + windowWidthOffset,
			    Title = $"{(is3D ? "3D" : "2D")} printing of \"{_familiesFolder.Name}\" folder",
			    WindowStyle = WindowStyle.ToolWindow,
			    Name = "Printing",
			    ResizeMode = ResizeMode.NoResize,
			    WindowStartupLocation = WindowStartupLocation.CenterScreen,
			    Content = _printProgress,
			    ShowActivated = true,
			    Topmost = true
		    };
		    _progressWindow.Show();
		    return _progressWindow;
	    }

	    public void Close()
	    {
		    if (_progressWindow == null) return;
			if(_progressWindow.IsVisible)
				_progressWindow.Close();
	    }

	    public void SetProgressBarMaximum(int max)
	    {
		    _printProgressBar.Value = 0;
		    _printProgressBar.Maximum = max;
	    }

	    public void SetProgressText(string text)
	    {
		    _processTextBlock.Text = text;
			R2018_HotFix();
		}

	    public void SubscribeOnLoadedFamily(UIApplication uiApp)
	    {
		    uiApp.Application.FamilyLoadedIntoDocument += ApplicationOnFamilyLoadedIntoDocument;
	    }

	    public void SubscribeOnViewActivated(UIApplication uiApp, bool is3D = false)
	    {
		    if (is3D)
			    uiApp.ViewActivated += ApplicationOn3DViewActivated;
			else
				uiApp.ViewActivated += ApplicationOnViewActivated;
	    }

	    private void R2018_HotFix()
	    {
		    if (App.Version != "2018") return;
		    var tmpWindow = new Window()
		    {
			    Width = 10,
			    Height = 10,
			    ShowInTaskbar = false,
				ShowActivated = false,
			    WindowStyle = WindowStyle.None
			};
		    tmpWindow.Show();
		    tmpWindow.Close();
		}

		#endregion

		#region Events

		private void ApplicationOnViewActivated(object sender, ViewActivatedEventArgs e)
	    {
		    string viewName = RevitPrintHelper.GetFileName(e.CurrentActiveView.Document);
		    if (PreviousViewName.Equals(viewName) || viewName.ToLower().Equals("empty")) return;
		    _printProgressBar.Value++;
		    _processTextBlock.Text = $"Printed: {_printProgressBar.Value} / {_printProgressBar.Maximum}";
		    PreviousViewName = viewName;
	    }

	    private void ApplicationOn3DViewActivated(object sender, ViewActivatedEventArgs e)
	    {
		    if (e.CurrentActiveView.ViewType != ViewType.ThreeD) return;
		    ApplicationOnViewActivated(sender, e);
	    }

		private void ApplicationOnFamilyLoadedIntoDocument(object sender, FamilyLoadedIntoDocumentEventArgs e) 
	    {
		    _printProgressBar.Value++;
		    _processTextBlock.Text = $"Loaded families: {_printProgressBar.Value} / {_familiesAmount}";
		    R2018_HotFix();
	    }

	    private void ApplicationOnDocumentSavedAs(object sender, DocumentSavedAsEventArgs e)
	    {
		    _printProgressBar.Value++;
		    _processTextBlock.Text = $"{e.Document.Title} has been created";
	    }

		#endregion

	    #region Variables

	    private Window _progressWindow;
	    private PrintProgress _printProgress;
	    private TextBlock _processTextBlock;
	    private ProgressBar _printProgressBar;
	    private DirectoryInfo _familiesFolder;
	    private int _familiesAmount;

	    #endregion
    }
}
