using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using RevitFamilyImagePrinter.Infrastructure;
using View = System.Windows.Forms.View;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print3D : IExternalCommand
	{
		#region User Values

		public UserImageValues UserValues { get; set; } = new UserImageValues()
		{
			UserScale = 100, UserZoomValue = 0.9
		};
		public string UserImagePath { get; set; } 

		#endregion

		#region Variables
		//private IList<ElementId> views = new List<ElementId>();
		private Document _doc => UIDoc?.Document;
		private readonly Logger _logger = Logger.GetLogger();
		public UIDocument UIDoc;
		public bool IsAuto = false;
		#endregion

		#region Constants
		private const int windowHeightOffset = 40;
		private const int windowWidthOffset = 20;
		private readonly string endl = Environment.NewLine;
		#endregion

		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			UIApplication uiapp = commandData.Application;
			UIDoc = uiapp.ActiveUIDocument;
			//uiapp.ViewActivated += SetViewParameters;
			
			RevitPrintHelper.SetActive3DView(UIDoc);
			ViewChangesCommit();

			if (!message.Equals("FolderPrint"))
			{
				UserImageValues userInputValues = 
					RevitPrintHelper.ShowOptionsDialog(UIDoc, windowHeightOffset, windowWidthOffset, true);
				if (userInputValues == null)
					return Result.Cancelled;
				this.UserValues = userInputValues;
				ViewChangesCommit();
			}

			PrintCommit();

			return Result.Succeeded;
		}

		public void ViewChangesCommit()
		{
			try
			{
				RevitPrintHelper.View3DChangesCommit(UIDoc, UserValues);
			}
			catch (Exception exc)
			{
				string errorMessage = "### ERROR ### - Error occured during current view correction";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"{errorMessage}{endl}{exc.Message}{endl}{exc.StackTrace}");
			}
		}

		private void PrintCommit()
		{
			try
			{
				RevitPrintHelper.PrintImageTransaction(UIDoc, UserValues, UserImagePath, IsAuto);
			}
			catch (Exception exc)
			{
				string errorMessage = "### ERROR ### - Error occured during printing of current view";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"{errorMessage}{endl}{exc.Message}{endl}{exc.StackTrace}");
			}
		}
	}
}
