using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Application = Autodesk.Revit.ApplicationServices.Application;
using View = Autodesk.Revit.DB.View;
using Ookii.Dialogs.Wpf;
using RevitFamilyImagePrinter.Infrastructure;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print2D : IExternalCommand
	{
		#region User Values
		public UserImageValues UserValues { get; set; } = new UserImageValues();
		public string UserImagePath { get; set; }
		#endregion

		#region Variables
		//private IList<ElementId> views = new List<ElementId>();
		private Document doc => UIDoc?.Document;
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

			RevitPrintHelper.SetActive2DView(UIDoc);

			if (!message.Equals("FolderPrint"))
			{
				UserImageValues userInputValues = RevitPrintHelper.ShowOptionsDialog(UIDoc, windowHeightOffset, windowWidthOffset);
				if (userInputValues == null)
					return Result.Cancelled;
				this.UserValues = userInputValues;
			}

			ViewChangesCommit();
			PrintCommit();

			return Result.Succeeded;
		}

		public void ViewChangesCommit()
		{
			try
			{
				RevitPrintHelper.View2DChangesCommit(UIDoc, UserValues);
			}
			catch (Exception exc)
			{
				string errorMessage = "### ERROR ### - Error occured during current view correction";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"{errorMessage}{endl}{exc.Message}");
			}
		}

		private void PrintCommit()
		{
			try
			{
				RevitPrintHelper.PrintImageTransaction(doc, UserValues, UserImagePath, IsAuto);
			}
			catch (Exception exc)
			{
				string errorMessage = "### ERROR ### - Error occured during printing of current view";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"{errorMessage}{endl}{exc.Message}");
			}
		}
	}
}
