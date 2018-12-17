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
			commandData.Application.ViewActivated += SetViewParameters;

			RevitCommandId commandId = RevitCommandId.LookupPostableCommandId(PostableCommand.Default3DView);
			if (commandData.Application.CanPostCommand(commandId))
			{
				commandData.Application.PostCommand(commandId);
			}

			if (!message.Equals("FamilyPrint"))
			{
				UserImageValues userInputValues = RevitPrintHelper.ShowOptionsDialog(UIDoc, windowHeightOffset, windowWidthOffset);
				if (userInputValues == null)
					return Result.Failed;
				this.UserValues = userInputValues;
			}

			return Result.Succeeded;
		}

		private void SetViewParameters(object sender, ViewActivatedEventArgs args)
		{
			try
			{
				FilteredElementCollector collector = new FilteredElementCollector(doc);
				collector.OfClass(typeof(View3D));

				View3D view3D = Get3DView();
				if (view3D == null)
				{
					var collectorF = new FilteredElementCollector(doc);
					var viewFamilyType = collectorF.OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
						.FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);
					using (Transaction trans = new Transaction(doc))
					{
						trans.Start("Add view");
						view3D = View3D.CreateIsometric(doc, viewFamilyType.Id);
						trans.Commit();
					}
				}

				//UIDoc.RequestViewChange(UIDoc.ActiveView);
				UIDoc.RefreshActiveView();
				UIDoc.ActiveView = view3D;

				IList<UIView> uiViews = UIDoc.GetOpenUIViews();
				foreach (var item in uiViews)
				{
					item.ZoomToFit();
					item.Zoom(UserValues.UserZoomValue);
					UIDoc.RefreshActiveView();
				}

				using (Transaction transaction = new Transaction(doc))
				{
					transaction.Start("SetView");
					UIDoc.ActiveView.DetailLevel = ViewDetailLevel.Fine;
					UIDoc.ActiveView.Scale = UserValues.UserScale;
					transaction.Commit();
				}

				using (Transaction transaction = new Transaction(doc))
				{
					transaction.Start("Print");
					RevitPrintHelper.PrintImage(doc, UserValues, string.Empty);
					transaction.Commit();
				}
			}
			catch (Exception exc)
			{
				string errorMessage = "### ERROR ### Error occured during printing of current view";
				new TaskDialog("Error")
				{
					TitleAutoPrefix = false,
					MainContent = errorMessage
				}.Show();
				_logger.WriteLine($"{errorMessage}{endl}{exc.Message}");
			}
		}

		private View3D Get3DView()
		{
			FilteredElementCollector collector
				= new FilteredElementCollector(doc)
					.OfClass(typeof(View3D));

			foreach (View3D view in collector)
			{
				if (view == null || view.IsTemplate) continue;
				return view;
			}
			return null;
		}
	}
}
