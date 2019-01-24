using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitFamilyImagePrinter.Infrastructure;

namespace RevitFamilyImagePrinter.Commands
{
	[Transaction(TransactionMode.Manual)]
	class Print3D : IExternalCommand
	{
		#region User Values

		public UserImageValues UserValues { get; set; } = new UserImageValues()
		{
			UserScale = 100,
			UserZoomValue = 0.9
		};
		public string UserImagePath { get; set; }

		#endregion

		#region Variables

		private readonly Logger _logger = App.Logger;
		public UIDocument UIDoc;

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

			PrintHelper.SetActive3DView(UIDoc);
			PrintHelper.View3DChangesCommit(UIDoc, UserValues);

			UserImageValues userInputValues =
				PrintHelper.ShowOptionsDialog(UIDoc, windowHeightOffset, windowWidthOffset, true);
			if (userInputValues == null)
				return Result.Cancelled;
			this.UserValues = userInputValues;

			PrintHelper.View3DChangesCommit(UIDoc, UserValues);
			PrintHelper.PrintImageTransaction(UIDoc, UserValues, UserImagePath);

			return Result.Succeeded;
		}
	}
}
