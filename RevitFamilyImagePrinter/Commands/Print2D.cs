using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitFamilyImagePrinter.Infrastructure;

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
			PrintHelper.CreateEmptyProject(commandData.Application.Application);

			PrintHelper.SetActive2DView(UIDoc);

			UserImageValues userInputValues = PrintHelper.ShowOptionsDialog(UIDoc, windowHeightOffset, windowWidthOffset);
			if (userInputValues == null)
				return Result.Cancelled;
			this.UserValues = userInputValues;

			PrintHelper.View2DChangesCommit(UIDoc, UserValues);
			PrintHelper.PrintImageTransaction(UIDoc, UserValues, UserImagePath, IsAuto);

			return Result.Succeeded;
		}
	}
}
