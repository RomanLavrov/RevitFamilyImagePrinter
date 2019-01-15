using Autodesk.Revit.DB;
using System;

namespace RevitFamilyImagePrinter
{
	public enum ImageAspectRatio
	{
		Ratio_16to9,
		Ratio_4to3,
		Ratio_1to1
	}

	[Serializable]
	public class UserImageValues
	{
		public int UserScale { get; set; }
		public int UserImageHeight { get; set; }
		public ImageResolution UserImageResolution { get; set; }
		public string UserExtension { get; set; }
		public double UserZoomValue { get; set; }
		public ViewDetailLevel UserDetailLevel { get; set; }
		public ImageAspectRatio UserAspectRatio { get; set; }
	}
}
