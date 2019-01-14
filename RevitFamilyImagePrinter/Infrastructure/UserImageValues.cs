using Autodesk.Revit.DB;
using System;

namespace RevitFamilyImagePrinter
{
	[Serializable]
	public class UserImageValues
	{
		public int UserScale { get; set; }
		public int UserImageSize { get; set; }
		public ImageResolution UserImageResolution { get; set; }
		public string UserExtension { get; set; }
		public double UserZoomValue { get; set; }
		public ViewDetailLevel UserDetailLevel { get; set; }
	}
}
