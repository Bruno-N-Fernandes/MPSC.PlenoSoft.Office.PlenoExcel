using System;
using System.Xml;
using System.Xml.Serialization;

namespace MPSC.PlenoSoft.Office.ZipXml
{
	[Serializable, XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
	[XmlRoot("sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
	public class SharedStrings
	{
		[XmlAttribute]
		public String uniqueCount;

		[XmlAttribute]
		public String count;

		[XmlElement("si")]
		public SharedString[] si;

		public SharedStrings() { }
	}

	public class SharedString
	{
		public String t;
	}
}