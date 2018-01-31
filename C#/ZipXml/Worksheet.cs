using System;
using System.Xml.Serialization;

namespace MPSC.PlenoSoft.Office.ZipXml
{
	[Serializable, XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
	[XmlRoot("worksheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
	public class Worksheet
	{
		[XmlArray("sheetData"), XmlArrayItem("row")]
		public Row[] Rows;

		[XmlIgnore]
		public Int32 NumberOfColumns; // Total number of columns in this worksheet

		public static Int32 MaxColumnIndex = 0; // Temporary variable for import

		public Worksheet() { }

		public void ExpandRows()
		{
			foreach (var row in Rows)
				row.ExpandCells(NumberOfColumns);
		}
	}
}