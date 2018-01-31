using System;
using System.Xml.Serialization;

namespace MPSC.PlenoSoft.Office.ZipXml
{
	public class Row
	{
		[XmlElement("c")]
		public Cell[] FilledCells;

		[XmlIgnore]
		public Cell[] Cells;

		public void ExpandCells(Int32 NumberOfColumns)
		{
			Cells = new Cell[NumberOfColumns];
			foreach (var cell in FilledCells)
				Cells[cell.ColumnIndex] = cell;
			FilledCells = null;
		}
	}
}