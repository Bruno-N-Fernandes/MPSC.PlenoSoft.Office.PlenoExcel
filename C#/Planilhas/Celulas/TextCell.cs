using DocumentFormat.OpenXml.Spreadsheet;
using MPSC.PlenoSoft.Office.Planilhas.Controller;
using System;

namespace MPSC.PlenoSoft.Office.Planilhas.Celulas
{
	public class TextCell : Cell
	{
		public TextCell(Celula celula, String texto)
		{
			DataType = CellValues.InlineString;
			CellReference = celula.Referencia;
			InlineString = new InlineString { Text = new Text { Text = texto ?? String.Empty } };
		}

		public TextCell(Celula celula, Char chr) : this(celula, chr.ToString()) { }
	}
}