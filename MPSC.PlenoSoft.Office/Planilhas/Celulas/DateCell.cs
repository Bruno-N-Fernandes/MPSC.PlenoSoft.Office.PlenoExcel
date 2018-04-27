using DocumentFormat.OpenXml.Spreadsheet;
using MPSC.PlenoSoft.Office.Planilhas.Controller;
using System;

namespace MPSC.PlenoSoft.Office.Planilhas.Celulas
{
	public class DateCell : Cell
	{
		public DateCell(Celula celula, DateTime? dateTime)
		{
			DataType = CellValues.Date;
			CellReference = celula.Referencia;
			StyleIndex = 3;
			CellValue = new CellValue
			{
				Text = (dateTime.HasValue && (dateTime.Value != default(DateTime)))
					? dateTime.Value.ToString("yyyy-MM-dd")
					: String.Empty
			};
		}
	}
}