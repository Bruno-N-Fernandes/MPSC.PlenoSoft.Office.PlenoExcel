using DocumentFormat.OpenXml.Spreadsheet;
using MPSC.PlenoSoft.Office.Planilhas.Controller;
using System;

namespace MPSC.PlenoSoft.Office.Planilhas.Celulas
{
	public class FormulaCell : Cell
	{
		public FormulaCell(Celula celula, String formula)
		{
			DataType = CellValues.Number;
			CellFormula = new CellFormula { CalculateCell = true, Text = formula.Substring(1) };
			CellReference = celula.Referencia;
			StyleIndex = 2;
		}
	}
}