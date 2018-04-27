using DocumentFormat.OpenXml.Spreadsheet;
using MPSC.PlenoSoft.Office.Planilhas.Controller;
using System;

namespace MPSC.PlenoSoft.Office.Planilhas.Celulas
{
	public class LogicalCell : Cell
	{
		private static String _valor0;
		private static String _valor1;

		private LogicalCell(Celula celula, String valor)
		{
			DataType = CellValues.InlineString;
			CellReference = celula.Referencia;
			InlineString = new InlineString { Text = new Text { Text = valor } };
		}

		private LogicalCell(Celula celula, Boolean? valor)
		{
			DataType = CellValues.Boolean;
			CellReference = celula.Referencia;
			CellValue = new CellValue(ObterValorLogico(valor, "0", "1"));
		}

		public static void Configurar(String descricaoParaValoresFalsos, String descricaoParaValoresVerdadeiros)
		{
			_valor0 = descricaoParaValoresFalsos;
			_valor1 = descricaoParaValoresVerdadeiros;
		}

		private static String ObterValorLogico(Boolean? valor, String valor0, String valor1)
		{
			return valor.HasValue ? (valor.Value ? valor1 : valor0) : String.Empty;
		}

		public static Cell Create(Celula celula, Boolean? valor)
		{
			if (String.IsNullOrWhiteSpace(_valor0) || String.IsNullOrWhiteSpace(_valor1))
				return new LogicalCell(celula, valor);

			return new LogicalCell(celula, ObterValorLogico(valor, _valor0, _valor1));
		}
	}
}