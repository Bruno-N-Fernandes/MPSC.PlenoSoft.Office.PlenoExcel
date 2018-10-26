using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Text.RegularExpressions;

namespace MPSC.PlenoSoft.Office.Planilhas.Util
{
	public static class Coluna
	{
		public static Column CreateColumn(Int32 indice, Double tamanho, Boolean ajustarTamanho = true)
		{
			return new Column
			{
				BestFit = ajustarTamanho,
				Min = (UInt32)indice,
				Max = (UInt32)indice,
				Width = tamanho,
				CustomWidth = true,
			};
		}

		public static Columns Dimensionar(params Double[] tamanhos)
		{
			var columns = new Columns(CreateColumn(1, 10));
			return Dimensionar(columns, tamanhos);
		}

		public static Columns Dimensionar(Columns columns, params Double[] tamanhos)
		{
			var index = 0;
			foreach (var tamanho in tamanhos)
				columns.Append(CreateColumn(++index, tamanho));
			return columns;
		}

		public static String ObterNomePor(Int32 coluna)
		{
			var col = Convert.ToByte(coluna);
			if (col < 26)
				return String.Format("{0}", (Char)(64 + col));

			return String.Format("{0}{1}", (Char)(64 + (col / 26)), (Char)(64 + (col % 26)));
		}

		public static Int32 ObterIndicePor(String coluna)
		{
			coluna = Regex.Replace(coluna, "[0-9]", String.Empty);

			if (coluna.Length > 1)
				return ((coluna[0] - 64) * 26) + (coluna[1] - 64);
			else
				return coluna[0] - 64;
		}
	}
}