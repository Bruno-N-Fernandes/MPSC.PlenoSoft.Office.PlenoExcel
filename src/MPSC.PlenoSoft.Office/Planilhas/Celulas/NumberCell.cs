using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using MPSC.PlenoSoft.Office.Planilhas.Controller;
using System;
using System.Globalization;

namespace MPSC.PlenoSoft.Office.Planilhas.Celulas
{
	public class NumberCell : Cell
	{
		private static readonly CultureInfo en_US = CultureInfo.GetCultureInfo("en-US");

		private static String Format(Object numero)
		{
			if ((numero == null) || ((numero is Double) && Double.IsNaN((Double)numero)))
				return String.Empty;

			return String.Format(en_US, "{0:#0.0000000000}", numero);
		}

		private NumberCell(Celula celula, String numero, UInt32Value styleIndex)
		{
			DataType = CellValues.Number;
			CellReference = celula.Referencia;
			StyleIndex = styleIndex;
			CellValue = new CellValue(numero);
		}

		private NumberCell(Celula celula, Object numero, UInt32Value styleIndex)
			: this(celula, Format(numero), styleIndex) { }

		public NumberCell(Celula celula, Int64 numero) : this(celula, numero.ToString(), 1) { }
		public NumberCell(Celula celula, Int32 numero) : this(celula, numero.ToString(), 1) { }
		public NumberCell(Celula celula, Int16 numero) : this(celula, numero.ToString(), 1) { }
		public NumberCell(Celula celula, Byte numero) : this(celula, numero.ToString(), 1) { }

		public NumberCell(Celula celula, UInt64 numero) : this(celula, numero.ToString(), 1) { }
		public NumberCell(Celula celula, UInt32 numero) : this(celula, numero.ToString(), 1) { }
		public NumberCell(Celula celula, UInt16 numero) : this(celula, numero.ToString(), 1) { }
		public NumberCell(Celula celula, SByte numero) : this(celula, numero.ToString(), 1) { }

		public NumberCell(Celula celula, Decimal numero) : this(celula, numero, 2) { }
		public NumberCell(Celula celula, Double numero) : this(celula, numero, 2) { }
		public NumberCell(Celula celula, Single numero) : this(celula, numero, 2) { }
	}
}