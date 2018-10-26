using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using MPSC.PlenoSoft.Office.Planilhas.Celulas;
using System;
using System.Text.RegularExpressions;

namespace MPSC.PlenoSoft.Office.Planilhas.Controller
{
	public struct Celula
	{
		public readonly String Coluna;
		public readonly Int32 Linha;
		public readonly Int32 NumeroColuna;
		public readonly String Referencia;

		public Celula(String coluna, Int32 linha)
		{
			Coluna = coluna;
			Linha = linha;
			NumeroColuna = Util.Coluna.ObterIndicePor(Coluna);
			Referencia = Coluna + Linha.ToString();
		}

		public static Celula From(String referencia)
		{
			var coluna = Regex.Replace(referencia, "[0-9]", String.Empty);
			var linha = Convert.ToInt32(Regex.Replace(referencia, "[^0-9]", String.Empty));
			return new Celula(coluna, linha);
		}

		public static Celula From(String coluna, Int32 linha)
		{
			return new Celula(coluna, linha);
		}

		public static Celula From(Int32 coluna, Int32 linha)
		{
			return new Celula(Util.Coluna.ObterNomePor(coluna), linha);
		}

		public static Cell Criar(String referencia, Object obj, UInt32? style)
		{
			return Criar(From(referencia), obj, style);
		}

		public static Cell Criar(String coluna, Int32 linha, Object obj, UInt32? style)
		{
			return Criar(From(coluna, linha), obj, style);
		}

		public static Cell Criar(Int32 coluna, Int32 linha, Object obj, UInt32? style)
		{
			return Criar(From(coluna, linha), obj, style);
		}

		public static Cell Criar(Celula celula, Object obj, UInt32? style)
		{
			Cell c;
			if (obj == null)
				c = new Cell { CellReference = celula.Referencia };

			else if (obj.GetType() == typeof(Boolean))
				c = LogicalCell.Create(celula, Convert.ToBoolean(obj));

			else if (obj.GetType() == typeof(DateTime))
				c = new DateCell(celula, Convert.ToDateTime(obj));

			else if (obj.GetType() == typeof(Decimal))
				c = new NumberCell(celula, Convert.ToDecimal(obj));

			else if (obj.GetType() == typeof(Double))
				c = new NumberCell(celula, Convert.ToDouble(obj));

			else if (obj.GetType() == typeof(Single))
				c = new NumberCell(celula, Convert.ToSingle(obj));

			else if (obj.GetType() == typeof(Int64))
				c = new NumberCell(celula, Convert.ToInt64(obj));

			else if (obj.GetType() == typeof(Int32))
				c = new NumberCell(celula, Convert.ToInt32(obj));

			else if (obj.GetType() == typeof(Int16))
				c = new NumberCell(celula, Convert.ToInt16(obj));

			else if (obj.GetType() == typeof(Byte))
				c = new NumberCell(celula, Convert.ToByte(obj));

			else if (obj.GetType() == typeof(UInt64))
				c = new NumberCell(celula, Convert.ToUInt64(obj));

			else if (obj.GetType() == typeof(UInt32))
				c = new NumberCell(celula, Convert.ToUInt32(obj));

			else if (obj.GetType() == typeof(UInt16))
				c = new NumberCell(celula, Convert.ToUInt16(obj));

			else if (obj.GetType() == typeof(SByte))
				c = new NumberCell(celula, Convert.ToSByte(obj));

			else if (obj.GetType() == typeof(Char))
				c = new TextCell(celula, Convert.ToChar(obj));

			else if (obj.GetType() == typeof(String))
				c = Convert.ToString(obj).StartsWith("=") ? new FormulaCell(celula, Convert.ToString(obj)) : new TextCell(celula, Convert.ToString(obj)) as Cell;

			else
				c = new TextCell(celula, Convert.ToString(obj));

			c.StyleIndex = (style.HasValue ? (UInt32Value)style.Value : (UInt32Value)null) ?? c.StyleIndex;

			return c;
		}
	}
}