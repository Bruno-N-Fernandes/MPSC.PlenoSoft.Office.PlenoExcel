using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using MPSC.PlenoSoft.Office.Planilhas.Integracao;
using MPSC.PlenoSoft.Office.Planilhas.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace MPSC.PlenoSoft.Office.Planilhas.Controller
{
	public class Planilha
	{
		private static readonly Cell NullCell = new Cell();
		private readonly PlenoExcel _plenoExcel;
		private readonly Dictionary<Style, UInt32Value> _styles;
		internal readonly SheetData _sheetData;

		public Planilha(PlenoExcel plenoExcel, Dictionary<Style, UInt32Value> styles, SheetData sheetData)
		{
			_plenoExcel = plenoExcel;
			_styles = styles;
			_sheetData = sheetData;
		}

		#region // "Leitura"

		public Tipo Obter<Tipo>(String referencia, Tipo valorPadrao) where Tipo : struct
		{
			var celula = Celula.From(referencia);
			return Obter(celula, valorPadrao);
		}

		public Tipo Obter<Tipo>(String coluna, Int32 linha, Tipo valorPadrao) where Tipo : struct
		{
			var celula = Celula.From(coluna, linha);
			return Obter(celula, valorPadrao);
		}

		public Tipo Obter<Tipo>(Int32 coluna, Int32 linha, Tipo valorPadrao) where Tipo : struct
		{
			var celula = Celula.From(coluna, linha);
			return Obter(celula, valorPadrao);
		}

		private Tipo Obter<Tipo>(Celula celula, Tipo valorPadrao) where Tipo : struct
		{
			var valor = Ler(celula, null);
			return String.IsNullOrWhiteSpace(valor) ? valorPadrao : (Tipo)Converter(valor, typeof(Tipo));
		}


		public String Ler(String referencia, String valorPadrao = null)
		{
			var celula = Celula.From(referencia);
			return Ler(celula, valorPadrao);
		}

		public String Ler(String coluna, Int32 linha, String valorPadrao = null)
		{
			var celula = Celula.From(coluna, linha);
			return Ler(celula, valorPadrao);
		}

		public String Ler(Int32 coluna, Int32 linha, String valorPadrao = null)
		{
			var celula = Celula.From(coluna, linha);
			return Ler(celula, valorPadrao);
		}

		private String Ler(Celula celula, String valorPadrao)
		{
			var row = _sheetData.GetRow(celula.Linha);
			var cell = row.GetCell(celula.Coluna, NullCell);
			return GetValue(cell, valorPadrao);
		}

		public String GetValue(Cell cell, String valorPadrao)
		{
			String retorno = null;
			if (cell != null)
			{
				if (cell.Is(CellValues.InlineString))
					retorno = cell.InnerText;
				else if (cell.CellValue != null)
				{
					retorno = cell.CellValue.Text;
					if (cell.Is(CellValues.SharedString))
					{
						int ssid = int.Parse(retorno);
						retorno = _plenoExcel._sharedStringTable.ChildElements[ssid].InnerText;
					}
					else if (cell.Is(CellValues.Boolean))
						retorno = (!(retorno ?? cell.InnerText).Equals("0")).ToString();
				}
				else
					retorno = cell.InnerText;
			}

			return retorno ?? valorPadrao;
		}

		#endregion // "Leitura"


		#region // "Escrita"

		public void Escrever(String referencia, Object valor, Style? style = null)
		{
			var celula = Celula.From(referencia);
			Escrever(celula, valor, style);
		}

		public void Escrever(String coluna, Int32 linha, Object valor, Style? style = null)
		{
			var celula = Celula.From(coluna, linha);
			Escrever(celula, valor, style);
		}

		public void Escrever(Int32 coluna, Int32 linha, Object valor, Style? style = null)
		{
			var celula = Celula.From(coluna, linha);
			Escrever(celula, valor, style);
		}

		public void Escrever(Celula celula, Object valor, Style? style = null)
		{
			var row = _sheetData.GetRow(celula.Linha);
			var cell = row.GetCell(celula.Coluna, null);
			if (cell != null)
				row.RemoveChild(cell);

			var estilo = style.HasValue ? _styles[style.Value] : null;
			cell = Celula.Criar(celula, valor, estilo);
			row.AppendChild(cell);
		}

		#endregion // "Escrita"


		#region // "Exportação"

		public void AdicionarDados<T>(IEnumerable<T> lista, IEnumerable<PlenoMapa> mapeamento = null)
		{
			AdicionarDados(lista, typeof(T), mapeamento);
		}

		public void AdicionarDados(IEnumerable lista, Type tipo, IEnumerable<PlenoMapa> mapeamento = null)
		{
			var campos = AdicionarCabecalho(tipo, mapeamento);
			AdicionarLinhas(lista, campos);
		}

		private Cabecalho[] AdicionarCabecalho(Type tipo, IEnumerable<PlenoMapa> mapeamento)
		{
			var campos = Cabecalho.ObterCabecalhos(tipo, mapeamento);

			if (campos.Any())
			{
				var header = EscreverUmaLinha(campos, _styles[Style.Header], campo => campo.Mapeamento.Titulo);

				DefinirTamanhoColunas(campos.OrderBy(c => c.Mapeamento.Posicao).Select(c => c.Mapeamento.Largura).ToArray());
			}

			return campos;
		}

		private void AdicionarLinhas(IEnumerable lista, IEnumerable<Cabecalho> campos)
		{
			if (campos.Any())
			{
				foreach (var item in lista)
				{
					var newRow = EscreverUmaLinha(campos, null, campo => campo.PropertyInfo.GetValue(item, null));

					_plenoExcel.Flush(newRow.RowIndex, this);
				}
				Flush();
			}
		}

		private Row EscreverUmaLinha(IEnumerable<Cabecalho> campos, uint? style, Func<Cabecalho, Object> obterValor)
		{
			var linha = _sheetData.ChildElements.Count() + 1;
			var newRow = _sheetData.GetRow(linha);

			foreach (var campo in campos)
			{
				var value = obterValor(campo);
				var cell = Celula.Criar(campo.Mapeamento.Posicao, linha, value, style);
				newRow.AppendChild(cell);
			}

			return newRow;
		}

		#endregion // "Exportação"


		#region // "LayOut"
		public void Flush()
		{
			_plenoExcel.Flush(0, this);
		}

		public void DefinirTamanhoColunas(params Double[] tamanhos)
		{
			var colunas = _sheetData.Parent.GetFirstChild<Columns>();
			colunas.RemoveAllChildren();
			Coluna.Dimensionar(colunas, tamanhos);
		}

		private static object Converter(string valor, Type tipo)
		{
			object retorno = null;

			if (tipo == typeof(Decimal))
				retorno = Math.Round(Convert.ToDecimal(valor.Replace(".", ",")), 2);

			return retorno;
		}

		#endregion // "LayOut"
	}
}