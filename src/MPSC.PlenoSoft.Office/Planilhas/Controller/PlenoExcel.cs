using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MPSC.PlenoSoft.Office.Planilhas.Integracao;
using MPSC.PlenoSoft.Office.Planilhas.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace MPSC.PlenoSoft.Office.Planilhas.Controller
{
	[Flags]
	public enum Modo
	{
		Leitura = 1,
		Escrita = 2,
		CriarSeNaoExistir = 4,
		SalvarAutomaticamente = 8,
		ApagarSeExistir = 16,
		FazerBackupAntes = 32,
		Padrao = Leitura | Escrita | CriarSeNaoExistir | SalvarAutomaticamente,
		Seguro = Padrao | FazerBackupAntes,
		SempreCriaNovo = Leitura | Escrita | ApagarSeExistir | CriarSeNaoExistir | SalvarAutomaticamente
	}

	public class PlenoExcel
	{
		private const UInt32 politicaAutoFlush = 1000;
		private readonly FileInfo _arquivoExcel;
		private readonly Modo _modo;
		private readonly SpreadsheetDocument _spreadsheetDocument;
		internal readonly SharedStringTable _sharedStringTable;
		private readonly Dictionary<String, Planilha> _planilhas;
		private readonly Dictionary<Style, UInt32Value> _styles;

		private FileVersion NewFileVersion { get { return new FileVersion { ApplicationName = "PlenoExcel", BuildVersion = Assembly.GetExecutingAssembly().ObterVersaoDoAssembly().ToString() }; } }
		private CalculationProperties NewCalculationProperties { get { return new CalculationProperties { ForceFullCalculation = true, FullCalculationOnLoad = true, CalculationOnSave = true }; } }

		protected PlenoExcel(FileInfo arquivoExcel)
		{
			_arquivoExcel = arquivoExcel;
			_planilhas = new Dictionary<String, Planilha>();
			_styles = new Dictionary<Style, UInt32Value>();
		}

		public PlenoExcel(FileInfo arquivoExcel, Modo modo = Modo.Padrao) : this(arquivoExcel)
		{
			_modo = modo;
			_spreadsheetDocument = Validar(arquivoExcel, modo);
			if (_spreadsheetDocument.WorkbookPart == null)
			{
				var _workbookPart = _spreadsheetDocument.AddWorkbookPart();
				_workbookPart.Workbook = new Workbook(NewFileVersion, new Sheets()) { CalculationProperties = NewCalculationProperties };
				_workbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = FabricaDeEstilo.Criar(_styles);
			}
			else if (modo.Is(Modo.Escrita))
			{
				var _workbookPart = _spreadsheetDocument.WorkbookPart;
				_workbookPart.Workbook.FileVersion = NewFileVersion;
				_workbookPart.Workbook.CalculationProperties = NewCalculationProperties;
				FabricaDeEstilo.Criar(_styles, _workbookPart.WorkbookStylesPart.Stylesheet);
			}

			_sharedStringTable = _spreadsheetDocument.WorkbookPart.GetSharedStringTable();
		}

		private SpreadsheetDocument Validar(FileInfo arquivoExcel, Modo modo)
		{
			var editavel = modo.Is(Modo.Escrita) || modo.Is(Modo.CriarSeNaoExistir);
			var autoSave = modo.Is(Modo.SalvarAutomaticamente);
			var backup = modo.Is(Modo.FazerBackupAntes);
			var apagar = modo.Is(Modo.ApagarSeExistir);
			var criar = modo.Is(Modo.CriarSeNaoExistir);

			if (!Directory.Exists(arquivoExcel.Directory.FullName))
				arquivoExcel.Directory.Create();

			if (backup)
				arquivoExcel.Backup();

			if (apagar && File.Exists(arquivoExcel.FullName))
				arquivoExcel.Delete();

			SpreadsheetDocument retorno = null;
			if (File.Exists(arquivoExcel.FullName))
				retorno = SpreadsheetDocument.Open(arquivoExcel.FullName, editavel, new OpenSettings() { AutoSave = autoSave });
			else if (criar)
				retorno = SpreadsheetDocument.Create(arquivoExcel.FullName, SpreadsheetDocumentType.Workbook, autoSave);

			if (retorno == null)
				throw new Exception("Configurações de [Modo] inválidas.");

			return retorno;
		}

		public Planilha this[String nomeDaPlanilha]
		{
			get
			{
				return _planilhas.TryGetValue(nomeDaPlanilha.ToUpper(), out Planilha planilha) ? planilha : ProvidenciarPlanilha(nomeDaPlanilha);
			}
		}

		public Planilha this[Int32 indiceDaPlanilha]
		{
			get
			{
				var planilha = _planilhas.Values.Skip(indiceDaPlanilha).FirstOrDefault();
				return planilha ?? ProvidenciarPlanilha(indiceDaPlanilha);
			}
		}

		private Planilha ProvidenciarPlanilha(Int32 index)
		{
			var nomeDaPlanilha = String.Format("Plan#{0}", index + 1);
			var planilha = ObterPlanilha((s, i) => { if (i == index) nomeDaPlanilha = s.Name; return (i == index); }) ?? ProvidenciarPlanilha(nomeDaPlanilha);
			return _planilhas[nomeDaPlanilha.ToUpper()] = planilha;
		}

		private Planilha ProvidenciarPlanilha(String nomeDaPlanilha)
		{
			var planilha = ObterPlanilha((s, i) => s.Name == nomeDaPlanilha) ?? CriarPlanilha(nomeDaPlanilha);
			return _planilhas[nomeDaPlanilha.ToUpper()] = planilha;
		}

		private Planilha ObterPlanilha(Func<Sheet, Int32, Boolean> filtro)
		{
			var sheetData = _spreadsheetDocument.WorkbookPart.GetSheetData(filtro);
			return (sheetData != null) ? new Planilha(this, _styles, sheetData) : null;
		}

		private Planilha CriarPlanilha(String nomeDaPlanilha, params Double[] tamanhos)
		{
			var wbP = _spreadsheetDocument.WorkbookPart;
			var sheetData = new SheetData();
			var worksheetPart = wbP.AddNewPart<WorksheetPart>();
			worksheetPart.Worksheet = new Worksheet(Coluna.Dimensionar(tamanhos), sheetData);
			worksheetPart.Worksheet.Save();
			wbP.Workbook.Sheets.Append(new Sheet { Name = nomeDaPlanilha, SheetId = (UInt32)wbP.GetPartsOfType<WorksheetPart>().Count(), Id = wbP.GetIdOfPart(worksheetPart) });
			return new Planilha(this, _styles, sheetData);
		}

		public void Fechar()
		{
			Salvar(_modo.Is(Modo.SalvarAutomaticamente));
			_spreadsheetDocument.Close();
			_spreadsheetDocument.Dispose();
			GC.Collect();
		}

		public void Salvar()
		{
			Salvar(true);
		}

		private void Salvar(Boolean deveSalvar)
		{
			if (_modo.Is(Modo.Escrita) && deveSalvar)
			{
				foreach (var worksheetParts in _spreadsheetDocument.WorkbookPart.WorksheetParts)
					worksheetParts.Worksheet.Save();

				_spreadsheetDocument.WorkbookPart.Workbook.Save();
				_spreadsheetDocument.Save();
			}
		}

		internal void Flush(UInt32 iteracao, Planilha planilha)
		{
			if (_modo.Is(Modo.Escrita) && (iteracao % politicaAutoFlush == 0) && ((iteracao == 0) || _modo.Is(Modo.SalvarAutomaticamente)))
			{
				foreach (var worksheetParts in _spreadsheetDocument.WorkbookPart.WorksheetParts)
				{
					if (worksheetParts.Worksheet.OfType<SheetData>().Any(d => d == planilha._sheetData))
						worksheetParts.Worksheet.Save();
				}
				GC.Collect();
			}
		}
		public static Type GetTypeOfItems<TCollection>(TCollection collection, out IEnumerable iEnumerable)
		{
			iEnumerable = ((collection as IEnumerable) ?? new[] { collection });
			return iEnumerable?.GetType()?.GetGenericArguments()?.FirstOrDefault()
				?? iEnumerable?.OfType<Object>()?.FirstOrDefault()?.GetType()
				?? typeof(TCollection);
		}

		public void Exportar<TClass>(TClass dto, IEnumerable<PlenoMapa> mapeamento = null)
		{
			if (dto is IEnumerable lista)
				ExportarDados(lista, mapeamento, null, false);
			else
			{
				var campos = Cabecalho.ObterCabecalhos(dto.GetType(), mapeamento);
				foreach (var campo in campos)
				{
					var property = campo.PropertyInfo;
					var value = GetValue(dto, property);
					ExportarDados(value, mapeamento, campo.Mapeamento.Titulo, property.PropertyType.IsPrimitive);
				}
			}
			Salvar();
		}

		private void ExportarDados(object value, IEnumerable<PlenoMapa> mapeamento, string planName, bool isPrimitive)
		{
			if (value != null)
			{
				var tipo = GetTypeOfItems(value, out var lista);
				var planilha = this[planName ?? tipo.Name];

				if (lista.OfType<Object>().Any())
				{
					if (isPrimitive)
						planilha.AdicionarDados(lista.OfType<Object>().Select(i => new Interno(i)), typeof(Interno));
					else
						planilha.AdicionarDados(lista, tipo, mapeamento);
				}
			}
		}

		private object GetValue<TClass>(TClass dto, PropertyInfo property)
		{
			try
			{
				return property.GetValue(dto, null);
			}
			catch (Exception)
			{
				var _ = GetTypeOfItems(dto, out var lista);
				return lista;
			}
		}

		public class Interno
		{
			public Object Valor { get; private set; }

			public Interno(Object valor)
			{
				Valor = valor;
			}
		}
	}
}