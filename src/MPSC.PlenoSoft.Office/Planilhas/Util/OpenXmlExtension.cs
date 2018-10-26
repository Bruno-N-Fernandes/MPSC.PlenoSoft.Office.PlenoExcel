using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MPSC.PlenoSoft.Office.Planilhas.Util
{
	public static class OpenXmlExtension
	{
		public static Boolean Is(this Cell cell, CellValues cellValues)
		{
			return (cell.DataType != null) && cell.DataType.HasValue && (cell.DataType.Value == cellValues);
		}

		public static SharedStringTable GetSharedStringTable(this WorkbookPart workbookPart)
		{
			var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
			return (sharedStringTablePart == null) ? new SharedStringTable() : sharedStringTablePart.SharedStringTable;
		}

		public static SheetData GetSheetDataBySheet(this WorkbookPart workbookPart, Sheet sheet)
		{
			if ((workbookPart == null) || (sheet == null)) return null;
			var worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
			return (worksheetPart != null) ? worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault() : null;
		}

		public static IEnumerable<Sheet> GetSheets(this WorkbookPart workbookPart, Func<Sheet, Int32, Boolean> filtro = null)
		{
			var sheets = workbookPart.Workbook.Sheets.Elements<Sheet>();
			return (filtro != null) ? sheets.Where(filtro) : sheets;
		}

		public static SheetData GetSheetData(this WorkbookPart workbookPart, Func<Sheet, Int32, Boolean> filtro)
		{
			var sheet = workbookPart.GetSheets(filtro).FirstOrDefault();
			return workbookPart.GetSheetDataBySheet(sheet);
		}

		public static IEnumerable<Row> GetRows(this SheetData sheetData, Func<Row, Int32, Boolean> filtro = null)
		{
			var rows = sheetData.ChildElements.OfType<Row>();
			return (filtro != null) ? rows.Where(filtro) : rows;
		}

		private static IEnumerable<Cell> GetCells(this Row row, Func<Cell, Int32, Boolean> filtro = null)
		{
			var cells = row.ChildElements.OfType<Cell>();
			return (filtro != null) ? cells.Where(filtro) : cells;
		}

		public static Row GetRow(this SheetData sheetData, Int32 linha)
		{
			var row = sheetData.GetRows((r, i) => r.RowIndex == linha).FirstOrDefault();
			if (row == null)
				sheetData.Append(row = new Row { RowIndex = (UInt32)linha });
			return row;
		}

		public static Cell GetCell(this Row row, String nomeDaColuna, Cell padrao)
		{
			return row.GetCells((c, i) => (c != null)
				&& (c.CellReference.HasValue)
				&& (!string.IsNullOrWhiteSpace(c.CellReference.Value))
				&& c.CellReference.Value.ToUpper().StartsWith(nomeDaColuna.ToUpper())
			)
			.FirstOrDefault() ?? padrao;
		}
	}
}