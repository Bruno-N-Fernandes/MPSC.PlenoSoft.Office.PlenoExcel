using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace MPSC.PlenoSoft.Office.Planilhas.Util
{
	public class CustomStylesheet : Stylesheet
	{
		public CustomStylesheet()
		{
			Fonts fts = new Fonts();
			DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
			FontName ftn = new FontName();
			ftn.Val = StringValue.FromString("Calibri");
			FontSize ftsz = new FontSize();
			ftsz.Val = DoubleValue.FromDouble(11);
			ft.FontName = ftn;
			ft.FontSize = ftsz;
			fts.Append(ft);

			ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
			ftn = new FontName();
			ftn.Val = StringValue.FromString("Palatino Linotype");
			ftsz = new FontSize();
			ftsz.Val = DoubleValue.FromDouble(18);
			ft.FontName = ftn;
			ft.FontSize = ftsz;
			fts.Append(ft);

			fts.Count = UInt32Value.FromUInt32((uint)fts.ChildElements.Count);

			Fills fills = new Fills();
			Fill fill;
			PatternFill patternFill;
			fill = new Fill();
			patternFill = new PatternFill();
			patternFill.PatternType = PatternValues.None;
			fill.PatternFill = patternFill;
			fills.Append(fill);

			fill = new Fill();
			patternFill = new PatternFill();
			patternFill.PatternType = PatternValues.Gray125;
			fill.PatternFill = patternFill;
			fills.Append(fill);

			fill = new Fill();
			patternFill = new PatternFill();
			patternFill.PatternType = PatternValues.Solid;
			patternFill.ForegroundColor = new ForegroundColor();
			patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString("00ff9728");
			patternFill.BackgroundColor = new BackgroundColor();
			patternFill.BackgroundColor.Rgb = patternFill.ForegroundColor.Rgb;
			fill.PatternFill = patternFill;
			fills.Append(fill);

			fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);

			Borders borders = new Borders();
			Border border = new Border();
			border.LeftBorder = new LeftBorder();
			border.RightBorder = new RightBorder();
			border.TopBorder = new TopBorder();
			border.BottomBorder = new BottomBorder();
			border.DiagonalBorder = new DiagonalBorder();
			borders.Append(border);

			//Boarder Index 1
			border = new Border();
			border.LeftBorder = new LeftBorder();
			border.LeftBorder.Style = BorderStyleValues.Thin;
			border.RightBorder = new RightBorder();
			border.RightBorder.Style = BorderStyleValues.Thin;
			border.TopBorder = new TopBorder();
			border.TopBorder.Style = BorderStyleValues.Thin;
			border.BottomBorder = new BottomBorder();
			border.BottomBorder.Style = BorderStyleValues.Thin;
			border.DiagonalBorder = new DiagonalBorder();
			borders.Append(border);


			//Boarder Index 2
			border = new Border();
			border.LeftBorder = new LeftBorder();
			border.RightBorder = new RightBorder();
			border.TopBorder = new TopBorder();
			border.TopBorder.Style = BorderStyleValues.Thin;
			border.BottomBorder = new BottomBorder();
			border.BottomBorder.Style = BorderStyleValues.Thin;
			border.DiagonalBorder = new DiagonalBorder();
			borders.Append(border);


			borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);

			CellStyleFormats csfs = new CellStyleFormats();
			CellFormat cf = new CellFormat();
			cf.NumberFormatId = 0;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			csfs.Append(cf);
			csfs.Count = UInt32Value.FromUInt32((uint)csfs.ChildElements.Count);

			uint iExcelIndex = 164;
			NumberingFormats nfs = new NumberingFormats();
			CellFormats cfs = new CellFormats();

			cf = new CellFormat();
			cf.NumberFormatId = 0;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cfs.Append(cf);

			NumberingFormat nfDateTime = new NumberingFormat();
			nfDateTime.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
			nfDateTime.FormatCode = StringValue.FromString("dd/MM/yyyy");
			nfs.Append(nfDateTime);

			NumberingFormat nf4decimal = new NumberingFormat();
			nf4decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
			nf4decimal.FormatCode = StringValue.FromString("#,##0.0000");
			nfs.Append(nf4decimal);

			// #,##0.00 is also Excel style index 4
			NumberingFormat nf2decimal = new NumberingFormat();
			nf2decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
			nf2decimal.FormatCode = StringValue.FromString("#,##0.00");
			nfs.Append(nf2decimal);

			// @ is also Excel style index 49
			NumberingFormat nfForcedText = new NumberingFormat();
			nfForcedText.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
			nfForcedText.FormatCode = StringValue.FromString("@");
			nfs.Append(nfForcedText);

			// index 1
			// Format dd/mm/yyyy
			cf = new CellFormat();
			cf.NumberFormatId = 14;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);

			// index 2
			// Format #,##0.00
			cf = new CellFormat();
			cf.NumberFormatId = 4;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);

			// index 3
			cf = new CellFormat();
			cf.NumberFormatId = nfDateTime.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);

			// index 4
			cf = new CellFormat();
			cf.NumberFormatId = nf4decimal.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);

			// index 5
			cf = new CellFormat();
			cf.NumberFormatId = nf2decimal.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);

			// index 6
			cf = new CellFormat();
			cf.NumberFormatId = nfForcedText.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);

			// index 7
			// Header text
			cf = new CellFormat();
			cf.NumberFormatId = nfForcedText.NumberFormatId;
			cf.FontId = 1;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);

			// index 8
			// column text
			cf = new CellFormat();
			cf.NumberFormatId = nfForcedText.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 1;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);

			// index 9
			// coloured 2 decimal text
			cf = new CellFormat();
			cf.NumberFormatId = nf2decimal.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 2;
			cf.BorderId = 2;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);

			// index 10
			// coloured column text
			cf = new CellFormat();
			cf.NumberFormatId = nfForcedText.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 2;
			cf.BorderId = 2;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			cfs.Append(cf);


			nfs.Count = UInt32Value.FromUInt32((uint)nfs.ChildElements.Count);
			cfs.Count = UInt32Value.FromUInt32((uint)cfs.ChildElements.Count);

			this.Append(nfs);
			this.Append(fts);
			this.Append(fills);
			this.Append(borders);
			this.Append(csfs);
			this.Append(cfs);

			CellStyles css = new CellStyles();
			CellStyle cs = new CellStyle();
			cs.Name = StringValue.FromString("Normal");
			cs.FormatId = 0;
			cs.BuiltinId = 0;
			css.Append(cs);
			css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
			this.Append(css);

			DifferentialFormats dfs = new DifferentialFormats();
			dfs.Count = 0;
			this.Append(dfs);

			TableStyles tss = new TableStyles();
			tss.Count = 0;
			tss.DefaultTableStyle = StringValue.FromString("TableStyleMedium9");
			tss.DefaultPivotStyle = StringValue.FromString("PivotStyleLight16");
			this.Append(tss);
		}


		private Cell CreateIntegerCell(string header, string text, int index)
		{
			Cell c = new Cell();
			c.DataType = CellValues.Number;
			c.CellReference = header + index;

			CellValue v = new CellValue();
			v.Text = text;
			c.AppendChild(v);
			return c;
		}

		private Cell CreateDecimalCell(string header, string text, int index, Stylesheet styles)
		{
			Cell c = new Cell();
			c.DataType = CellValues.Number;
			c.CellReference = header + index;
			UInt32Value fontId = CreateFont(styles, "Arial", 11, false, System.Drawing.Color.Black);
			UInt32Value fillId = CreateFill(styles, System.Drawing.Color.White);
			UInt32Value formatId = CreateCellFormat(styles, fontId, fillId, 171);
			c.StyleIndex = formatId;

			CellValue v = new CellValue();
			v.Text = text;
			c.AppendChild(v);
			return c;
		}

		private Cell CreateFomulaCell(string header, string formula, int index, Stylesheet styles)
		{
			Cell c = new Cell();
			c.DataType = CellValues.Number;
			c.CellReference = header + index;
			UInt32Value fontId = CreateFont(styles, "Arial", 11, false, System.Drawing.Color.Black);
			UInt32Value fillId = CreateFill(styles, System.Drawing.Color.White);
			UInt32Value formatId = CreateCellFormat(styles, fontId, fillId, 171);
			c.StyleIndex = formatId;

			CellFormula f = new CellFormula();
			f.CalculateCell = true;
			f.Text = formula;
			c.Append(f);

			CellValue v = new CellValue();
			c.AppendChild(v);
			return c;
		}


		private Cell CreateDateCell(string header, string text, int index, Stylesheet styles)
		{
			Cell c = new Cell();
			c.DataType = CellValues.Date;
			c.CellReference = header + index;

			UInt32Value fontId = CreateFont(styles, "Arial", 11, false, System.Drawing.Color.Black);
			UInt32Value fillId = CreateFill(styles, System.Drawing.Color.White);
			UInt32Value formatId = CreateCellFormat(styles, fontId, fillId, 14);
			c.StyleIndex = formatId;

			CellValue v = new CellValue();
			v.Text = text;
			c.CellValue = v;


			return c;
		}

		private Cell CreateTextCell(string header, string text, int index)
		{

			//Create a new inline string cell.
			Cell c = new Cell();
			c.DataType = CellValues.InlineString;
			c.CellReference = header + index;

			//Add text to the text cell.
			InlineString inlineString = new InlineString();
			Text t = new Text();
			t.Text = text;
			inlineString.AppendChild(t);
			c.AppendChild(inlineString);
			return c;
		}

		private Cell CreateHeaderCell(string header, string text, int index, Stylesheet styles)
		{
			//Create a new inline string cell.
			Cell c = new Cell();
			c.DataType = CellValues.InlineString;
			c.CellReference = header + index;

			UInt32Value fontId = CreateFont(styles, "Arial", 12, true, System.Drawing.Color.Black);
			UInt32Value fillId = CreateFill(styles, System.Drawing.Color.ForestGreen);
			UInt32Value formatId = CreateCellFormat(styles, fontId, fillId, 0);
			c.StyleIndex = formatId;

			//Add text to the text cell.
			InlineString inlineString = new InlineString();
			Text t = new Text();
			t.Text = text;
			inlineString.AppendChild(t);
			c.AppendChild(inlineString);
			return c;
		}
		private List<string> GetPropertyInfo<T>()
		{

			PropertyInfo[] propertyInfos = typeof(T).GetProperties();
			// write property names
			return propertyInfos.Select(propertyInfo => propertyInfo.Name).ToList();
		}

		private Stylesheet CreateStylesheet()
		{
			var ss = new Stylesheet();

			var fts = new Fonts();
			var ftn = new FontName { Val = "Arial" };
			var ftsz = new FontSize { Val = 11 };
			var ft = new DocumentFormat.OpenXml.Spreadsheet.Font { FontName = ftn, FontSize = ftsz };
			fts.Append(ft);
			fts.Count = (uint)fts.ChildElements.Count;


			var fills = new Fills();
			var fill = new Fill();
			var patternFill = new PatternFill { PatternType = PatternValues.None };
			fill.PatternFill = patternFill;
			fills.Append(fill);

			fill = new Fill();
			patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
			fill.PatternFill = patternFill;
			fills.Append(fill);

			fills.Count = (uint)fills.ChildElements.Count;

			var borders = new Borders();
			var border = new Border
			{
				LeftBorder = new LeftBorder(),
				RightBorder = new RightBorder(),
				TopBorder = new TopBorder(),
				BottomBorder = new BottomBorder(),
				DiagonalBorder = new DiagonalBorder()
			};
			borders.Append(border);
			borders.Count = (uint)borders.ChildElements.Count;

			var csfs = new CellStyleFormats();
			var cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 };
			csfs.Append(cf);
			csfs.Count = (uint)csfs.ChildElements.Count;

			// dd/mm/yyyy is also Excel style index 14

			uint iExcelIndex = 164;
			var nfs = new NumberingFormats();
			var cfs = new CellFormats();

			cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 };
			cfs.Append(cf);

			var nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "dd/mm/yyyy hh:mm:ss" };
			nfs.Append(nf);

			cf = new CellFormat
			{
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);


			iExcelIndex = 165;
			nfs = new NumberingFormats();
			cfs = new CellFormats();

			cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 };
			cfs.Append(cf);

			nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "MMM yyyy" };
			nfs.Append(nf);

			cf = new CellFormat
			{
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);


			iExcelIndex = 170;
			nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "#,##0.0000" };
			nfs.Append(nf);
			cf = new CellFormat
			{
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);

			// #,##0.00 is also Excel style index 4
			iExcelIndex = 171;
			nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "#,##0.00" };
			nfs.Append(nf);
			cf = new CellFormat
			{
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);

			// @ is also Excel style index 49
			iExcelIndex = 172;
			nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "@" };
			nfs.Append(nf);
			cf = new CellFormat
			{
				NumberFormatId = nf.NumberFormatId,
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				ApplyNumberFormat = true
			};
			cfs.Append(cf);

			nfs.Count = (uint)nfs.ChildElements.Count;
			cfs.Count = (uint)cfs.ChildElements.Count;

			ss.Append(nfs);
			ss.Append(fts);
			ss.Append(fills);
			ss.Append(borders);
			ss.Append(csfs);
			ss.Append(cfs);

			var css = new CellStyles();
			var cs = new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 };
			css.Append(cs);
			css.Count = (uint)css.ChildElements.Count;
			ss.Append(css);

			var dfs = new DifferentialFormats { Count = 0 };
			ss.Append(dfs);

			var tss = new TableStyles
			{
				Count = 0,
				DefaultTableStyle = "TableStyleMedium9",
				DefaultPivotStyle = "PivotStyleLight16"
			};
			ss.Append(tss);

			return ss;
		}

		private static UInt32Value CreateCellFormat(
			Stylesheet styleSheet,
			UInt32Value fontIndex,
			UInt32Value fillIndex,
			UInt32Value numberFormatId)
		{
			CellFormat cellFormat = new CellFormat();

			if (fontIndex != null)
				cellFormat.FontId = fontIndex;

			if (fillIndex != null)
				cellFormat.FillId = fillIndex;

			if (numberFormatId != null)
			{
				cellFormat.NumberFormatId = numberFormatId;
				cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			}

			styleSheet.CellFormats.Append(cellFormat);

			UInt32Value result = styleSheet.CellFormats.Count;
			styleSheet.CellFormats.Count++;
			return result;
		}

		private UInt32Value CreateFill(
			Stylesheet styleSheet,
			System.Drawing.Color fillColor)
		{


			PatternFill patternFill =
				new PatternFill(
					new ForegroundColor()
					{
						Rgb = new HexBinaryValue()
						{
							Value =
							System.Drawing.ColorTranslator.ToHtml(
								System.Drawing.Color.FromArgb(
									fillColor.A,
									fillColor.R,
									fillColor.G,
									fillColor.B)).Replace("#", "")
						}
					});

			patternFill.PatternType = fillColor ==
						System.Drawing.Color.White ? PatternValues.None : PatternValues.LightDown;

			Fill fill = new Fill(patternFill);

			styleSheet.Fills.Append(fill);

			UInt32Value result = styleSheet.Fills.Count;
			styleSheet.Fills.Count++;
			return result;
		}

		private UInt32Value CreateFont(
			Stylesheet styleSheet,
			string fontName,
			double? fontSize,
			bool isBold,
			System.Drawing.Color foreColor)
		{

			Font font = new Font();

			if (!string.IsNullOrEmpty(fontName))
			{
				FontName name = new FontName()
				{
					Val = fontName
				};
				font.Append(name);
			}

			if (fontSize.HasValue)
			{
				FontSize size = new FontSize()
				{
					Val = fontSize.Value
				};
				font.Append(size);
			}

			if (isBold == true)
			{
				Bold bold = new Bold();
				font.Append(bold);
			}


			Color color = new Color()
			{
				Rgb = new HexBinaryValue()
				{
					Value =
						System.Drawing.ColorTranslator.ToHtml(
							System.Drawing.Color.FromArgb(
								foreColor.A,
								foreColor.R,
								foreColor.G,
								foreColor.B)).Replace("#", "")
				}
			};
			font.Append(color);

			styleSheet.Fonts.Append(font);
			UInt32Value result = styleSheet.Fonts.Count;
			styleSheet.Fonts.Count++;
			return result;
		}

		private Column CreateColumnData(UInt32 columnIndex, double columnWidth)
		{
			Column column = new Column();
			column.Min = columnIndex;
			column.Max = columnIndex;
			column.Width = columnWidth;
			column.CustomWidth = true;
			return column;
		}

		private int GetExcelSerialDate(DateTime input)
		{
			int nDay = input.Day;
			int nMonth = input.Month;
			int nYear = input.Year;
			// Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
			// leap year, but Excel/Lotus 123 think it is...
			if (nDay == 29 && nMonth == 02 && nYear == 1900)
				return 60;

			// DMY to Modified Julian calculatie with an extra substraction of 2415019.
			long nSerialDate =
					(int)((1461 * (nYear + 4800 + (int)((nMonth - 14) / 12))) / 4) +
					(int)((367 * (nMonth - 2 - 12 * ((nMonth - 14) / 12))) / 12) -
					(int)((3 * ((int)((nYear + 4900 + (int)((nMonth - 14) / 12)) / 100))) / 4) +
					nDay - 2415019 - 32075;

			if (nSerialDate < 60)
			{
				// Because of the 29-02-1900 bug, any serial date 
				// under 60 is one off... Compensate.
				nSerialDate--;
			}

			return (int)nSerialDate;

		}
	}
}