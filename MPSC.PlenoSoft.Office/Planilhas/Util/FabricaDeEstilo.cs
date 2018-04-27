using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace MPSC.PlenoSoft.Office.Planilhas.Util
{
	public enum Style : uint
	{
		Geral = 0,
		Long = 1,
		Double = 4,
		Date = 14,
		Header = 999
	}

	public class FabricaDeEstilo
	{
		private FabricaDeEstilo() { }

		internal static Stylesheet Criar(Dictionary<Style, UInt32Value> styles, Stylesheet stylesheet = null)
		{
			var fabricaDeEstilo = new FabricaDeEstilo();
			if (stylesheet == null)
			{
				stylesheet = new Stylesheet();
				fabricaDeEstilo.CreateChildren(stylesheet);
				fabricaDeEstilo.CreateDefaults(stylesheet);
			}
			fabricaDeEstilo.CreateStyles(stylesheet, styles);
			return stylesheet;
		}

		private void CreateChildren(Stylesheet stylesheet)
		{
			stylesheet.Borders = new Borders { Count = 0 };
			stylesheet.CellFormats = new CellFormats { Count = 0 };
			stylesheet.CellStyleFormats = new CellStyleFormats { Count = 0 };
			stylesheet.CellStyles = new CellStyles { Count = 0 };
			stylesheet.Colors = new Colors();
			stylesheet.DifferentialFormats = new DifferentialFormats { Count = 0 };
			stylesheet.Fills = new Fills { Count = 0 };
			stylesheet.Fonts = new Fonts { Count = 0 };
			stylesheet.NumberingFormats = new NumberingFormats { Count = 0 };
			stylesheet.StylesheetExtensionList = new StylesheetExtensionList();
			stylesheet.TableStyles = new TableStyles { Count = 0 };
		}

		private void CreateDefaults(Stylesheet stylesheet)
		{
			New(stylesheet, CreateBorder(true, true, true, true, true));
			New(stylesheet, CreateCellStyleFormat(0, 0, 0));
			New(stylesheet, CreateCellStyle("Normal", 0, 0));
			New(stylesheet, CreateFill(PatternValues.None));
			New(stylesheet, CreateFont("Calibri", 11D, null, null, null, null));
		}

		private void CreateStyles(Stylesheet stylesheet, Dictionary<Style, UInt32Value> styles)
		{
			var fontId = New(stylesheet, CreateFont("Calibri", 11, true, null, null, null));
			styles[Style.Geral] = New(stylesheet, CreateCellFormat((uint)Style.Geral, 0, 0));
			styles[Style.Long] = New(stylesheet, CreateCellFormat((uint)Style.Long, 0, 0));
			styles[Style.Double] = New(stylesheet, CreateCellFormat((uint)Style.Double, 0, 0));
			styles[Style.Date] = New(stylesheet, CreateCellFormat((uint)Style.Date, 0, 0));
			styles[Style.Header] = New(stylesheet, CreateCellFormat((uint)Style.Header, fontId, 0));
		}

		private Border CreateBorder(Boolean? top, Boolean? bottom, Boolean? left, Boolean? right, Boolean? diagonal)
		{
			var border = new Border();

			if (left.HasValue && left.Value)
				border.Append(new LeftBorder());

			if (right.HasValue && right.Value)
				border.Append(new RightBorder());

			if (top.HasValue && top.Value)
				border.Append(new TopBorder());

			if (bottom.HasValue && bottom.Value)
				border.Append(new BottomBorder());

			if (diagonal.HasValue && diagonal.Value)
				border.Append(new DiagonalBorder());

			return border;
		}

		private CellFormat CreateCellFormat(UInt32? numberFormatId, UInt32Value fontId, UInt32? fillId)
		{
			var cellFormat = new CellFormat();

			if (fontId != null)
				cellFormat.FontId = fontId;

			if (fillId.HasValue)
				cellFormat.FillId = fillId.Value;

			if (numberFormatId.HasValue)
			{
				cellFormat.NumberFormatId = numberFormatId.Value;
				cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			}

			return cellFormat;
		}

		private CellStyleFormat CreateCellStyleFormat(UInt32? numberFormatId, UInt32Value fontId, UInt32? fillId)
		{
			var cellStyleFormat = new CellStyleFormat();

			if (fontId != null)
				cellStyleFormat.FontId = fontId;

			if (fillId.HasValue)
				cellStyleFormat.FillId = fillId.Value;

			if (numberFormatId.HasValue)
			{
				cellStyleFormat.NumberFormatId = numberFormatId.Value;
				cellStyleFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
			}

			return cellStyleFormat;
		}

		private CellStyle CreateCellStyle(String name, UInt32? formatId, UInt32? builtinId)
		{
			return new CellStyle
			{
				Name = name,
				FormatId = formatId,
				BuiltinId = builtinId
			};
		}

		private Fill CreateFill(PatternValues? patternValues)
		{
			var fill = new Fill();

			if (patternValues.HasValue)
				fill.PatternFill = new PatternFill { PatternType = patternValues.Value };

			return fill;
		}

		private Font CreateFont(String fontName, Double? fontSize, Boolean? isBold, UInt32? color, Int32? fontFamilyNumbering, FontSchemeValues? fontSchemeValues)
		{
			var font = new Font();

			if (!String.IsNullOrWhiteSpace(fontName))
				font.Append(new FontName() { Val = fontName });

			if (fontSize.HasValue)
				font.Append(new FontSize() { Val = fontSize.Value });

			if (isBold.HasValue && isBold.Value)
				font.Append(new Bold());

			if (color.HasValue)
				font.Append(new Color { Theme = color.Value });

			if (fontFamilyNumbering.HasValue)
				font.Append(new FontFamilyNumbering { Val = fontFamilyNumbering.Value });

			if (fontSchemeValues.HasValue)
				font.Append(new FontScheme { Val = fontSchemeValues.Value });

			return font;
		}

		#region // "News"
		private UInt32Value New(Stylesheet stylesheet, Border border)
		{
			stylesheet.Borders.Append(border);
			return stylesheet.Borders.Count++;
		}
		private UInt32Value New(Stylesheet stylesheet, CellFormat cellFormat)
		{
			stylesheet.CellFormats.Append(cellFormat);
			return stylesheet.CellFormats.Count++;
		}
		private UInt32Value New(Stylesheet stylesheet, CellStyleFormat cellStyleFormat)
		{
			stylesheet.CellStyleFormats.Append(cellStyleFormat);
			return stylesheet.CellStyleFormats.Count++;
		}
		private UInt32Value New(Stylesheet stylesheet, CellStyle cellStyle)
		{
			stylesheet.CellStyles.Append(cellStyle);
			return stylesheet.CellStyles.Count++;
		}
		private UInt32Value New(Stylesheet stylesheet, Color color)
		{
			stylesheet.Colors.Append(color);
			return new UInt32Value((uint)stylesheet.Colors.ChildElements.Count);
		}
		private UInt32Value New(Stylesheet stylesheet, DifferentialFormat differentialFormat)
		{
			stylesheet.DifferentialFormats.Append(differentialFormat);
			return stylesheet.DifferentialFormats.Count++;
		}
		private UInt32Value New(Stylesheet stylesheet, Fill fill)
		{
			stylesheet.Fills.Append(fill);
			return stylesheet.Fills.Count++;
		}
		private UInt32Value New(Stylesheet stylesheet, Font font)
		{
			stylesheet.Fonts.Append(font);
			return stylesheet.Fonts.Count++;
		}
		private UInt32Value New(Stylesheet stylesheet, NumberingFormat numberingFormat)
		{
			stylesheet.NumberingFormats.Append(numberingFormat);
			return stylesheet.NumberingFormats.Count++;
		}
		private UInt32Value New(Stylesheet stylesheet, StylesheetExtension stylesheetExtension)
		{
			stylesheet.StylesheetExtensionList.Append(stylesheetExtension);
			return new UInt32Value((uint)stylesheet.StylesheetExtensionList.ChildElements.Count);
		}
		private UInt32Value New(Stylesheet stylesheet, TableStyle tableStyle)
		{
			stylesheet.TableStyles.Append(tableStyle);
			return stylesheet.TableStyles.Count++;
		}
		#endregion // "News"

		private class CellStyleFormat : CellFormat { }
	}
}