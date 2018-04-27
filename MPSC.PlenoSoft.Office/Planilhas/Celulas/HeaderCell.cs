using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using MPSC.PlenoSoft.Office.Planilhas.Controller;
using System;
using Cor = System.Drawing.Color;
using Translator = System.Drawing.ColorTranslator;

namespace MPSC.PlenoSoft.Office.Planilhas.Celulas
{
	public class HeaderCell : TextCell
	{
		public HeaderCell(Celula celula, String text, Stylesheet styles, Cor fillColor, Double? fontSize, Boolean isBold)
			: base(celula, text)
		{
			UInt32Value fontId = CreateFont(styles, "", fontSize, isBold, Cor.Black);
			UInt32Value fillId = CreateFill(styles, fillColor);
			UInt32Value formatId = CreateCellFormat(styles, fontId, fillId, 0);
			this.StyleIndex = formatId;
		}

		private static UInt32Value CreateCellFormat(Stylesheet styleSheet, UInt32Value fontIndex, UInt32Value fillIndex, UInt32Value numberFormatId)
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
			return styleSheet.CellFormats.Count++;
		}

		private static UInt32Value CreateFill(Stylesheet styleSheet, Cor cor)
		{
			PatternFill patternFill = new PatternFill(new ForegroundColor { Rgb = ObterCorRGB(cor) })
			{
				PatternType = ((cor == System.Drawing.Color.White) ? PatternValues.None : PatternValues.LightDown)
			};

			styleSheet.Fills.Append(new Fill(patternFill));
			return styleSheet.Fills.Count++;
		}


		private static UInt32Value CreateFont(Stylesheet styleSheet, string fontName, double? fontSize, bool isBold, Cor cor)
		{
			Font font = new Font();

			if (!string.IsNullOrEmpty(fontName))
				font.Append(new FontName() { Val = fontName });

			if (fontSize.HasValue)
				font.Append(new FontSize() { Val = fontSize.Value });

			if (isBold == true)
				font.Append(new Bold());

			font.Append(new Color() { Rgb = ObterCorRGB(cor) });

			styleSheet.Fonts.Append(font);
			return styleSheet.Fonts.Count++;
		}

		private static HexBinaryValue ObterCorRGB(Cor cor)
		{
			return new HexBinaryValue { Value = Translator.ToHtml(Cor.FromArgb(cor.A, cor.R, cor.G, cor.B)).Replace("#", String.Empty) };
		}
	}
}