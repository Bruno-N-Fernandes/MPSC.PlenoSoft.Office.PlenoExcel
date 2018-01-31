using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace MPSC.PlenoSoft.Office.ZipXml
{
	public class Cell
	{
		private String _value = "";

		[XmlAttribute("r")]
		public String CellReference
		{
			get { return ColumnIndex.ToString(); }
			set { setCellReference(value); }
		}

		[XmlAttribute("t")]
		public String tType = "";

		[XmlElement("v")]
		public String Value
		{
			get { return _value; }
			set { setValue(value); }
		}

		[XmlIgnore]
		public Int32 ColumnIndex;

		[XmlIgnore]
		public String Text = "";

		[XmlIgnore]
		public Double Amount;

		[XmlIgnore]
		public Boolean IsAmount;

		private void setCellReference(String value)
		{
			ColumnIndex = GetColumnIndex(value);
			if (ColumnIndex > Worksheet.MaxColumnIndex)
				Worksheet.MaxColumnIndex = ColumnIndex;
		}

		private void setValue(String value)
		{
			_value = value;
			if (tType.Equals("s"))
			{
				Text = Workbook.SharedStrings.si[Convert.ToInt32(_value)].t;
				return;
			}
			if (tType.Equals("str"))
			{
				Text = _value;
				return;
			}
			try
			{
				Amount = Convert.ToDouble(_value, CultureInfo.InvariantCulture);
				Text = Amount.ToString("#,##0.##");
				IsAmount = true;
			}
			catch (Exception ex)
			{
				Amount = 0;
				Text = String.Format("Cell Value '{0}': {1}", _value, ex.Message);
			}
		}

		private Int32 GetColumnIndex(String CellReference)
		{
			var colLetter = new Regex("[A-Za-z]+").Match(CellReference).Value.ToUpper();
			var colIndex = 0;

			for (var i = 0; i < colLetter.Length; i++)
			{
				colIndex *= 26;
				colIndex += (colLetter[i] - 'A' + 1);
			}
			return colIndex - 1;
		}
	}
}