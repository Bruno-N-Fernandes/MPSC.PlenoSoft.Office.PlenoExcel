using System;
using System.Reflection;

namespace MPSC.PlenoSoft.Office.Planilhas.Integracao
{
	public enum TipoEnum
	{
		String,
		Int,
		Long,
		Double,
		Decimal,
		DateTime,
		ExcelGeneral
	}

	[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
	public class PlenoMapa : Attribute
	{
		public PropertyInfo PropertyInfo { get; private set; }
		public Double Largura { get; set; }
		public String PropertyName { get; set; }
		public String Titulo { get; set; }
		public Int32 Posicao { get; set; }
		public String Formato { get; set; }
		public TipoEnum Tipo { get; set; }

		public String NumberFormat { get; set; }

		public String DataFormat
		{
			get { return Formato ?? (Formato = "{0:yyyy/MM/dd}"); }
			set { Formato = value; }
		}

		public PlenoMapa(String titulo, Int32 posicao) : this(null, titulo, posicao, null, null) { }

		public PlenoMapa(String propertyName, String titulo, Int32 posicao) : this(propertyName, titulo, posicao, null, null) { }

		public PlenoMapa(String propertyName, String titulo, Int32 posicao, TipoEnum? tipo = null, String formato = null)
		{
			PropertyName = propertyName;
			Largura = 20.00;
			Titulo = titulo;
			Posicao = posicao;
			Tipo = tipo ?? TipoEnum.String;
			if (formato != null) Formato = formato;
		}

		public PlenoMapa Mapear(PropertyInfo propertyInfo)
		{
			PropertyName = propertyInfo.DeclaringType.Name + "." + propertyInfo.Name;
			PropertyInfo = propertyInfo;
			return this;
		}
	}
}