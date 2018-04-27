using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;

namespace MPSC.PlenoSoft.Office.Planilhas.Integracao
{
	public class Cabecalho
	{
		public static readonly Func<PropertyInfo, Boolean> Filtro = (p => p.Name != "ExtensionData");

		public PropertyInfo PropertyInfo { get; set; }
		public PlenoMapa Mapeamento { get; set; }

		public Cabecalho(PropertyInfo propertyInfo, PlenoMapa mapeamento)
		{
			PropertyInfo = propertyInfo;
			Mapeamento = mapeamento;
		}

		public static Cabecalho[] ObterCabecalhos(Type type, IEnumerable<PlenoMapa> mapeamento)
		{
			var allProperties = type.GetProperties();

			if ((mapeamento != null) && mapeamento.Any())
			{
				var mapa = allProperties.Join(mapeamento, p => p.DeclaringType.Name + "." + p.Name, m => m.PropertyName, (p, m) => m.Mapear(p));
				return mapa.Select(m => new Cabecalho(m.PropertyInfo, m)).OrderBy(c => c.Mapeamento.Posicao).ToArray();
			}

			var properties = allProperties.Where(p => p.GetCustomAttributes(true).OfType<PlenoMapa>().Any());

			if (!properties.Any())
				properties = allProperties.Where(Filtro).OrderBy(p => p.Name);

			return properties.Select((p, i) => new Cabecalho(p, ObterCabecalho(p, i + 1))).OrderBy(c => c.Mapeamento.Posicao).ToArray();
		}

		private static PlenoMapa ObterCabecalho(PropertyInfo propertyInfo, int posicao)
		{
			return ObterCabecalho(propertyInfo) ?? CriarCabecalho(propertyInfo, posicao);
		}

		private static PlenoMapa ObterCabecalho(PropertyInfo propertyInfo)
		{
			return propertyInfo.GetCustomAttributes(true).OfType<PlenoMapa>().FirstOrDefault();
		}

		private static PlenoMapa CriarCabecalho(PropertyInfo propertyInfo, int posicao)
		{
			return new PlenoMapa(propertyInfo.DeclaringType.Name + "." + propertyInfo.Name, propertyInfo.Name, posicao);
		}

		/// <summary>
		/// Exemplo:
		/// &lt;Mapeamento&gt;
		/// 	&lt;Class Name="Classe1"&gt;
		/// 		&lt;Property Name="Propriedade1'&gt;Nome&lt;/Property&gt;
		/// 		&lt;Property Name="Propriedade2'&gt;Opções&lt;/Property&gt;
		/// 	&lt;/Class&gt;
		/// 	&lt;Class Name="Pessoa"&gt;
		/// 		&lt;Property Name="PrimeiroNome"&gt;Primeiro Nome&lt;/Property&gt;
		/// 		&lt;Property Name="NomeCompleto"&gt;Nome Completo&lt;/Property&gt;
		/// 	&lt;/Class&gt;
		/// &lt;/Mapeamento&gt;
		/// </summary>
		/// <param name="arquivoXML">Informações da localização do arquivo XML</param>
		/// <returns>um array de todos os PlenoMapa[]</returns>
		public static PlenoMapa[] LerMapeamento(FileInfo arquivoXML)
		{
			var lista = new List<PlenoMapa>();
			var xmlDocument = new XmlDocument();
			xmlDocument.Load(arquivoXML.FullName);
			foreach (XmlNode classe in xmlDocument.DocumentElement.ChildNodes)
			{
				lista.AddRange(
					classe.ChildNodes
					.OfType<XmlNode>()
					.Select((prop, i) =>
						new PlenoMapa(
							$"{classe.Attributes["Name"].Value}.{prop.Attributes["Name"].Value}",
							$"{prop.InnerText}",
							i + 1
						)
						{ Largura = Parse(prop, "Largura", 20.0) }
					)
				);
			}
			return lista.ToArray();
		}

		private static Double Parse(XmlNode prop, String attributeName, Double value)
		{
			var attribute = prop.Attributes[attributeName];
			if (attribute != null)
				Double.TryParse(attribute.Value, out value);
			return value;
		}
	}
}