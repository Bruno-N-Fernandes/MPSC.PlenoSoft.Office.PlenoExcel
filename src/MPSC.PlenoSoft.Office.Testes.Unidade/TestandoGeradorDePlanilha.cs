using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MPSC.PlenoSoft.Office.Planilhas.Controller;
using MPSC.PlenoSoft.Office.Planilhas.Integracao;
using MPSC.PlenoSoft.Office.Planilhas.Util;

namespace MPSC.PlenoSoft.Office.Testes.Unidade
{
	[TestClass]
	public class TestandoGeradorDePlanilha
	{
		private static readonly String cRoot = File.Exists(@"C:\Temp\") ? @"C:\Temp" : Path.GetTempPath();

		[TestMethod]
		public void Quando_Converte0()
		{
			var leads = new List<Lead>();
			for (int i = 0; i < 10; i++)
			{
				var lead = ObterLead(i + 1);
				leads.Add(lead);
			}

			var arquivo1 = new FileInfo(cRoot + @"\OfficeDTO.xlsx");
			var plenoExcel1 = new PlenoExcel(arquivo1, Modo.Padrao | Modo.ApagarSeExistir);
			plenoExcel1.Exportar(leads);
			plenoExcel1.Fechar();
		}

		private static Lead ObterLead(int i)
		{
			return new Lead
			{
				bMarketinCommunication = i % 2 == 0,
				dPredictedAmount = 2.85M / i,
				dtCreatedDate = DateTime.Today.AddDays(-i * 2),
				dtUpdatedDate = DateTime.Today.AddDays(-i),
				iLeadId = i,
				iLeadSourceId = i * 2,
				iLeadStatus = i % 5,
				iOwnerId = i * 4 - 3,
				iServiceTechnicianId = i % 5 - 2,
				nvchAreaCode = string.Join("", Enumerable.Range(1, i * 7 % 29 + 4).Select(j => (char)(j + 64))),
				nvchAreaDesc = string.Join("", Enumerable.Range(1, i * 7 % 28 + 5).Select(j => (char)(j + 64))),
				nvchContactName = string.Join("", Enumerable.Range(1, i * 7 % 27 + 6).Select(j => (char)(j + 64))),
				nvchCreatedBy = string.Join("", Enumerable.Range(1, i * 7 % 26 + 7).Select(j => (char)(j + 64))),
				nvchCustomerNo_ = string.Join("", Enumerable.Range(1, i * 7 % 25 + 8).Select(j => (char)(j + 64))),
				nvchEmail = string.Join("", Enumerable.Range(1, i * 7 % 24 + 9).Select(j => (char)(j + 64))),
				nvchLeadServices = string.Join("", Enumerable.Range(1, i * 7 % 23 + 0).Select(j => (char)(j + 64))),
				nvchLeadSourceType = string.Join("", Enumerable.Range(1, i * 7 % 22 + 1).Select(j => (char)(j + 64))),
				nvchLeadStatus = string.Join("", Enumerable.Range(1, i * 7 % 21 + 2).Select(j => (char)(j + 64))),
				nvchLeadTypeCode = string.Join("", Enumerable.Range(1, i * 7 % 20 + 3).Select(j => (char)(j + 64))),
				nvchName = string.Join("", Enumerable.Range(1, i * 7 % 19 + 4).Select(j => (char)(j + 64))),
				nvchOwnerName = string.Join("", Enumerable.Range(1, i * 7 % 18 + 5).Select(j => (char)(j + 64))),
				nvchPhone = string.Join("", Enumerable.Range(1, i * 7 % 17 + 6).Select(j => (char)(j + 64))),
				nvchReason = string.Join("", Enumerable.Range(1, i * 7 % 19 + 7).Select(j => (char)(j + 64))),
				nvchServiceTechnicianName = string.Join("", Enumerable.Range(1, i * 7 % 7 + 8).Select(j => (char)(j + 64))),
				nvchUpdatedBy = string.Join("", Enumerable.Range(1, i * 7 % 15 + 9).Select(j => (char)(j + 64))),
				nvchUser1 = string.Join("", Enumerable.Range(1, i * 7 % 14 + 0).Select(j => (char)(j + 64))),
				nvchUser2 = string.Join("", Enumerable.Range(1, i * 7 % 13 + 1).Select(j => (char)(j + 64))),
				nvchUser3 = string.Join("", Enumerable.Range(1, i * 7 % 12 + 2).Select(j => (char)(j + 64))),
				nvchUser4 = string.Join("", Enumerable.Range(1, i * 7 % 11 + 3).Select(j => (char)(j + 64))),
				nvchUser5 = string.Join("", Enumerable.Range(1, i * 7 % 10 + 4).Select(j => (char)(j + 64))),
				nvchVATNo = string.Join("", Enumerable.Range(1, i * 7 % 9 + 5).Select(j => (char)(j + 64))),
				nvchZipCode = string.Join("", Enumerable.Range(1, i * 7 % 8 + 6).Select(j => (char)(j + 64)))
			};
		}

		[TestMethod]
		public void Quando_Converte()
		{
			Validar(1);
			Validar(2);
			Validar(3);
			Validar(24);
			Validar(25);
			Validar(26);
			Validar(27);
			Validar(28);
			Validar(29);
		}

		private void Validar(Int32 c0)
		{
			var c1 = Coluna.ObterNomePor(c0);
			var c2 = Coluna.ObterIndicePor(c1);

			Assert.AreEqual(c0, c2, "{0}: {1} - {2}", c0, c1, c2);
			Console.WriteLine("{0}: {1}", c0, c1);
		}

		[TestMethod]
		public void Quando_Grava_Uma_Planilha_Excel()
		{
			var arquivoExcel = new FileInfo(cRoot + @"\PlenoExcel.xlsx");
			var plenoExcel = new PlenoExcel(arquivoExcel, Modo.Seguro | Modo.SempreCriaNovo);

			var plan1 = plenoExcel["Plan1"];

			plan1.Escrever("A", 1, "Numero 1", Style.Header);
			plan1.Escrever("B", 1, "Número 2", Style.Header);
			plan1.Escrever("C", 1, "Soma", Style.Header);

			plan1.Escrever("A", 2, 6, Style.Geral);
			plan1.Escrever("B", 2, 4, Style.Geral);
			plan1.Escrever("C", 2, "=SUM(A2:B2)", Style.Geral);

			plenoExcel.Salvar();
			plenoExcel.Fechar();
		}

		[TestMethod]
		public void Quando_Grava_Uma_Lista_De_Dados_Em_Uma_Planilha_Excel()
		{
			//LogicalCell.Configurar("Não", "Sim");
			var mapeamento = new PlenoMapa[]
			{
				new PlenoMapa("Package.DateOrder", "D", 1),
				new PlenoMapa("Package.Company", "Company", 5)
			};

			var arquivo = new FileInfo(cRoot + @"\Office.xlsx");
			var packages = ObterDados();
			var plenoExcel = new PlenoExcel(arquivo, Modo.Padrao | Modo.ApagarSeExistir);

			var plan1 = plenoExcel["Plan1"];

			plan1.AdicionarDados(packages, mapeamento);
			plan1.DefinirTamanhoColunas(40, 20, 20, 25, 15, 15);
			plan1.Escrever("F", 1, "Fórmula", Style.Header);
			plan1.Escrever("A", 9, "= SUM(A2:A8)", Style.Geral);
			plan1.Escrever("C", 9, "= SUM(C2:C8)", Style.Geral);

			plenoExcel.Fechar();
			Console.WriteLine("Completed");

			var excel = new PlenoExcel(arquivo, Modo.Padrao);
			var plan01 = excel["Plan1"];
			var a1 = plan01.Ler("A", 1);
			var a2 = plan01.Ler("A", 2);
			var a3 = plan01.Ler("A", 3);
			var b1 = plan01.Ler("B", 1);
			var b2 = plan01.Ler("B", 2);
			var b3 = plan01.Ler("B", 3);
			var b9 = plan01.Ler("A", 9);
			var c9 = plan01.Ler("C", 9);

			excel.Fechar();
			Console.WriteLine("Completed");
		}

		[TestMethod]
		public void Quando_Exporta_Uma_Lista_De_Dados_Em_Uma_Planilha_Excel()
		{
			var mapeamento = new PlenoMapa[]
			{
				new PlenoMapa("Listagem.MyProperty", "Propriedade 1", 3),
				new PlenoMapa("Listagem.Packages2", "Pacote 2", 1),
				new PlenoMapa("Listagem.Packages1", "Pacote 1", 2),
				new PlenoMapa("Package.DateOrder", "Data", 1),
				new PlenoMapa("Package.Company", "Company", 5)
			};

			var listagem = new Listagem
			{
				MyProperty = 52,
				Packages1 = ObterDados(),
				Packages2 = ObterDados(),
			};

			var arquivo1 = new FileInfo(cRoot + @"\OfficeExport1.xlsx");
			var plenoExcel1 = new PlenoExcel(arquivo1, Modo.Padrao | Modo.ApagarSeExistir);
			plenoExcel1.Exportar(listagem);
			plenoExcel1.Fechar();

			var arquivo2 = new FileInfo(cRoot + @"\OfficeExport2.xlsx");
			var plenoExcel2 = new PlenoExcel(arquivo2, Modo.Padrao | Modo.ApagarSeExistir);
			plenoExcel2.Exportar(listagem, mapeamento);
			plenoExcel2.Fechar();
		}

		private static List<Package> ObterDados()
		{
			return new List<Package>
			{
				new Package("Coho Vineyard Ltd1", 25.250, 0089453312L, DateTime.Today, false ),
				new Package("Coho Vineyard Ltd2", 25.250, 0089453312L, DateTime.Today, false ),
				new Package("Lucerne Publishing", 18.778, 0089112755L, DateTime.Today, false ),
				new Package("Wingtip Toys Ltda.", 06.000, 0299456122L, DateTime.Today, false ),
				new Package("Adventure Works ME", 33.812, 4665518773L, DateTime.Today.AddDays(-4), true ),
				new Package("Test Works Ltda ME", 89.823, 4665518774L, DateTime.Today.AddDays(-2), true ),
				new Package("Good Works Ltda ME", 48.789, 4665518775L, DateTime.Today.AddDays(-1), true )
			};
		}
	}

	public class Listagem
	{
		public int MyProperty { get; set; }

		public List<Package> Packages1 { get; set; }
		public List<Package> Packages2 { get; set; }
	}

	public class Package
	{
		[PlenoMapa("Empresa", 2, Largura = 40)]
		public string Company { get; set; }

		[PlenoMapa("Weight", 1, Largura = 10)]
		public double Weight { get; set; }

		[PlenoMapa("TrackingNumber", 4, Largura = 15)]
		public long TrackingNumber { get; set; }

		[PlenoMapa("Data", 3, Largura = 25)]
		public DateTime DateOrder { get; set; }

		[PlenoMapa("HasCompleted", 5, Largura = 20)]
		public bool HasCompleted { get; set; }

		public Package(String company, double weight, long trackingNumber, DateTime dateOrder, Boolean hasCompleted)
		{
			Company = company;
			Weight = weight;
			TrackingNumber = trackingNumber;
			DateOrder = dateOrder;
			HasCompleted = hasCompleted;
		}
	}

	[Serializable]
	public class Lead
	{
		[Display(Name = "ID")]
		public int iLeadId { get; set; }
		[Required]
		[Display(Name = "Type")]
		public string nvchLeadTypeCode { get; set; }
		public string nvchCustomerNo_ { get; set; }
		[Required]
		[StringLength(50, ErrorMessage = "VAT should be less characters in length.")]
		[Display(Name = "VAT")]
		public string nvchVATNo { get; set; }
		[Required]
		[Display(Name = "Name")]
		[StringLength(50, ErrorMessage = "Name should be less characters in length.")]
		public string nvchName { get; set; }
		[Required]
		[StringLength(20, ErrorMessage = "Phone should be less characters in length.")]
		[Display(Name = "Phone")]
		public string nvchPhone { get; set; }
		[Required]
		[Display(Name = "Zip Code")]
		[StringLength(10, ErrorMessage = "Zip Code should be less characters in length.")]
		public string nvchZipCode { get; set; }
		[Required]
		[Display(Name = "Contact Name")]
		[StringLength(50, ErrorMessage = "Contact name should be less characters in length.")]
		public string nvchContactName { get; set; }
		[Required]
		[Display(Name = "E-mail")]
		[RegularExpression(@"^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$", ErrorMessage = "Email is not valid.")]
		[StringLength(50, ErrorMessage = "E-mail should be less characters in length.")]
		public string nvchEmail { get; set; }
		[Required]
		[Display(Name = "Marketing Communication")]
		public bool bMarketinCommunication { get; set; }
		[Required]
		[Display(Name = "Area Code")]
		public string nvchAreaCode { get; set; }
		[Display(Name = "Area")]
		public string nvchAreaDesc { get; set; }
		[Required]
		[Display(Name = "Reason")]
		[StringLength(1000, ErrorMessage = "Reason should be less characters in length.")]
		public string nvchReason { get; set; }
		[Required]
		[Display(Name = "Source")]
		public int iLeadSourceId { get; set; }
		[Required]
		[Display(Name = "Source Type")]
		public string nvchLeadSourceType { get; set; }
		public Nullable<int> iServiceTechnicianId { get; set; }
		[Display(Name = "Technician")]
		public string nvchServiceTechnicianName { get; set; }
		[Required]
		[Display(Name = "Owner")]
		public int iOwnerId { get; set; }
		[Display(Name = "Owner")]
		public string nvchOwnerName { get; set; }
		public string nvchCreatedBy { get; set; }
		public System.DateTime dtCreatedDate { get; set; }
		public string nvchUpdatedBy { get; set; }
		public Nullable<System.DateTime> dtUpdatedDate { get; set; }
		public int iLeadStatus { get; set; }
		[Display(Name = "Status")]
		public string nvchLeadStatus { get; set; }
		[Required]
		[Display(Name = "List Services")]
		public string nvchLeadServices { get; set; }
		public decimal? dPredictedAmount { get; set; }
		//public List<LeadService> lstLeadServices { get; set; }
		[Required]
		[Display(Name = "Address")]
		public string nvchUser1 { get; set; }
		public string nvchUser2 { get; set; }
		public string nvchUser3 { get; set; }
		public string nvchUser4 { get; set; }
		public string nvchUser5 { get; set; }

		public Lead()
		{
			iLeadId = 0;
			iOwnerId = 0;
			iLeadSourceId = 0;
			iLeadStatus = 1;
			nvchLeadTypeCode = string.Empty;
			nvchAreaCode = string.Empty;
			nvchAreaDesc = string.Empty;
			nvchContactName = string.Empty;
			nvchCreatedBy = string.Empty;
			nvchCustomerNo_ = string.Empty;
			nvchEmail = string.Empty;
			nvchName = string.Empty;
			nvchPhone = string.Empty;
			nvchReason = string.Empty;
			nvchUpdatedBy = string.Empty;
			nvchVATNo = string.Empty;
			nvchZipCode = string.Empty;
			nvchLeadServices = string.Empty;
			dtCreatedDate = DateTime.Now;
			dtUpdatedDate = DateTime.Now;
			dPredictedAmount = 0;

			nvchUser1 = string.Empty;
			nvchUser2 = string.Empty;
			nvchUser3 = string.Empty;
			nvchUser4 = string.Empty;
			nvchUser5 = string.Empty;

			bMarketinCommunication = false;
			//lstLeadServices = new List<LeadService>();
		}
	}


}