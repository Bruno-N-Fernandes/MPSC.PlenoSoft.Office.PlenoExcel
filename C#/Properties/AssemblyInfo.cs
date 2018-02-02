using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

#if DEBUG
[assembly: AssemblyConfiguration("Debug")]
#else
[assembly: AssemblyConfiguration("Release")]
#endif

[assembly: AssemblyTitle("Pleno Excel @ Mercado Pleno Soluções em Computação")]
[assembly: AssemblyDescription("Leitor e Gerador de Planilhas do Excel.")]
[assembly: AssemblyCompany("Mercado Pleno Soluções em Computação")]
[assembly: AssemblyProduct("MPSC.PlenoSoft.Office.dll")]
[assembly: AssemblyCopyright("Copyright © Mercado Pleno 2017")]
[assembly: AssemblyTrademark("software@mercadopleno.com.br ™")]
[assembly: AssemblyCulture("")]
[assembly: ComVisible(false)]
[assembly: Guid("7539D6F3-9829-4D41-9D2D-1DAD24A5C63A")]

[assembly: AssemblyVersion("1.0.0.115")]
[assembly: AssemblyFileVersion("16.39.12.685")]
[assembly: AssemblyInformationalVersion("1.0.2018.0202")]


public class VersaoDoAssembly
{
	public String NomeDoAssembly { get; private set; }
	public String NomeCompletoDoAssembly { get; private set; }
	public DateTime PublishDate { get; private set; }
	public String AssemblyVersion { get; private set; }
	public String ProductVersion { get; private set; }
	public String FileVersion { get; private set; }

	public VersaoDoAssembly(Assembly assembly)
	{
		var version = assembly.GetName().Version;
		var fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);

		NomeDoAssembly = Path.GetFileNameWithoutExtension(assembly.Location);
		NomeCompletoDoAssembly = assembly.FullName;
		AssemblyVersion = version.ToString();
		ProductVersion = fileVersionInfo.ProductVersion;
		FileVersion = fileVersionInfo.FileVersion;
		PublishDate = new FileInfo(assembly.Location).CreationTime;
	}

	public override string ToString()
	{
		return String.Format("{0} - {1} ({2} - {3})", NomeDoAssembly, AssemblyVersion, ProductVersion, FileVersion);
	}
}

public static class VersaoDoSistema
{
	public static IEnumerable<VersaoDoAssembly> Obter(String nomeDoAssembly)
	{
		return ObterTodosOsAssemblies()
			.Where(a => a.FullName.ToUpper().Contains(nomeDoAssembly.ToUpper()))
			.Select(ObterVersaoDoAssembly)
			.Where(v => v != null)
			.OrderBy(v => v.NomeDoAssembly);
	}

	public static VersaoDoAssembly ObterVersaoDoAssembly(this Assembly assembly)
	{
		try { return new VersaoDoAssembly(assembly); }
		catch (Exception)
		{
			return null;
		}
	}

	private static IEnumerable<Assembly> ObterTodosOsAssemblies()
	{
		return AppDomain.CurrentDomain.GetAssemblies();
	}
}