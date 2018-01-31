using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace MPSC.PlenoSoft.Office.ZipXml
{
	public class Workbook
	{
		public static SharedStrings SharedStrings;

		public static IEnumerable<Worksheet> Worksheets(FileInfo arquivoExcel)
		{
			using (ZipArchive zipArchive = ZipFile.Open(arquivoExcel.FullName, ZipArchiveMode.Read))
			{
				SharedStrings = DeserializedZipEntry<SharedStrings>(GetZipArchiveEntry(zipArchive, @"xl/sharedStrings.xml"));
				foreach (var worksheetEntry in (WorkSheetFileNames(zipArchive)).OrderBy(x => x.FullName))
				{
					var ws = DeserializedZipEntry<Worksheet>(worksheetEntry);
					ws.NumberOfColumns = Worksheet.MaxColumnIndex + 1;
					ws.ExpandRows();
					yield return ws;
				}
			}
		}

		public static DateTime DateFromExcelFormat(String excelCellValue)
		{
			return DateTime.FromOADate(Convert.ToDouble(excelCellValue));
		}

		private static ZipArchiveEntry GetZipArchiveEntry(ZipArchive zipArchive, String zipEntryName)
		{
			return zipArchive.Entries.FirstOrDefault(n => n.FullName.Equals(zipEntryName));
		}

		private static IEnumerable<ZipArchiveEntry> WorkSheetFileNames(ZipArchive zipArchive)
		{
			return zipArchive.Entries.Where(zipEntry => zipEntry.FullName.StartsWith("xl/worksheets/sheet"));
		}

		private static T DeserializedZipEntry<T>(ZipArchiveEntry zipArchiveEntry)
		{
			using (Stream stream = zipArchiveEntry.Open())
				return (T)new XmlSerializer(typeof(T)).Deserialize(XmlReader.Create(stream));
		}
	}
}