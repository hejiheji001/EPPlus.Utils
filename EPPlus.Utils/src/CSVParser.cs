using System.Linq;
using System;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace EPPlus.Utils.src
{
	public static class CSVParser
	{
		private static string FormatCSV(this string file, string delimiter = "\t", string EOL = "\r\n")
		{
			return File.ReadAllText(file, Encoding.UTF8).Replace(delimiter, ",").Replace(EOL, "\n");
		}

		private static string FormatCSV(this FileInfo file, string delimiter = "\t", string EOL = "\r\n")
		{
			return file.FullName.FormatCSV(delimiter, EOL);
		}

		public static Stream CsvToXlsStream(this string csvfile, char delimiter = ',', string EOL = "\n", string sheetName = "Data")
		{
			var excelTextFormat = new ExcelTextFormat()
			{
				Encoding = Encoding.UTF8,
				Delimiter = delimiter,
				EOL = EOL
			};

			var package = new ExcelPackage(new FileInfo($"{csvfile.Split('.')[0]}.xls"));
			var sheet = package.Workbook.Worksheets.Add(sheetName);
			sheet.Cells["A1"].LoadFromText(csvfile.FormatCSV(), excelTextFormat, OfficeOpenXml.Table.TableStyles.Medium25, false);
			var result = package.Stream;
			package.SaveAs(result);
			return result;
		}

		public static Stream CsvToXlsStream(this FileInfo csvfile, char delimiter = ',', string EOL = "\n", string sheetName = "Data")
		{
			return csvfile.FullName.CsvToXlsStream(delimiter, EOL, sheetName);
		}
	}
}
