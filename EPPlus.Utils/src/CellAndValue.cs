using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Utils.src
{
	public static class CellAndValue
	{
		public static void ReplaceValue(this ExcelRange range, string target, string value)
		{
			range.Value = range.Value.ToString().Replace(target, value);
		}

		public static void ReplaceValue(this ExcelRangeBase range, string target, string value)
		{
			if (range.Value == null) return;
			range.Value = range.Value.ToString().Replace(target, value);
		}

		public static void MultipleReplaceValue(this ExcelRangeBase range, Dictionary<string, string> replacements)
		{
			if (range.Value == null) return;
			range.Value = range.Value.ToString().MultipleReplace(replacements);
		}

		public static void MultipleReplaceValue(this ExcelRangeBase range, object replacements)
		{
			if (range.Value == null) return;
			range.Value = range.Value.ToString().MultipleReplace(replacements);
		}

		public static void MultipleReplaceValue(this ExcelRangeBase range, params string[] replacements)
		{
			if (range.Value == null) return;
			range.Value = range.Value.ToString().MultipleReplace(replacements);
		}

		public static void SetValue(this ExcelRange range, string value)
		{
			range.Value = value;
		}

		public static void SetValue(this ExcelRangeBase range, object value)
		{
			range.Value = value;
		}

		public static void SetWidth(this ExcelRangeBase cell, double width)
		{
			cell.Worksheet.Column(cell.Address.AddressToNumber()[1]).SetWidth(width);
		}

		public static void SetHeight(this ExcelRangeBase cell, double height)
		{
			cell.Worksheet.Row(cell.Address.AddressToNumber()[0]).SetHeight(height);
		}

		public static bool WithValue(this ExcelRangeBase range, string value)
		{
			return range.NotNullOrEmpty() && range.Value.ToString().Contains(value);
		}

		public static bool NotNullOrEmpty(this ExcelRangeBase range)
		{
			return range.Value != null && !range.Value.Equals("");
		}

		public static bool WithinValues(this ExcelRangeBase range, params string[] values)
		{
			return range.NotNullOrEmpty() && range.Value.ToString().In(values);
		}

		public static int[] AddressToNumber(this string text, bool reverse = false)
		{
			var col = text.Where(char.IsLetter).Select(c => c - 'A' + 1).Aggregate((sum, next) => sum * 26 + next);
			var row = int.Parse(string.Join("", text.Where(char.IsDigit)));
			return reverse ? new[] { col, row } : new[] { row, col };
		}

		public static string GetColumnLetter(int column)
		{
			if (column < 1) return string.Empty;
			return GetColumnLetter((column - 1) / 26) + (char)('A' + (column - 1) % 26);
		}

		public static IEnumerable<string> Seperate(this int[] addr)
		{
			return addr.Select((a, index) => index == 0 ? GetColumnLetter(a) : a.ToString());
		}

		public static string NumberToAddress(this int[] index)
		{
			var CS = GetColumnLetter(index[1]);
			var CE = GetColumnLetter(index[3]);
			return $"{CS}{index[0]}:{CE}{index[2]}";
		}
	}
}
