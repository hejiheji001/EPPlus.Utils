using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace EPPlus.Utils.src
{
	public static class Range
	{
		public static ExcelRange GetRange(this ExcelWorksheet sheet, int[] index)
		{
			return sheet.Cells.GetRange(index);
		}

		public static int[] GetRangeIndex(this ExcelRange range)
		{
			return new[] { range.Start.Row, range.Start.Column, range.End.Row, range.End.Column };
		}

		public static int[] GetRangeIndex(this ExcelRangeBase range)
		{
			return new[] { range.Start.Row, range.Start.Column, range.End.Row, range.End.Column };
		}

		public static ExcelRange GetRange(this ExcelRange range, int[] index)
		{
			return range[index[0], index[1], index[2], index[3]];
		}

		public static void AllBorder(this ExcelRange range, ExcelBorderStyle borderStyle)
		{
			range.ForEach(r => r.Style.Border.BorderAround(borderStyle));
		}

		public static void BackgroundColor(this ExcelRange range, Color color, ExcelFillStyle defaultFillStyle = ExcelFillStyle.Solid)
		{
			range.Style.Fill.PatternType = defaultFillStyle;
			range.Style.Fill.BackgroundColor.SetColor(color);
		}
	}
}
