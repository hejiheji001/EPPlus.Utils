using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace EPPlus.Utils.src
{
	public static class Range
	{
		public enum RemoveMode
		{
			RightShift,
			LowerShift,
			Column,
			Row
		}

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

		public static void RemoveRange(this ExcelRange range, RemoveMode Mode)
		{
			var sheet = range.Worksheet;

			if (Mode.Equals(RemoveMode.Column))
			{
				var cs = range.Start.Column;
				var ce = range.End.Column;
				sheet.DeleteColumn(cs);
			}

			if (Mode.Equals(RemoveMode.Row))
			{
				var rs = range.Start.Row;
				var re = range.End.Row;
				sheet.DeleteRow(rs, re - rs + 1);
			}

			if (Mode.Equals(RemoveMode.RightShift))
			{
				var cs = range.End.Column + 1;
				var ce = sheet.Dimension.Columns * 2;
				var rs = range.Start.Row;
				var re = range.End.Row;

				var RightRange = sheet.Cells[rs, cs, re, ce];

				cs = range.Start.Column;
				ce = sheet.Dimension.Columns;
				rs = range.Start.Row;
				re = range.End.Row;

				var newRange = sheet.Cells[rs, cs, re, ce];

				RightRange.Copy(newRange);

			}

			if (Mode.Equals(RemoveMode.LowerShift))
			{
				var cs = range.Start.Column;
				var ce = range.End.Column;
				var rs = range.End.Row + 1;
				var re = sheet.Dimension.Rows * 2;

				var LowerRange = sheet.Cells[rs, cs, re, ce];

				cs = range.Start.Column;
				ce = range.End.Column;
				rs = range.Start.Row;
				re = sheet.Dimension.Rows;

				var newRange = sheet.Cells[rs, cs, re, ce];

				LowerRange.Copy(newRange);
			}
		}
	}
}
