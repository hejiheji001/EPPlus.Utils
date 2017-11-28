using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using System;
using System.Collections;

namespace EPPlus.Utils.src
{
	public static class ColumnAndRow
	{
		public static void InsertRowsBelow(this ExcelRangeBase range, IList YIndex, int XSize = 0, string defaultValue = "")
		{
			var rcs = range.Start.Address.AddressToNumber();
			var rce = range.End.Address.AddressToNumber();
			range.Worksheet.InsertRow(rcs[0] + 1, YIndex.Count, 2);
			for (var i = 0; i < YIndex.Count; i++)
			{
				range.Worksheet.SetValue(i + rcs[0] + 1, rcs[1], YIndex[i]);
				for (var j = 0; j < XSize; j++)
				{
					range.Worksheet.SetValue(i + rcs[0] + 1, rcs[1] + 1 + j, defaultValue); // You must set value to newly inserted cells, otherwise the sheet can't get its range address
				}
			}
		}

		public static int[] ExpandRow(this int[] index, int offset)
		{
			var newIndex = index;
			newIndex[2] += offset;
			return newIndex;
		}

		public static void RemoveRangeRow(this ExcelRangeBase range)
		{
			range.Worksheet.DeleteRow(range.Start.Address.AddressToNumber()[0]);
		}

		public static void SetStyle(this ExcelRangeBase cell, string format)
		{
			cell.Style.Numberformat.Format = format;
		}


		public static void SetWidth(this ExcelColumn column, double width)
		{
			var num1 = width >= 1.0 ? Math.Round((Math.Round(7.0 * (width - 0.0), 0) - 5.0) / 7.0, 2) : Math.Round((Math.Round(12.0 * (width - 0.0), 0) - Math.Round(5.0 * width, 0)) / 12.0, 2);
			var num2 = width - num1;
			var num3 = width >= 1.0 ? Math.Round(7.0 * num2 - 0.0, 0) / 7.0 : Math.Round(12.0 * num2 - 0.0, 0) / 12.0 + 0.0;
			if (num1 > 0.0)
			{
				column.Width = width + num3;
			}
			else
			{
				column.Width = 0.0;
			}
		}

		//
		// Summary:
		//     Expand the current range index to new index, only expand column range
		//
		// Parameters:
		//	 index:
		//		index of current range
		//   offset:
		//     positive for right expandation, negative for left expandation
		public static int[] ExpandColumn(this int[] index, int offset)
		{
			var newIndex = index;
			newIndex[3] += offset;
			return newIndex;
		}

		//
		// Summary:
		//     Move the current range index to another index
		//
		// Parameters:
		//	 index:
		//		index of current range
		//   offset:
		//     positive for move right, negative for move left
		public static int[] MoveColumn(this int[] index, int offset)
		{
			var newIndex = index;
			newIndex[1] += offset;
			newIndex[3] += offset;
			return newIndex;
		}

		public static ExcelRangeBase MoveColumn(this ExcelRangeBase cell, int offset)
		{
			var newCellIndex = cell.GetRangeIndex().MoveColumn(offset);
			var newCell = cell.Worksheet.GetRange(newCellIndex);
			return newCell;
		}

		//
		// Summary:
		//     Move the current range index to another index
		//
		// Parameters:
		//	 index:
		//		index of current range
		//   offset:
		//     positive for move down, negative for move up
		public static int[] MoveRow(this int[] index, int offset)
		{
			var newIndex = index;
			newIndex[0] += offset;
			newIndex[2] += offset;
			return newIndex;
		}

		//
		// Summary:
		//     copy the style of old range to new range
		//
		// Parameters:
		//	 to:
		//		the new range where style will applied to
		//   from:
		//     the old range where style will be copied
		public static ExcelRange CopyStyleFrom(this ExcelRange to, ExcelRange from, bool withValue = false)
		{
			var fromIndex = from.GetRangeIndex();
			var toIndex = to.GetRangeIndex();
			var offset = fromIndex[3] - fromIndex[1] + 1;
			var range = toIndex[3] - toIndex[1] + 1;
			var rest = range % offset;
			var loop = (range - rest) / offset;
			var index = 0;
			for (; index < loop; index++)
			{
				var tmpRange = to.GetRange(new[] { fromIndex[0], fromIndex[1] + (index + 1) * offset, fromIndex[2], fromIndex[3] + (index + 1) * offset });
				from.Copy(tmpRange);
				if (!withValue)
				{
					tmpRange.ForEach(t => t.Value = "");
				}
			}

			if (rest > 0)
			{
				var restRange = to.GetRange(new[] { fromIndex[0], fromIndex[1] + index * offset, fromIndex[2], fromIndex[1] + index * offset + rest - 1 });
				var x = from.Start.Address.AddressToNumber();
				var y = from.End.Address.AddressToNumber();
				var restFrom = from[x[0], x[1], y[0], x[1] + rest - 1];
				restFrom.Copy(restRange);
				if (!withValue)
				{
					restRange.ForEach(t => t.Value = "");
				}
			}

			return to;
		}
	}
}
