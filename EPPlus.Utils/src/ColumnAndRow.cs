using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Linq;

namespace EPPlus.Utils.src
{
	public enum InsertMode
	{
		RowBefore,
		RowAfter,
		ColumnRight,
		ColumnLeft
	}

	public static class ColumnAndRow
	{
		private static int[] GetIndex(this ExcelRangeBase range, IList valuesToInsert, InsertMode Mode, out int mode, int cellsToExpand = 0, IList expandValues = null)
		{
			mode = 0;
			var index = new int[2];

			if (Mode == InsertMode.RowBefore || Mode == InsertMode.ColumnLeft)
			{
				index = range.Start.Address.AddressToNumber();
				mode = -1;
			}

			if (Mode == InsertMode.RowAfter || Mode == InsertMode.ColumnRight)
			{
				index = range.End.Address.AddressToNumber();
				mode = 1;
			}

			if (Math.Abs(cellsToExpand) * valuesToInsert.Count != expandValues.Count)
			{
				throw new Exception("Not All Expand Cells Have Value");
			}

			return index;
		}

		private static void ExpandValue(this ExcelRangeBase range, IList valuesToInsert, int[] index, int mode, InsertMode Mode, int cellsToExpand = 0, IList expandValues = null)
		{
			for (var i = valuesToInsert.Count - 1; i >= 0; i--)
			{
				var col = 0;
				var row = 0;
				var colIndicator = 0;
				var rowIndicator = 0;

				if (Mode == InsertMode.ColumnRight || Mode == InsertMode.ColumnLeft)
				{
					col = index[1] + i + ((mode + 1) / 2);
					row = index[0];
					colIndicator = cellsToExpand == 0 ? 0 : (cellsToExpand / Math.Abs(cellsToExpand));
				}

				if (Mode == InsertMode.RowAfter || Mode == InsertMode.RowBefore)
				{
					col = index[1];
					row = index[0] + i + ((mode + 1) / 2);
					rowIndicator = cellsToExpand == 0 ? 0 : (cellsToExpand / Math.Abs(cellsToExpand));
				}

				range.Worksheet.SetValue(row, col, valuesToInsert[i]);

				for (var j = 0; j < Math.Abs(cellsToExpand); j++)
				{
					try
					{
						range.Worksheet.SetValue(row + (1 + j) * colIndicator, col + (1 + j) * rowIndicator, expandValues == null ? null : expandValues[j + i * Math.Abs(cellsToExpand)]); // You must set value to newly inserted cells, otherwise the sheet can't get its range address
					}
					catch (Exception e)
					{
						if (!e.Message.Contains("out of range"))
						{
							throw e;
						}
					}
				}
			}
		}

		/// <summary>
		/// Insert rows before or aftert current <paramref name="ExcelRangeBase"/>
		/// </summary>
		/// <param name="range">current range</param>
		/// <param name="valuesToInsert">values of each inserted row(same column as current range)</param>
		/// <param name="copyStyleFromRowIndex">the style of which row should be copied to the new inserted rows</param>
		/// <param name="Mode">insert mode, before or after</param>
		/// <param name="columnsToExpand">expand columns for inserted rows. positive for right, negative for left. Default is 0</param>
		/// <param name="expandValues">values for expanded columns</param>
		/// <returns></returns>
		public static void InsertRows(this ExcelRangeBase range, IList valuesToInsert, int copyStyleFromRowIndex, InsertMode Mode, int columnsToExpand = 0, IList expandValues = null)
		{
			var rowIndex = range.GetIndex(valuesToInsert, Mode, out int modeValue, columnsToExpand, expandValues);
			range.Worksheet.InsertRow(rowIndex[0] + ((modeValue + 1) / 2), valuesToInsert.Count, copyStyleFromRowIndex);
			range.ExpandValue(valuesToInsert, rowIndex, modeValue, Mode, columnsToExpand, expandValues);
		}

		/// <summary>
		/// Remove row which current range locates
		/// </summary>
		/// <param name="range">current range</param>
		/// <param name="valuesToInsert">current range</param>
		/// <param name="copyStyleFromColumnIndex">current range</param>
		/// <param name="Mode">current range</param>
		/// <param name="rowsToExpand">current range</param>
		/// <param name="defaultValue">current range</param>
		/// <returns></returns>
		public static void InsertColumns(this ExcelRangeBase range, IList valuesToInsert, int copyStyleFromColumnIndex, InsertMode Mode, int rowsToExpand = 0, IList expandValues = null)
		{
			var columnIndex = range.GetIndex(valuesToInsert, Mode, out int modeValue, rowsToExpand, expandValues);
			range.Worksheet.InsertColumn(columnIndex[1] + ((modeValue + 1) / 2), valuesToInsert.Count, copyStyleFromColumnIndex);
			range.ExpandValue(valuesToInsert, columnIndex, modeValue, Mode, rowsToExpand, expandValues);
		}

		//public static void InsertRows(this ExcelRangeBase range, IEnumerable<IList> listOfValuesToInsert, int copyStyleFromRowIndex, InsertMode Mode, int columnsToExpand = 0, string defaultValue = "")
		//{
		//	listOfValuesToInsert.ForEach(valuesToInsert => range.InsertOneRow(valuesToInsert, copyStyleFromRowIndex, Mode, columnsToExpand, defaultValue));
		//}

		//public static void InsertColumns(this ExcelRangeBase range, IEnumerable<IList> listOfValuesToInsert, int copyStyleFromColumnIndex, InsertMode Mode, int rowsToExpand = 0, string defaultValue = "")
		//{
		//	listOfValuesToInsert.ForEach(valuesToInsert => range.InsertOneColumn(valuesToInsert, copyStyleFromColumnIndex, Mode, rowsToExpand, defaultValue));
		//}

		/// <summary>
		/// Remove row which current range locates
		/// </summary>
		/// <param name="range">current range</param>
		/// <returns></returns>
		public static void RemoveRangeRow(this ExcelRangeBase range)
		{
			range.Worksheet.DeleteRow(range.Start.Address.AddressToNumber()[0]);
		}

		/// <summary>
		/// Remove column which current range locates
		/// </summary>
		/// <param name="range">current range</param>
		/// <returns></returns>
		public static void RemoveRangeColumn(this ExcelRangeBase range)
		{
			range.Worksheet.DeleteColumn(range.Start.Address.AddressToNumber()[1]);
		}

		/// <summary>
		/// Set width for column
		/// </summary>
		/// <param name="row">current column</param>
		/// <param name="height">value of width</param>
		/// <returns></returns>
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

		/// <summary>
		/// Set height for row
		/// </summary>
		/// <param name="row">current row</param>
		/// <param name="height">value of height</param>
		/// <returns></returns>
		public static void SetHeight(this ExcelRow row, double height)
		{
			row.Height = height;
		}

		/// <summary>
		/// Expand the current range index to new index, only expand column range
		/// </summary>
		/// <param name="index">index of current range</param>
		/// <param name="offset">positive for right expandation, negative for left expandation</param>
		/// <returns>The new index</returns>
		public static int[] ExpandColumn(this int[] index, int offset)
		{
			var newIndex = index;
			newIndex[3] += offset;
			return newIndex;
		}

		/// <summary>
		/// Expand the current range to new index, only expand column range
		/// </summary>
		/// <param name="cell">cell of current range</param>
		/// <param name="offset">positive for right expandation, negative for left expandation</param>
		/// <returns>The <paramref name="ExcelRangeBase"/> with new index</returns>
		public static ExcelRangeBase ExpandColumn(this ExcelRangeBase cell, int offset)
		{
			var newCellIndex = cell.GetRangeIndex().ExpandColumn(offset);
			var newCell = cell.Worksheet.GetRange(newCellIndex);
			return newCell;
		}

		/// <summary>
		/// Expand the current range index to new index, only expand row range
		/// </summary>
		/// <param name="index">index of current range</param>
		/// <param name="offset">positive for down expandation, negative for up expandation</param>
		/// <returns>The new index</returns>
		public static int[] ExpandRow(this int[] index, int offset)
		{
			var newIndex = index;
			newIndex[2] += offset;
			return newIndex;
		}

		/// <summary>
		/// Expand the current range to new index, only expand row range
		/// </summary>
		/// <param name="cell">cell of current range</param>
		/// <param name="offset">positive for down expandation, negative for up expandation</param>
		/// <returns>The <paramref name="ExcelRangeBase"/> with new index</returns>
		public static ExcelRangeBase ExpandRow(this ExcelRangeBase cell, int offset)
		{
			var newCellIndex = cell.GetRangeIndex().ExpandRow(offset);
			var newCell = cell.Worksheet.GetRange(newCellIndex);
			return newCell;
		}

		/// <summary>
		/// Move the current range index to another index
		/// </summary>
		/// <param name="index">index of current range</param>
		/// <param name="offset">positive for move right, negative for move left</param>
		/// <returns>The new index</returns>
		public static int[] MoveColumn(this int[] index, int offset)
		{
			var newIndex = index;
			newIndex[1] += offset;
			newIndex[3] += offset;
			return newIndex;
		}

		/// <summary>
		/// Move the current range index to another index
		/// </summary>
		/// <param name="cell">cell of current range</param>
		/// <param name="offset">positive for move right, negative for move left</param>
		/// <returns>The <paramref name="ExcelRangeBase"/> with new index</returns>
		public static ExcelRangeBase MoveColumn(this ExcelRangeBase cell, int offset)
		{
			var newCellIndex = cell.GetRangeIndex().MoveColumn(offset);
			var newCell = cell.Worksheet.GetRange(newCellIndex);
			return newCell;
		}

		/// <summary>
		/// Move the current range index to another index
		/// </summary>
		/// <param name="index">index of current range</param>
		/// <param name="offset">positive for move down, negative for move up</param>
		/// <returns>The new index</returns>
		public static int[] MoveRow(this int[] index, int offset)
		{
			var newIndex = index;
			newIndex[0] += offset;
			newIndex[2] += offset;
			return newIndex;
		}

		/// <summary>
		/// Move the current range index to another index
		/// </summary>
		/// <param name="cell">cell of current range</param>
		/// <param name="offset">positive for move down, negative for move up</param>
		/// <returns>The <paramref name="ExcelRangeBase"/> with new index</returns>
		public static ExcelRangeBase MoveRow(this ExcelRangeBase cell, int offset)
		{
			var newCellIndex = cell.GetRangeIndex().MoveRow(offset);
			var newCell = cell.Worksheet.GetRange(newCellIndex);
			return newCell;
		}

		/// <summary>
		/// copy the style of old range to new range
		/// </summary>
		/// <param name="to">the new range where style will applied to</param>
		/// <param name="offset">the old range where style will be copied</param>
		/// <returns>The <paramref name="ExcelRange"/> after new style applied</returns>
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

		/// <summary>
		/// get the row where current cell located
		/// </summary>
		/// <param name="cell">the current cell</param>
		/// <param name="size">the length of the row</param>
		/// <returns>The <paramref name="ExcelRange"/> the row</returns>
		public static ExcelRange GetRow(this ExcelRangeBase cell, int size)
		{
			var cellIndex = (cell.Address.AddressToNumber().Concat(cell.Address.AddressToNumber())).ToArray();
			return cell.Worksheet.GetRange(cellIndex.ExpandColumn(size));
		}

		/// <summary>
		/// get the column where current cell located
		/// </summary>
		/// <param name="cell">the current cell</param>
		/// <param name="size">the height of the column</param>
		/// <returns>The <paramref name="ExcelRange"/> the column</returns>
		public static ExcelRange GetColumn(this ExcelRangeBase cell, int size)
		{
			var cellIndex = (cell.Address.AddressToNumber().Concat(cell.Address.AddressToNumber())).ToArray();
			return cell.Worksheet.GetRange(cellIndex.ExpandRow(size));
		}
	}
}
