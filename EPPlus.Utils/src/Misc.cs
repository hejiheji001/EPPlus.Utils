using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Utils.src
{
	public static class Misc
	{
		public static bool MultiArrayEquals<T>(this T[,] left, T[,] right)
		{
			var equal =
				left.Rank == right.Rank &&
				Enumerable.Range(0, left.Rank).All(dimension => left.GetLength(dimension) == right.GetLength(dimension)) &&
				left.Cast<T>().SequenceEqual(right.Cast<T>());
			return equal;
		}

		public static bool In<T>(this T t, params T[] c)
		{
			return c.Any(i => i.Equals(t));
		}

		public static bool In<T>(this T t, IEnumerable<T> c)
		{
			return c.Any(i => i.Equals(t));
		}

		public static bool NotIn<T>(this T t, params T[] c)
		{
			return c.All(i => !i.Equals(t));
		}

		public static bool NotIn<T>(this T t, IEnumerable<T> c)
		{
			return c.All(i => !i.Equals(t));
		}
	}
}
