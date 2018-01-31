using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

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

		public static IDictionary<string, T> ToDictionary<T>(this object source)
		{
			if (source == null)
			{
				throw new ArgumentNullException("source", "Unable to convert object to a dictionary. The source object is null.");
			}

			var dictionary = new Dictionary<string, T>();
			foreach (PropertyDescriptor property in TypeDescriptor.GetProperties(source))
			{
				var value = property.GetValue(source) ?? "";
				if (value is T)
				{
					dictionary.Add(property.Name, (T)value);
				}
			}

			return dictionary;
		}

		public static IDictionary<string, object> ToDictionary(this object source)
		{
			return source.ToDictionary<object>();
		}

		public static string MultipleReplace(this string str, Dictionary<string, string> replacements)
		{
			var sb = new StringBuilder(str, str.Length * 2);
			replacements.Keys.ForEach(k => sb.Replace(k, replacements[k]));
			return sb.ToString();
		}

		public static string MultipleReplace(this string str, object replacements)
		{
			return str.MultipleReplace(replacements.ToDictionary());
		}

		public static string MultipleReplace(this string str, params Dictionary<string, string>[] replacements)
		{
			var result = new StringBuilder();
			foreach (var replacement in replacements)
			{
				result.Append(str.MultipleReplace(replacement));
			}
			return result.ToString();
		}

		public static string MultipleReplace(this string str, params string[] replacements)
		{
			if (replacements.Length % 2 == 0)
			{
				var dic = new Dictionary<string, string>();
				for (var i = 0; i < replacements.Length; i += 2)
				{
					dic.Add(replacements[i], replacements[i + 1]);
				}
				return str.MultipleReplace(dic);
			}
			else
			{
				throw new Exception("replacements' length must be an even value");
			}
		}
	}
}
