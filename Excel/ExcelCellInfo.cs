using System;
using System.Collections.Generic;

namespace MGEditor
{
	public class ExcelCell : Object
	{
		public string content { get; set; }
		public string format  { get; set; }
		public string prefix  { get; set; }
		public string formula { get; set; }

		public void Fill(string nContent, string nFormat, string nPrefix, string nFormula)
		{
			content = nContent;
			format  = nFormat;
			prefix  = nPrefix;
			formula = nFormula;
		}

		public ExcelCell ()
		{
			Fill ("", "", "", "");
		}

		public ExcelCell (string nContent)
		{
			Fill (nContent, "", "", "");
		}

		public ExcelCell (string nContent, string nFormat)
		{
			Fill (nContent, nFormat, "", "");
		}

		public ExcelCell (string nContent, string nFormat, string nPrefix, string nFormula)
		{
			Fill (nContent, nFormat, nPrefix, nFormula);
		}

		public static bool operator ==(ExcelCell x, ExcelCell y)
		{
			return x.IsEqual (y);
		}
		public static bool operator !=(ExcelCell x, ExcelCell y)
		{
			return !x.IsEqual (y);
		}


		#region operators == and != handles
		public override int GetHashCode()
		{
			unchecked
			{
				int hash = 17;
				hash = hash * 23 + (prefix + content).GetHashCode();
				hash = hash * 23 + formula.GetHashCode();
				hash = hash * 23 + format.GetHashCode();
				return hash;
			}
		}
		public override bool Equals(object otherObj)
		{
			if (!(otherObj is ExcelCell))
				return false;
			return this == (ExcelCell)otherObj;
		}

		private bool IsEqual(ExcelCell other)
		{
			if (formula != "" || other.formula != "") 
			{
				if (formula != other.formula)
					return false;
				if (format != "" || other.format != "") 
				{
					if (format != other.format)
						return false;
				}
				return true;
			}
			if (content != "" || other.content != "") 
			{
				if  (prefix + content != other.prefix + other.content)
					return false;
			}
			if (format != "" || other.format != "") 
			{
				return format == other.format;
			}
			else
				return true;
		}
		#endregion
	}




	#region ExcelCellInfo
	public class ExcelCellInfo : ExcelCell,  IComparable
	{
		public int row { get; set; }
		public int column { get; set; }


		public ExcelCellInfo () : base()
		{
			row = -1;
			column = -1;
		}

		public ExcelCellInfo (int nRow, int nColumn, string nContent) : base(nContent)
		{
			row = nRow;
			column = nColumn;
		}

		public ExcelCellInfo (int nRow, int nColumn, string nContent, string nFormat)
			: base(nContent, nFormat)
		{
			row = nRow;
			column = nColumn;
		}

		public ExcelCellInfo (int nRow, int nColumn, string nContent, string nFormat, string nPrefix, string nFormula) 
			: base(nContent, nFormat, nPrefix, nFormula)
		{
			row = nRow;
			column = nColumn;
		}

		public ExcelCellInfo (string info)
		{
			row = 0;
			column = 0;
			string [] fields= info.Split ('\t');
			if (fields.Length < 5)
				return;

			int indx = fields [0].IndexOf ("C");
			try {
				row= Convert.ToInt32(fields [0].Substring (1, indx - 1));
				column= Convert.ToInt32(fields [0].Substring (indx+1));
			}   
			catch (FormatException) {
				return;
			}

			content = fields [1];
			format  = fields [2];
			prefix  = fields [3];
			formula = fields [4];
		}

		public string Address {
			get{
				if (row <= 0 || column <= 0)
					return "";
				else
					return "R" + row.ToString () + "C" + column.ToString ();
			}
		}

		public void Print ()
		{
			if (row <= 0 || column <= 0)
				System.Console.Out.WriteLine ("Wrong data");
			else 
			{
				string prfx = prefix == "" ? " " : prefix;
				Console.Out.Write ("R{0}C{1}: {2}{3}", row, column, prfx, content);
				if (format != "")
					Console.Out.Write (" format={0}", format);
				if (formula != "")
					Console.Out.WriteLine (" <--- {0}", formula);
				else
					Console.Out.WriteLine ("");
			}
		}

		public int CompareTo(object obj)
		{
			ExcelCellInfo orderToCompare = obj as ExcelCellInfo;
			if (orderToCompare.row < row )
			{
				return 1;
			}
			if (orderToCompare.row == row) 
			{
				if (orderToCompare.column < column)
					return 1;
				if (orderToCompare.column == column)
					return 0;
				else
					return -1;
			}
			else
				return -1;
		}
	}
	#endregion
}

