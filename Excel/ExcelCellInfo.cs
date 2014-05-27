﻿using System;
using System.Collections.Generic;

namespace MGEditor
{
	public class ExcelCell
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
	}

	public class ExcelCellInfo : ExcelCell
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
			
	}
}

