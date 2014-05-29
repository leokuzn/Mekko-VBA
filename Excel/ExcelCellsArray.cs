using System;
using System.Collections.Generic;

namespace MGEditor
{
	public static class ExcelCellsArray
	{
		public static ExcelCell[,] RebuildSessionArray(List<ExcelCellInfo> cellsList, out int minRow, out int minCol, out int numRow, out int numCol)
		{
			numRow = numCol = 0;
			ExcelCell[,] cellsArray = ToCellArray(cellsList, out minRow, out minCol);
			if (cellsArray == null) 
				return null;

			numRow = cellsArray.GetLength (0);
			numCol = cellsArray.GetLength (1);
			return cellsArray;
		}

		//---------------------------------
		public static string[,] ToFormatArray(List<ExcelCellInfo> cells, out int minRow, out int minCol)
		{
			if (cells == null || cells.Count == 0) {
				minRow = minCol = 0;
				return null;
			}

			minRow = minCol = int.MaxValue;
			int maxRow, maxCol;
			maxRow = maxCol = 0;
			bool found = false;

			foreach (ExcelCellInfo info in cells) 
			{
				if (info.format != "") 
				{
					found = true;
					if (info.row > maxRow)
						maxRow = info.row;
					if (info.col > maxCol)
						maxCol = info.col; 
					if (info.row < minRow)
						minRow = info.row;
					if (info.col < minCol)
						minCol = info.col; 
				}
			}
			if (!found) {
				minRow = minCol = 0;
				return null;
			}

			int numRow = maxRow - minRow + 1;
			int numCol = maxCol - minCol + 1;
			string[,] fmtArray = new string[numRow, numCol];

			for (int iR = 0; iR < numRow; iR++)
				for (int iC = 0; iC < numCol; iC++)
					fmtArray[iR, iC]= ""; 

			foreach (ExcelCellInfo info in cells) 
			{
				if (info.format != "") 
				{
					int iR = info.row - minRow;
					int iC = info.col - minCol;
					fmtArray[iR, iC]= info.format;
				}
			}

			return fmtArray;
		}

		//---------------------------------
		public static string[,] ToValueArray(List<ExcelCellInfo> cells, out int minRow, out int minCol)
		{
			if (cells == null || cells.Count == 0) {
				minRow = minCol = 0;
				return null;
			}

			minRow = minCol = int.MaxValue;
			int maxRow, maxCol;
			maxRow = maxCol = 0;
			bool found = false;

			foreach (ExcelCellInfo info in cells) 
			{
				if (info.content != "" || info.formula != "") 
				{
					found = true;
					if (info.row > maxRow)
						maxRow = info.row;
					if (info.col > maxCol)
						maxCol = info.col; 
					if (info.row < minRow)
						minRow = info.row;
					if (info.col < minCol)
						minCol = info.col; 
				}
			}
			if (!found) {
				minRow = minCol = 0;
				return null;
			}

			int numRow = maxRow - minRow + 1;
			int numCol = maxCol - minCol + 1;
			string[,] valueArray = new string[numRow, numCol];

			for (int iR = 0; iR < numRow; iR++)
				for (int iC = 0; iC < numCol; iC++)
					valueArray[iR, iC]= ""; 

			foreach (ExcelCellInfo info in cells) 
			{
				if (info.content != "" || info.formula != "") 
				{
					int iR = info.row - minRow;
					int iC = info.col - minCol;
					if (info.formula != "") {
						valueArray[iR, iC]= info.formula;
					}
					else{
						valueArray[iR, iC]= info.prefix + info.content;
					}
				}
			}

			return valueArray;
		}

		//---------------------------------
		public static ExcelCell[,] ToCellArray(List<ExcelCellInfo> cells, out int minRow, out int minCol)
		{
			if (cells == null || cells.Count == 0) {
				minRow = minCol = 0;
				return null;
			}

			minRow = minCol = int.MaxValue;
			int maxRow, maxCol;
			maxRow = maxCol = 0;

			foreach (ExcelCellInfo info in cells) 
			{
				if (info.row > maxRow)
					maxRow = info.row;
				if (info.col > maxCol)
					maxCol = info.col; 
				if (info.row < minRow)
					minRow = info.row;
				if (info.col < minCol)
					minCol = info.col; 
			}

			int numRow = maxRow - minRow + 1;
			int numCol = maxCol - minCol + 1;
			ExcelCell[,] CellsArray = new ExcelCell[numRow, numCol];

			for (int iR = 0; iR < numRow; iR++)
				for (int iC = 0; iC < numCol; iC++)
					CellsArray[iR, iC]=new ExcelCell (); 

			foreach (ExcelCellInfo info in cells) 
			{
				int iR = info.row - minRow;
				int iC = info.col - minCol;
				CellsArray[iR, iC].Fill(info.content, info.format, info.prefix, info.formula);
			}

			return CellsArray;
		}

		//---------------------------------
		public static List<ExcelCellInfo> ToCellList(ExcelCell[,] cellsArray, int minRow, int minCol)
		{
			if (cellsArray == null || minRow < 1 || minCol < 1)
				return null;

			List<ExcelCellInfo> cells= new List<ExcelCellInfo>();

			int numRow = cellsArray.GetLength (0);
			int numCol = cellsArray.GetLength (1);
			for (int iR = 0; iR < numRow; iR++) 
			{
				for (int iC = 0; iC < numCol; iC++) 
				{
					ExcelCell cell = cellsArray [iR, iC];
					if (cell.content != "" || cell.format != "" || cell.formula != "") {
						ExcelCellInfo cellInfo = new ExcelCellInfo (minRow + iR, minCol + iC, cell.content, cell.format, cell.prefix, cell.formula);
						cells.Add (cellInfo);
					}
				}
			}

			return cells;
		}

		//-------------------------------------
		public static bool CompareListAndArray(List<ExcelCellInfo> dataList, ExcelCell[,] cellsArray, int minRow, int minCol)
		{
			int numRow = cellsArray.GetLength (0);
			int numCol = cellsArray.GetLength (1);
			int maxRow = minRow + numRow - 1;
			int maxCol = minCol + numCol - 1;

			foreach (ExcelCellInfo cellInfo in dataList) 
			{
				int iR = cellInfo.row;
				int iC = cellInfo.col;

				if (iR < minRow || maxRow < iR || iC < minCol || maxCol < iC) 
					return false;

				ExcelCell cell = cellsArray [iR-minRow, iC-minCol];
				if (cell.content != cellInfo.content) 
				{
					if (cell.content == "" && cellInfo.content != "" && cell.formula != "" && cell.formula == cellInfo.formula) 
						cell.content = cellInfo.content;
					else 
						return false;
				} 
				else if ( cell.format != cellInfo.format || cell.formula != cellInfo.formula || cell.prefix != cellInfo.prefix )
					return false;
			}
			return true;
		}


		//-------------------------------------
		public static bool Compare(List<ExcelCellInfo> list1, List<ExcelCellInfo> list2)
		{
			if (list1 == null || list2 == null)
				return false;
			if (list1.Count != list2.Count)
				return false;

			list1.Sort ();
			list2.Sort ();

			for (int i=0; i < list1.Count; i++) 
			{
				ExcelCellInfo info1 = list1[i], info2 = list2[i];
				if (!info1.IsEqual (info2))
					return false;
			}
			return true;
		}


	}
}

