using System;
using MonoMac.Foundation;
using MonoMac.AppKit;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.ComponentModel;
using System.Collections.Generic;

namespace MGEditor
{
	public static class ExcelExchange
	{
		private static NSObject session= null;

		private static ExcelCell[,] CellsArray= null;
		private static List<ExcelCellInfo> CellsList= null;

		private static int minRow= 0;
		private static int minCol= 0;
		private static int numRow= 0;
		private static int numCol= 0;
		private static long timeStartSession;
		private static long timeOpenExcel;
		private static long timeFirstRespond;
		private static bool firstRun= true;


		public static List<ExcelCellInfo> TestList()
		{
			List<ExcelCellInfo> myCells = new List<ExcelCellInfo> ();
			myCells.Add (new ExcelCellInfo (2, 2, "Series1"));
			myCells.Add (new ExcelCellInfo (2, 3, "Series2"));
			myCells.Add (new ExcelCellInfo (2, 4, "Series3"));
			int j = 0;
			int numRow = 49;
			for (int i = 3; i < numRow; i++) 
			{
				myCells.Add (new ExcelCellInfo (i, 1, i.ToString()));
				myCells.Add (new ExcelCellInfo (i, 2, (51.0/i).ToString(), "0.000"));
				myCells.Add (new ExcelCellInfo (i, 3, (i+2000).ToString()));
				myCells.Add (new ExcelCellInfo (i, 4, (-14.0 + 0.5*i).ToString(), "$##,##0.00_);[Red]($#,##0.00)"));
				if (i % 5 == 0) {
					j++;
					myCells.Add (new ExcelCellInfo (i, 5, " <--Case " + j.ToString ()));
				}
				myCells.Add (new ExcelCellInfo (i,  6, (100-i).ToString()));
				myCells.Add (new ExcelCellInfo (i,  7, (200+i).ToString()));
				myCells.Add (new ExcelCellInfo (i,  8, (300-i).ToString()));
				myCells.Add (new ExcelCellInfo (i,  9, (400+i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 10, (500-i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 11, (600+i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 12, (700-i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 13, (800+i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 14, (900-i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 15, (900+i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 16, (i*i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 17, (100-i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 18, (200-i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 19, (300-i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 20, (i*i+2).ToString()));
				myCells.Add (new ExcelCellInfo (i, 21, (18+i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 22, (22+i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 23, (2-0.3*i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 24, (-i*i).ToString()));
				myCells.Add (new ExcelCellInfo (i, 25, (i*(i+1)/2).ToString()));
				myCells.Add (new ExcelCellInfo (i, 26, (i*(i-1)/2).ToString()));
			}
			myCells.Add (new ExcelCellInfo (numRow+1, 2, "", "0.000", "", "=SUM(B3:B" + (numRow-1).ToString() + ")"));
			myCells.Add (new ExcelCellInfo (numRow+1, 3, "", "", "", "=AVERAGE(C3:C" + (numRow-1).ToString() + ")"));
			myCells.Add (new ExcelCellInfo (numRow+1, 4, "", "$##,##0.00_);[Red]($#,##0.00)", "", "=SUM(D3:D" + (numRow-1).ToString() + ")"));
			myCells.Add (new ExcelCellInfo (numRow+1, 6, "", "", "", "=AVERAGE(F3:F" + (numRow-1).ToString() + ")"));
			myCells.Add (new ExcelCellInfo (numRow+1,16, "", "", "", "=SUM(P3:P" + (numRow-1).ToString() + ")"));
			return myCells;
		}

		private static void ClearSession()
		{
			CellsArray= null;
			CellsList= null;
			minRow= numRow= 0;
			minCol= numCol= 0;
			timeStartSession = DateTime.Now.Ticks;
		}

		private static void RebuildSessionArray()
		{
			CellsArray = ToCellArray(CellsList, out minRow, out minCol);
			if (CellsArray == null) 
			{
				ClearSession ();
				return;
			}

			numRow = CellsArray.GetLength (0);
			numCol = CellsArray.GetLength (1);
		}

		public static void StartSession(List<ExcelCellInfo> cells)
		{
			ClearSession ();
			if (cells == null || cells.Count == 0)
				//return;
				cells = TestList ();

			CellsList = new List<ExcelCellInfo> ();
			foreach (ExcelCellInfo info in cells) {
				CellsList.Add (info);
			}
			RebuildSessionArray ();

			string data = "";

			int minRowFmt=0, minColFmt=0, numRowFmt=0, numColFmt=0;
			string[,] fmtArray= ToFormatArray(CellsList, out minRowFmt, out minColFmt);
			if (fmtArray != null) 
			{
				numRowFmt = fmtArray.GetLength (0);
				numColFmt = fmtArray.GetLength (1);

				for (int iR = 0; iR < numRowFmt; iR++) {
					for (int iC = 0; iC < numColFmt; iC++) {
						data += fmtArray[iR,iC] + "\t";
					}
				}
			}

			int minRowVal=0, minColVal=0, numRowVal=0, numColVal=0;
			string[,] valArray= ToValueArray(CellsList, out minRowVal, out minColVal);
			if (valArray != null) 
			{
				numRowVal = valArray.GetLength (0);
				numColVal = valArray.GetLength (1);

				for (int iR = 0; iR < numRowVal; iR++) {
					for (int iC = 0; iC < numColVal; iC++) {
						data += valArray[iR,iC] + "\t";
					}
				}
			}

			session= NSNotificationCenter.DefaultCenter.AddObserver (ExcelDataReceiver.notificationName, OnExcelTableChanged);
			ExcelAppleScript.StartExcel ();
			firstRun = true;
			timeOpenExcel = DateTime.Now.Ticks;

			long elapsedTime = (timeOpenExcel - timeStartSession)/TimeSpan.TicksPerMillisecond;
			Console.Out.WriteLine ("Excel session started {0} sec", elapsedTime/1000.0);

			ExcelDataSender dataSender = new ExcelDataSender ("SetRangeValueByMekko", 
											minRowFmt.ToString (), 
											minColFmt.ToString (), 
											numRowFmt.ToString (), 
											numColFmt.ToString (),
											minRowVal.ToString (), 
											minColVal.ToString (), 
											numRowVal.ToString (), 
											numColVal.ToString ());
			dataSender.Send (data);
		}

		public static void StopSession()
		{
			if ( session != null )
				NSNotificationCenter.DefaultCenter.RemoveObserver (session);

			ClearSession ();
			ExcelAppleScript.RunMacroAsync ("HideByMekko");
			Console.Out.WriteLine ("Excel session stopped");
		}

		private static void OnExcelTableChanged(NSNotification notification)
		{
			ExcelDataReceiver data = notification.Object as ExcelDataReceiver;

			if (data.dataList.Count == 0) {
				StopSession ();
				return;
			}

			int maxRow = minRow + numRow - 1;
			int maxCol = minCol + numCol - 1;


			if (firstRun) 
			{
				timeFirstRespond = DateTime.Now.Ticks;

				long elapsedTime = (timeFirstRespond - timeOpenExcel)/TimeSpan.TicksPerMillisecond;
				Console.Out.WriteLine ("Session first respond time= {0} sec", elapsedTime/1000.0);

				firstRun = false;

				foreach (ExcelCellInfo cellInfo in data.dataList) 
				{
					int iR = cellInfo.row;
					int iC = cellInfo.column;

					if (iR < minRow || maxRow < iR || iC < minCol || maxCol < iC) 
					{
						Console.Out.WriteLine ("Excel communication error");
						StopSession ();
						return;
					}

					ExcelCell cell = CellsArray [iR-minRow, iC-minCol];
					if (cell.content != cellInfo.content) 
					{
						if (cell.content == "" && cellInfo.content != "" && cell.formula != "" && cell.formula == cellInfo.formula) 
						{
							cell.content = cellInfo.content;
						} 
						else 
						{
							Console.Out.WriteLine ("Excel communication error");
							StopSession ();
							return;
						}
					} 
					else if ( cell.format != cellInfo.format || cell.formula != cellInfo.formula || cell.prefix != cellInfo.prefix )
					{
						System.Console.Out.WriteLine ("Excel communication error");
						StopSession ();
						return;
					}
				}
				return;
			}

			List<ExcelCellInfo> cellsAdded   = new List<ExcelCellInfo> ();
			List<ExcelCellInfo> cellsChanged = new List<ExcelCellInfo> ();

			foreach (ExcelCellInfo cellInfo in data.dataList)
			{
				int iR= cellInfo.row;
				int iC = cellInfo.column;

				if (iR < minRow || maxRow < iR || iC < minCol || maxCol < iC) 
				{
					cellsAdded.Add (cellInfo);
				} 
				else 
				{
					ExcelCell cell = CellsArray [iR-minRow, iC-minCol];
					if (cell.content != cellInfo.content  ||
						cell.formula != cellInfo.formula  ||
						cell.format  != cellInfo.format   ||
						cell.prefix  != cellInfo.prefix) 
					{
						if (cell.content == "" && cell.format == "" && cell.formula == "")
							cellsAdded.Add (cellInfo);
						else 
						{
							cellsChanged.Add (cellInfo);
							cell.content = cellInfo.content;
							cell.formula = cellInfo.formula;
							cell.format  = cellInfo.format;
							cell.prefix  = cellInfo.prefix;
						}
					}
				}
			}

			CellsList= null;
			CellsList= ToCellList(CellsArray, minRow, minCol);
			CellsArray = null;
			if (cellsAdded.Count > 0) 
			{
				foreach (ExcelCellInfo cellInfo in cellsAdded) {
					CellsList.Add (cellInfo);
				}
			}
			RebuildSessionArray ();


			if (cellsAdded.Count + cellsChanged.Count > 0) 
			{
				System.Console.Out.WriteLine ("------beg----------");
				foreach (ExcelCellInfo cellInfo in cellsAdded) {
					Console.Out.Write ("Added  : ");
					cellInfo.Print ();
				}
				foreach (ExcelCellInfo cellInfo in cellsChanged) {
					Console.Out.Write ("Changed: ");
					cellInfo.Print ();
				}
				Console.Out.WriteLine ("------end---------- total {0} cells\n", cellsAdded.Count + cellsChanged.Count);
			}

			data.Dispose ();
		}

		#region Array Helper Functions
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
					if (info.column > maxCol)
						maxCol = info.column; 
					if (info.row < minRow)
						minRow = info.row;
					if (info.column < minCol)
						minCol = info.column; 
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
					int iC = info.column - minCol;
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
					if (info.column > maxCol)
						maxCol = info.column; 
					if (info.row < minRow)
						minRow = info.row;
					if (info.column < minCol)
						minCol = info.column; 
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
					int iC = info.column - minCol;
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
				if (info.column > maxCol)
					maxCol = info.column; 
				if (info.row < minRow)
					minRow = info.row;
				if (info.column < minCol)
					minCol = info.column; 
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
				int iC = info.column - minCol;
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
		#endregion
	}
}

