using System;
using System.Collections.Generic;
using System.Linq;
using MonoMac.CoreFoundation;
using MonoMac.Foundation;
using MonoMac.AppKit;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using MonoMac.ObjCRuntime;

namespace MGEditor
{
	#region ExcelDataReceiver
	public class ExcelDataReceiver : NSObject
	{
		private static string codeDollar    = ((char)2).ToString ();
		private static string codePound     = ((char)3).ToString ();
		private static string codePrefLeft  = ((char)5).ToString ();
		private static string codePrefRight = ((char)6).ToString ();
		private static string codePrefBlank = ((char)7).ToString ();
		private static string codeCommand   = ((char)8).ToString ();

		public static string notificationExcelChanged{ get{ return "ExcelCellsChanged"; } }
		public static string notificationExcelClosed { get{ return "ExcelClosed"; } }
		public static string notificationExcelEmpty  { get{ return "ExcelEmpty"; } }

		public List<ExcelCellInfo> dataList;

		public ExcelDataReceiver() : base() 
		{ 
			dataList = new List<ExcelCellInfo> (); 
		}

		public static void Post(string str)
		{
			ExcelDataReceiver data = new ExcelDataReceiver();

			if (str != "") 
			{
				string content = str.Replace (codeDollar, "$");
				content = content.Replace (codePound, "#");
				content = content.Replace (codePrefLeft, "\'");
				content = content.Replace (codePrefRight, "\"");
				content = content.Replace (codePrefBlank, "\\");

				if (content.Substring (0, 1) == codeCommand) 
				{
					if (content.Length == 1 || (content.Length >= 6 && content.Substring (1) == "close")) 
					{
						DispatchQueue.MainQueue.DispatchAsync (() => {
							NSNotificationCenter.DefaultCenter.PostNotificationName (notificationExcelClosed, data, null);
						});
					}
					else
						Console.WriteLine("Excel: {0}", content.Substring(1));
					return;
				}

				string[] lines = content.Split ('\n');
				foreach (string strInfo in lines) {
					if (strInfo != "") {
						ExcelCellInfo cellInfo = new ExcelCellInfo (strInfo);
						if (cellInfo.row > 0 && cellInfo.column > 0)
							data.dataList.Add (cellInfo);
					}
				}
			}

			string notificationName = data.dataList.Count > 0 ? notificationExcelChanged : notificationExcelEmpty;

			DispatchQueue.MainQueue.DispatchAsync (() => {
				NSNotificationCenter.DefaultCenter.PostNotificationName (notificationName, data, null);
			});

		}
	}
	#endregion


	#region ExcelDataServer
	public class ExcelDataServer
	{
		private static string codeEOF = ((char)4).ToString ();
		private static Socket listener= null;
		public static int port{ 
			get{
				if (listener != null)
					return ((IPEndPoint)listener.LocalEndPoint).Port;
				else
					return 0;
			} 
		}

		// Thread signal.
		private static ManualResetEvent allDone = new ManualResetEvent(false);

		public ExcelDataServer ()
		{
			IPEndPoint localEndPoint = new IPEndPoint(IPAddress.Loopback, 0);
			listener = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp );
			listener.Bind(localEndPoint);
		}

		public void Start ()
		{
			Thread tData = new Thread (StartDataListening);
			tData.Start();
		}

		private void StartDataListening()
		{
			try 
			{
				listener.Listen(100);
				while (true) 
				{
					allDone.Reset();
					listener.BeginAccept( new AsyncCallback(AcceptDataCallback), listener );
					allDone.WaitOne();
				}

			} 
			catch (Exception e) 
			{
				Console.WriteLine(e.ToString());
			}
		}

		private static void AcceptDataCallback(IAsyncResult ar) 
		{
			allDone.Set();

			// Get the socket that handles the client request.
			Socket listener = (Socket) ar.AsyncState;
			Socket handler = listener.EndAccept(ar);

			// Create the state object.
			StateObject state = new StateObject();
			state.workSocket = handler;
			handler.BeginReceive( state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(DataReadCallback), state);
		}

		private static void DataReadCallback(IAsyncResult ar) 
		{
			String content = String.Empty;

			// Retrieve the state object and the handler socket
			// from the asynchronous state object.
			StateObject state = (StateObject) ar.AsyncState;
			Socket handler = state.workSocket;

			// Read data from the client socket. 
			int bytesRead = handler.EndReceive(ar);

			if (bytesRead > 0) {
				// There  might be more data, so store the data received so far.
				state.sb.Append(Encoding.ASCII.GetString(state.buffer,0,bytesRead));

				// Check for end-of-file tag. If it is not there, read 
				// more data.
				content = state.sb.ToString();
				int indx = content.IndexOf (codeEOF);
				if (indx > -1 ) 
				{
					content= content.Substring(0, indx);
					DispatchQueue.DefaultGlobalQueue.DispatchAsync (() => {
						ExcelDataReceiver.Post (content);
					});
					handler.Close();
				} 
				else 
				{
					// Not all data received. Get more.
					handler.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(DataReadCallback), state);
				}
			}
		}
	}

	public class StateObject 
	{
		public Socket workSocket = null;
		public const int BufferSize = 1024;
		public byte[] buffer = new byte[BufferSize];
		public StringBuilder sb = new StringBuilder(); 
	}
	#endregion



	#region ExcelDataSender
	public class ExcelDataSender
	{
		private ManualResetEvent allDone;
		private bool waiting;
		private byte[] byteData;
		private string promptCmd;
		private static string codeEOF = ((char)4).ToString ();

		public ExcelDataSender (string macroName, params string [] Arguments)
		{
			waiting= true;
			allDone = null;
			byteData = null;
			promptCmd = ExcelAppleScript.CreateMacroWithArguments(macroName, Arguments);
		}

		public void Send (string data)
		{
			if (data != null && data.Length > 0 && promptCmd != null && promptCmd.Length > 0) 
			{
				byteData = Encoding.ASCII.GetBytes (data);
				allDone = new ManualResetEvent(false);

				Thread tControl = new Thread (StartListening);
				tControl.Start ();
			}
		}

		private void StartListening()
		{
			IPEndPoint localEndPoint = new IPEndPoint(IPAddress.Loopback, 0);
			Socket listener = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp );

			List<string> promptCmdLines= new List<string> ();
			try 
			{
				listener.Bind(localEndPoint);
				int port= ((IPEndPoint)listener.LocalEndPoint).Port;
				listener.Listen(100);

				promptCmdLines.Add (ExcelAppleScript.CreateMacroWithArguments ("SetMekkoSenderPort", port.ToString()));
				promptCmdLines.Add (promptCmd);
				string [] commands= promptCmdLines.ToArray();
				DispatchQueue.MainQueue.DispatchAsync (() => {
					ExcelAppleScript.Run(commands);
				});

				while (waiting) 
				{
					allDone.Reset();
					listener.BeginAccept( new AsyncCallback(ControlAcceptCallback), listener );
					allDone.WaitOne();
				}

			} 
			catch (Exception e) 
			{
				Console.WriteLine(e.ToString());
			}
		}

		private void ControlAcceptCallback(IAsyncResult ar) 
		{
			allDone.Set();
			Socket listener = (Socket) ar.AsyncState;
			Socket handler = listener.EndAccept(ar);

			StateObject state = new StateObject();
			state.workSocket = handler;
			handler.BeginReceive( state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(ReadCallback), state);
		}

		private void ReadCallback(IAsyncResult ar) 
		{
			String content = String.Empty;
			StateObject state = (StateObject) ar.AsyncState;
			Socket handler = state.workSocket;
			int bytesRead = handler.EndReceive(ar);

			if (bytesRead > 0) {
				state.sb.Append(Encoding.ASCII.GetString(state.buffer,0,bytesRead));

				content = state.sb.ToString();
				int indx = content.IndexOf (codeEOF);
				if (indx > -1 ) 
				{
					handler.BeginSend(byteData, 0, byteData.Length, 0, new AsyncCallback(SendCallback), handler);
				} 
				else 
				{
					handler.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(ReadCallback), state);
				}
			}
		}

		private void SendCallback(IAsyncResult ar) {
			try {
				// Retrieve the socket from the state object.
				Socket handler = (Socket) ar.AsyncState;

				// Complete sending the data to the remote device.
				//int bytesSent = handler.EndSend(ar);
				//Console.WriteLine("ExcelDataSender: Sent {0} bytes to Excel.", bytesSent);

				handler.EndSend(ar);
				handler.Close();
				waiting= false;

			} catch (Exception e) {
				Console.WriteLine(e.ToString());
			}
		}
		#endregion

	}
}

