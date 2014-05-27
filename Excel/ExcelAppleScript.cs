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
	public class ExcelAppleScript
	{
		public ExcelAppleScript (){}

		private static ExcelDataServer dataServer= null;
		private static string MekkoExcel= "__MekkoExcel__.xlsb!ThisWorkbook.";
		public static bool LogScriptCommands= false;

		public static void ExcelCommunicationInit()
		{
			dataServer = new ExcelDataServer ();
			dataServer.Start ();
		}

		public static bool StartExcel()
		{
			string script= Path.Combine( AppDelegate.GetResourcesDirectory(), "StartExcel.applescript");
			script += " " + ExcelDataServer.port.ToString ();
			return RunOsaScript (script);
		}

		public static bool ReStartExcel()
		{
			string script= Path.Combine( AppDelegate.GetResourcesDirectory(), "StartExcel.applescript");
			return RunOsaScript (script);
		}

		public static string CreateMacroWithArguments(string macroName, params string [] Arguments)
		{
			if (macroName != null) 
			{
				string cmd = "run VB macro \"" + MekkoExcel + macroName + "\"";
				if (Arguments != null) 
				{
					for (int i = 0; i < Arguments.Length; i++) {
						cmd += " arg" + (i+1).ToString() + " \"" + Arguments [i] + "\"";
					}
				}
				return cmd;
			}
			else
				return null;
		}

		public static bool RunMacroWithArguments(string macroName, params string [] Arguments)
		{
			string cmd= CreateMacroWithArguments(macroName, Arguments);
			if (cmd != null) 
			{
				List<string> cmdLines = new List<string> ();
				cmdLines.Add (cmd);
				System.Console.Out.WriteLine (cmd);
				return Run (cmdLines.ToArray());
			}
			else
				return Run (null);
		}

		public static string [] CreateRunMacroList(params string [] macros)
		{
			List<string> cmdList = new List<string> ();
			if (macros != null) 
			{
				foreach (string s in macros) 
				{
					if (s != "") 
					{
						string cmd = "run VB macro \"" + MekkoExcel + s + "\"";
						cmdList.Add (cmd);
					}
				}
			}
			if ( cmdList.Count > 0 )
				return cmdList.ToArray();
			else
				return null;
		}

		public static bool RunMacro(params string [] macros)
		{	string [] cmdLines= CreateRunMacroList(macros);
			return Run (cmdLines);
		}

		public static void RunMacroAsync(params string [] macros)
		{	string [] cmdLines= CreateRunMacroList(macros);
			RunAsync (cmdLines);
		}

		public static void RunAsync(params string [] cmdLines)
		{
			Thread t = new Thread( () => Run(cmdLines) );
			t.Start();
		}

		public static bool Run(params string [] cmdLines)
		{
			string script = "-e 'tell application \"Microsoft Excel\"' -e 'activate'";
			if (cmdLines != null) 
			{
				foreach(string s in cmdLines){
					if (s != "") 
					{
						script += " -e '" + s + "'";
						if ( LogScriptCommands )
							System.Console.Out.WriteLine (s);
					}
				}
			}
			script += " -e 'end tell'";
			return RunOsaScript(script);
		}

		public static bool RunOsaScript(string script)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();

			proc.EnableRaisingEvents=false; 
			proc.StartInfo.UseShellExecute = false;
			proc.StartInfo.FileName = "/usr/bin/osascript";
			proc.StartInfo.Arguments = script;
			proc.StartInfo.RedirectStandardOutput = true;
			proc.StartInfo.RedirectStandardError = true;
			procOutput = new StringBuilder("");
			proc.OutputDataReceived += new DataReceivedEventHandler(ProcOutputHandler);

			proc.Start();
			proc.BeginOutputReadLine();
			string stdError = proc.StandardError.ReadToEnd();
			proc.WaitForExit();
			proc.Close();

			string[] stringSeparators = new string[] {Environment.NewLine};
			string [] stdOut= procOutput.ToString().Split(stringSeparators, StringSplitOptions.None);
			if (stdError == "")
				return true;
			if (script.IndexOf ("CloseByMekko") == -1) 
			{
				System.Console.Out.WriteLine ("osascript {0}", script);
				System.Console.Out.WriteLine ("\n--- stderr----\n{0}", stdError);
				System.Console.Out.WriteLine ("\n--- stdout----\n{0}", stdOut);
			}
			return false;
		}

		static StringBuilder procOutput = null;
		private static void ProcOutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
		{
			if (!String.IsNullOrEmpty(outLine.Data))
			{
				procOutput.Append(Environment.NewLine + outLine.Data);
			}
		}
	}
}

