/*
 * Erstellt mit SharpDevelop.
 * Benutzer: mri0
 * Datum: 06.10.2009
 * Zeit: 16:00
 * 
 * Sie können diese Vorlage unter Extras > Optionen > Codeerstellung > Standardheader ändern.
 */
using System;
using System.Runtime.InteropServices;

namespace MXDReader
{
	class Program
	{
		public static void Main(string[] args)
		{
			// erlaubt Umlaute in der Konsolenausgabe
			ESRI.ArcGIS.RuntimeManager.Bind(ESRI.ArcGIS.ProductCode.Desktop);
			Console.OutputEncoding = System.Text.Encoding.GetEncoding("Windows-1252");
			if (args.Length == 1 && args[0] != "/?")
			{
				MXD datei = new MXDReader.MXD(@args[0]);
			} else
			{
				Console.Error.WriteLine("Analysiert MXD-Files.");
				Console.Error.WriteLine("");
				Console.Error.WriteLine("Verwendung:");
				Console.Error.WriteLine("MXDREADER mxdfile");
			}
			
		}
	}
	
	public class Win32APICall
	{
		[DllImport("user32.dll", EntryPoint="GetDesktopWindow")]
		public static extern IntPtr GetDesktopWindow();
	}
}
