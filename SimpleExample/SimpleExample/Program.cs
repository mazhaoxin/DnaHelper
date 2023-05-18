/*
 * 由SharpDevelop创建。
 * 用户： Ma Zhaoxin
 * 日期: 2023/5/18
 * 时间: 21:37
 */
using System;
using System.Windows.Forms;

namespace SimpleExample
{
	/// <summary>
	/// Class with program entry point.
	/// </summary>
	internal sealed class Program
	{
		/// <summary>
		/// Program entry point.
		/// </summary>
		[STAThread]
		private static void Main(string[] args)
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}
		
	}
}
