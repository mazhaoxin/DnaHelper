/*
 * 由SharpDevelop创建。
 * 用户： Ma Zhaoxin
 * 日期: 2023/5/18
 * 时间: 21:37
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using XlApp = Microsoft.Office.Interop.Excel.Application;


namespace SimpleExample
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		XlApp app;
		Workbook wb;
		Worksheet ws;
		
		const int N = 8;

		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			app = new XlApp();
			app.Visible = true;
			wb = (Workbook)app.Workbooks.Add();
			ws = (Worksheet)wb.Worksheets.Add();
		}
		
		void Button1Click(object sender, EventArgs e)
		{
			// TODO: Implement Button1Click
		
			// draw board
			var c = (Range)ws.Cells[1, 1];
			c.ColumnWidth = 0.5;
			c.RowHeight = 6;
			
			c = (Range)ws.Cells[N+2+2, N+2+2];
			c.ColumnWidth = 0.5;
			c.RowHeight = 6;
			
			for (int i = 1; i <= N+2+2; i++) {
				c = (Range)ws.Cells[1, i];
				c.Interior.Color = XlRgbColor.rgbDarkGray;
				c = (Range)ws.Cells[N+2+2, i];
				c.Interior.Color = XlRgbColor.rgbDarkGray;
				c = (Range)ws.Cells[i, 1];
				c.Interior.Color = XlRgbColor.rgbDarkGray;
				c = (Range)ws.Cells[i, N+2+2];
				c.Interior.Color = XlRgbColor.rgbDarkGray;
			}
			
			// draw label
			c = (Range)ws.Cells[2, 2];
			c.ColumnWidth = 2.5;
			c.RowHeight = 15;
			
			c = (Range)ws.Cells[N+2+1, N+2+1];
			c.ColumnWidth = 2.5;
			c.RowHeight = 15;
			
			for (int i = 2; i <= N+2+1; i++) {
				c = (Range)ws.Cells[2, i];
				c.Interior.Color = XlRgbColor.rgbDarkGoldenrod;
				c.Font.Color = XlRgbColor.rgbWhite;
				c = (Range)ws.Cells[N+2+1, i];
				c.Interior.Color = XlRgbColor.rgbDarkGoldenrod;
				c.Font.Color = XlRgbColor.rgbWhite;
				c = (Range)ws.Cells[i, 2];
				c.Interior.Color = XlRgbColor.rgbDarkGoldenrod;
				c.Font.Color = XlRgbColor.rgbWhite;
				c = (Range)ws.Cells[i, N+2+1];
				c.Interior.Color = XlRgbColor.rgbDarkGoldenrod;
				c.Font.Color = XlRgbColor.rgbWhite;
			}
			
			// draw main part
			for (int i = 3; i <= N+2; i++) {
				c = (Range)ws.Cells[3, i];
				c.ColumnWidth = 10;
				c = (Range)ws.Cells[i, 3];
				c.RowHeight = 60;
			}
			
			for (int i = 3; i <= N+2; i+=2) {
				for (int j = 3; j <= N+2; j+=2) {
					c = (Range)ws.Cells[i, j];
					c.Interior.Color = XlRgbColor.rgbGhostWhite;
				}
				for (int j = 3+1; j <= N+2; j+=2) {
					c = (Range)ws.Cells[i+1, j];
					c.Interior.Color = XlRgbColor.rgbGhostWhite;
				}
			}
			for (int i = 3; i <= N+2; i+=2) {
				for (int j = 3+1; j <= N+2; j+=2) {
					c = (Range)ws.Cells[i, j];
					c.Interior.Color = XlRgbColor.rgbLightGoldenrodYellow;
				}
				for (int j = 3; j <= N+2; j+=2) {
					c = (Range)ws.Cells[i+1, j];
					c.Interior.Color = XlRgbColor.rgbLightGoldenrodYellow;
				}
			}
		}
		
		void Button2Click(object sender, EventArgs e)
		{
			// TODO: Implement Button2Click
			var c = (Range)ws.Range["A1:L12"];
			c.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			c.VerticalAlignment = XlVAlign.xlVAlignCenter;
			
			const string s1 = "abcdefgh";
			const string s2 = "87654321";
			
			for (int i = 3; i <= N+2; i++) {
				c = (Range)ws.Cells[2, i];
				c.Formula = s1[i-3].ToString();
				c = (Range)ws.Cells[N+2+1, i];
				c.Formula = s1[i-3].ToString();
				c = (Range)ws.Cells[i, 2];
				c.Formula = s2[i-3].ToString();
				c = (Range)ws.Cells[i, N+2+1];
				c.Formula = s2[i-3].ToString();
			}
		}

		void Button3Click(object sender, EventArgs e)
		{
			// TODO: Implement Button3Click
			var c = (Range)ws.Range["A1:L12"];
			c.EntireColumn.Clear();
		}
	}
}
