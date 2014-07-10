using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConsoleApplication3
{
	internal sealed class Program
	{
		private Program()
		{

		}

		public static void Main(string[] args)
		{
			string path = 1 <= args.Length ? args[0] : "";
			if (path == "")
			{
				Console.WriteLine("path to Excel Workbook.");
				return;
			}

			Microsoft.Office.Interop.Excel.Application app = null;
			try
			{
				app = new Microsoft.Office.Interop.Excel.Application();
				app.Visible = false;
				app.Workbooks.Open(path, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					false, Type.Missing, Type.Missing);
				Microsoft.Office.Interop.Excel.Workbook book = app.Workbooks[1];
				for (int i = 0; i < book.Worksheets.Count; i++)
				{
					Console.WriteLine(book.Worksheets[1 + i].Name);
				}

				Console.WriteLine("" + book.Worksheets.Count + "個のシートを検出しました。");
			}
			catch (Exception e)
			{
				MessageBox.Show(e.ToString(), "エラー", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
			}
			finally
			{
				if (app != null)
				{
					app.Workbooks.Close();
					app.Quit();
				}
			}
		}
	}
}
