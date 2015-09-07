using System;
using InteropExcel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel
{
	
	public class IOWrite
	{
		private DataStruct _data;
		private InteropExcel.Application excel;

		public IOWrite (DataStruct data)
		{
			_data = data;
		}

		public bool exportTable()
		{
			try{
				excel = new InteropExcel.ApplicationClass();

				if (excel==null) return false;

				excel.Visible=false;

				InteropExcel.Workbook workbook = excel.Workbooks.Add();
				if (workbook==null) return false;

				InteropExcel.Worksheet sheet = (InteropExcel.Worksheet) workbook.Worksheets[1];
				sheet.Name="Таблица 1";
				//попълване на таблицата

				int i=1;
				addRow(new DataRow("Първо име","Фамилия","Години"),i++,true,254);i++;
				foreach (DataRow row in _data.table)
				{
					addRow(row,i++,false,-1);
				}
				i++;addRow(new DataRow("Брой редове","",_data.table.Count.ToString()),i++,true,-1);i++;
				workbook.SaveCopyAs(getPath());
				excel.DisplayAlerts=false;
				excel.Quit();

				if (workbook!=null) Marshal.ReleaseComObject(workbook);
				if (sheet!=null) Marshal.ReleaseComObject(sheet);
				if (excel!=null) Marshal.ReleaseComObject(excel);
				workbook=null;
				sheet=null;
				excel=null;

				GC.Collect();

				return true;
			}catch{
			}
			return false;
		}

		public void addRow(DataRow _row,int _indexRow,bool isBold, int color)
		{
			try {
				InteropExcel.Range range;
				//форматиране
				range=excel.Range["A"+_indexRow.ToString(),"C"+_indexRow.ToString()];
				if (color > 0) range.Interior.Color=color;
				range.Font.Bold=isBold;

				//Попълване на данните
				
				range=excel.Range["A"+_indexRow.ToString(),"A"+_indexRow.ToString()];
				range.Value2=_row.firstName;

				range=excel.Range["B"+_indexRow.ToString(),"B"+_indexRow.ToString()];
				range.Value2=_row.lastName;

				range=excel.Range["C"+_indexRow.ToString(),"C"+_indexRow.ToString()];
				range.Value2=_row.age;


			} catch {
			}

		}

		public void runFile()
		{
			try{
			System.Diagnostics.Process.Start (getPath ());
			}catch{
			}
		}

		public string getPath()
		{
			return System.IO.Path.Combine (AppDomain.CurrentDomain.BaseDirectory, "table.xlsx");
		}

	}
}

