using System;
using System.Collections.Generic;

namespace Excel
{
	public class DataStruct
	{
		public List<DataRow> table=new List<DataRow> ();
		public DataStruct ()
		{
		}

		public void AddRow(string _fName, string _lName, string _age)
		{
			table.Add(new DataRow(_fName,_lName,_age));
		}

		public void PrintTable()
		{
			try{
				foreach (DataRow row in table)
				{
					Console.WriteLine(row.firstName+" "+row.lastName+", "+row.age);
				}

			}catch{
			}
		}
	}

	public class DataRow
	{
		private string _firstName="";
		private string _lastName="";
		private string _age="";

		public DataRow (string __firstName, string __lastName, string __age)
		{
			_firstName = __firstName;
			_lastName = __lastName;
			_age = __age;

		}

		public string firstName{get { return _firstName; } set{_firstName=value;}}

		public string lastName{get { return _lastName; } set{_lastName=value;}}

		public string age{get { return _age; } set{_age=value;}}

	}

}

