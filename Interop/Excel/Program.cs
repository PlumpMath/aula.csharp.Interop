using System;
using System.Collections.Generic;

namespace Excel
{
	class MainClass
	{
		
		public static void Main (string[] args)
		{
			DataStruct data=new DataStruct();
			IOWrite write=new IOWrite(data);

			//Попълване на данни
			data.AddRow ("Мартин", "Симеонов", "33");
			data.AddRow ("Симеон", "Мартинов", "37");

			//Проверка на таблицата
			data.PrintTable();

			write.exportTable ();
			write.runFile ();
		}
	}
}
