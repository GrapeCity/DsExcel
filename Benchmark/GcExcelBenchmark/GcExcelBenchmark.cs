using GrapeCity.Documents.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace GCExcelBenchMark
{
	public class GcExcelBenchmark
	{

		public static void TestSetRangeValues_Double(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
		{

			Console.WriteLine();
			//JAVA TO C# CONVERTER TODO TASK: The following line has a Java format specifier which cannot be directly translated to .NET:
			//ORIGINAL LINE: System.out.println(String.format("GcExcel benchmark for double values with {0} rows and {1} columns", rowCount, columnCount));
			Console.WriteLine(string.Format("GcExcel benchmark for double values with {0} rows and {1} columns", rowCount, columnCount));

			double startMem = GetMemory();

			IWorkbook workbook = new Workbook();
			IWorksheet worksheet = workbook.Worksheets[0];

			DateTime start = DateTime.Now;

			//JAVA TO C# CONVERTER NOTE: The following call to the 'RectangularArrays' helper class reproduces the rectangular array initialization that is automatic in Java:
			//ORIGINAL LINE: double[][] values = new double[rowCount][columnCount];
			double[,] values =new double[rowCount, columnCount];

			for (int i = 0; i < rowCount; i++)
			{
				for (int j = 0; j < columnCount; j++)
				{
					values[i,j] = i + j;
				}
			}
			
			worksheet.Range[0, 0, rowCount, columnCount].Value = values;
			DateTime end = DateTime.Now;

			setTime = (end - start).TotalSeconds;
			Console.WriteLine(string.Format("GcExcel set double values: {0:N3} s", setTime));

			start = DateTime.Now;
			object tmpValues = worksheet.Range[0, 0, rowCount, columnCount].Value;
			end = DateTime.Now;

			getTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel get double values: {0:N3} s", getTime));

			start = DateTime.Now;
			workbook.Save("../../output/gcexcel-saved-doubles.xlsx");
			end = DateTime.Now;
			saveTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel save doubles to Excel: {0:N3} s", saveTime));

			double endMem = GetMemory();
			usedMem = (endMem - startMem);
			Console.WriteLine(string.Format("GcExcel used memory: {0:N3} MB", usedMem));
		}


		public static void TestSetRangeValues_String(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
		{

			Console.WriteLine();
			//JAVA TO C# CONVERTER TODO TASK: The following line has a Java format specifier which cannot be directly translated to .NET:
			//ORIGINAL LINE: System.out.println(String.format("GcExcel benchmark for string values with {0} rows and {1} columns", rowCount, columnCount));
			Console.WriteLine(string.Format("GcExcel benchmark for string values with {0} rows and {1} columns", rowCount, columnCount));

			double startMem = GetMemory();

			IWorkbook workbook = new Workbook();
			IWorksheet worksheet = workbook.Worksheets[0];

			Random random = new Random();
			string AlphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

			DateTime start = DateTime.Now;

			//JAVA TO C# CONVERTER NOTE: The following call to the 'RectangularArrays' helper class reproduces the rectangular array initialization that is automatic in Java:
			//ORIGINAL LINE: string[][] values = new string[rowCount][columnCount];
			string[,] values = new string[rowCount, columnCount];

			for (int i = 0; i < rowCount; i++)
			{
				for (int j = 0; j < columnCount; j++)
				{
					values[i,j] = AlphaNumericString[random.Next(25)].ToString();
				}
			}

			worksheet.Range[0, 0, rowCount, columnCount].Value = values;
			DateTime end = DateTime.Now;

			setTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel set string values: {0:N3} s", setTime));

			start = DateTime.Now;
			object tmpValues = worksheet.Range[0, 0, rowCount, columnCount].Value;
			end = DateTime.Now;

			getTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel get string values: {0:N3} s", getTime));

			start = DateTime.Now;
			workbook.Save("../../output/gcexcel-saved-string.xlsx");
			end = DateTime.Now;
			saveTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel save string to Excel: {0:N3} s", saveTime));

			double endMem = GetMemory();
			usedMem = (endMem - startMem);
			Console.WriteLine(string.Format("GcExcel used memory: {0:N3} MB", usedMem));
		}

		public static void TestSetRangeValues_Date(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
		{

			Console.WriteLine();
			//JAVA TO C# CONVERTER TODO TASK: The following line has a Java format specifier which cannot be directly translated to .NET:
			//ORIGINAL LINE: System.out.println(String.format("GcExcel benchmark for date values with {0} rows and {1} columns", rowCount, columnCount));
			Console.WriteLine(string.Format("GcExcel benchmark for date values with {0} rows and {1} columns", rowCount, columnCount));

			double startMem = GetMemory();

			IWorkbook workbook = new Workbook();
			IWorksheet worksheet = workbook.Worksheets[0];

			DateTime start = DateTime.Now;

			//JAVA TO C# CONVERTER NOTE: The following call to the 'RectangularArrays' helper class reproduces the rectangular array initialization that is automatic in Java:
			//ORIGINAL LINE: DateTime[][] values = new DateTime[rowCount][columnCount];
			DateTime[,] values = new DateTime[rowCount, columnCount];

			for (int i = 0; i < rowCount; i++)
			{
				for (int j = 0; j < columnCount; j++)
				{
					values[i,j] = DateTime.Now;
				}
			}
			worksheet.Range[0, 0, rowCount, columnCount].Value = values;
			DateTime end = DateTime.Now;

			setTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel set date values: {0:N3} s", setTime));

			start = DateTime.Now;
			object tmpValues = worksheet.Range[0, 0, rowCount, columnCount].Value;
			end = DateTime.Now;

			getTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel get date values: {0:N3} s", getTime));

			start = DateTime.Now;
			workbook.Save("../../output/gcexcel-saved-date.xlsx");
			end = DateTime.Now;
			saveTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel save date to Excel: {0:N3} s", saveTime));

			double endMem = GetMemory();
			usedMem = (endMem - startMem);
			//        System.out.println(String.format("GcExcel used memory: %.1f MB", usedMem));
		}

		public static void TestSetRangeFormulas(int rowCount, int columnCount, ref double setTime, ref double calcTime, ref double saveTime, ref double usedMem)
		{

			Console.WriteLine();
			//JAVA TO C# CONVERTER TODO TASK: The following line has a Java format specifier which cannot be directly translated to .NET:
			//ORIGINAL LINE: System.out.println(String.format("GcExcel benchmark for double values with {0} rows and {1} columns", rowCount, columnCount));
			Console.WriteLine(string.Format("GcExcel benchmark for double values with {0} rows and {1} columns", rowCount, columnCount));

			double startMem = GetMemory();

			IWorkbook workbook = new Workbook();
			workbook.ReferenceStyle = ReferenceStyle.R1C1;
			IWorksheet worksheet = workbook.Worksheets[0];


			//JAVA TO C# CONVERTER NOTE: The following call to the 'RectangularArrays' helper class reproduces the rectangular array initialization that is automatic in Java:
			//ORIGINAL LINE: double[][] values = new double[rowCount][2];
			double[,] values = new double[ rowCount, 2];

			for (int i = 0; i < rowCount; i++)
			{
				for (int j = 0; j < 2; j++)
				{
					values[i,j] = i + j;
				}
			}
			worksheet.Range[0, 0, rowCount, 2].Value = values;

			DateTime start = DateTime.Now;
			worksheet.Range[0, 2, rowCount-2, columnCount].Formula = "=SUM(RC[-2],RC[-1])";
			DateTime end = DateTime.Now;

			setTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel set formulas: {0:N3} s", setTime));

			start = DateTime.Now;
			workbook.Calculate();
			end = DateTime.Now;

			calcTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel calculates formula: {0:N3} s", calcTime));

			workbook.ReferenceStyle = ReferenceStyle.A1;

			start = DateTime.Now;
			workbook.Save("../../output/gcexcel-saved-formulas.xlsx");
			end = DateTime.Now;
			saveTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel save formulas to Excel: {0:N3} s", saveTime));

			double endMem = GetMemory();
			usedMem = (endMem - startMem);
			//        System.out.println(String.format("GcExcel used memory: %.1f MB", usedMem));
		}


		public static void TestBigExcelFile(int rowCount, int columnCount, ref double openTime, ref double calcTime, ref double saveTime, ref double usedMem)
		{

			Console.WriteLine();
			Console.WriteLine(string.Format("GcExcel benchmark for test-performance.xlsx which is 20.5MB with a lot of values, formulas and styles"));

			double startMem = GetMemory();

			IWorkbook workbook = new Workbook();

			DateTime start = DateTime.Now;
			workbook.Open("../../files/test-performance.xlsx");
			DateTime end = DateTime.Now;

			openTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel open big Excel: {0:N3} s", openTime));

			start = DateTime.Now;
			workbook.Dirty();
			workbook.Calculate();
			end = DateTime.Now;

			calcTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel calculate formulas for big Excel: {0:N3} s", calcTime));

			start = DateTime.Now;
			workbook.Save("../../output/gcexcel-saved-test-performance.xlsx");
			end = DateTime.Now;
			saveTime = (end - start).TotalSeconds;;
			Console.WriteLine(string.Format("GcExcel save back to big Excel: {0:N3} s", saveTime));

			double endMem = GetMemory();
			usedMem = (endMem - startMem);
			Console.WriteLine(string.Format("GcExcel used memory: {0:N3} MB", usedMem));
		}

		public static double GetMemory()
		{
			Process proc = Process.GetCurrentProcess();
			long b = proc.PrivateMemorySize64;
			for (int i = 0; i < 2; i++)
			{
				b /= 1024;
			}
			return b;
		}
	}

	//Helper class added by Java to C# Converter:

	//---------------------------------------------------------------------------------------------------------
	//	Copyright © 2007 - 2019 Tangible Software Solutions, Inc.
	//	This class can be used by anyone provided that the copyright notice remains intact.
	//
	//	This class is used to replace calls to Java's System.currentTimeMillis with the C# equivalent.
	//	Unix time is defined as the number of seconds that have elapsed since midnight UTC, 1 January 1970.
	//---------------------------------------------------------------------------------------------------------
	internal static class DateTimeHelper
	{
		private static readonly System.DateTime Jan1st1970 = new System.DateTime(1970, 1, 1, 0, 0, 0, System.DateTimeKind.Utc);
		public static long CurrentUnixTimeMillis()
		{
			return (long)(System.DateTime.UtcNow - Jan1st1970).TotalSeconds;
		}
	}

	//Helper class added by Java to C# Converter:

	//----------------------------------------------------------------------------------------
	//	Copyright © 2007 - 2019 Tangible Software Solutions, Inc.
	//	This class can be used by anyone provided that the copyright notice remains intact.
	//
	//	This class includes methods to convert Java rectangular arrays (jagged arrays
	//	with inner arrays of the same length).
	//----------------------------------------------------------------------------------------
	internal static class RectangularArrays
	{
		public static double[][] RectangularDoubleArray(int size1, int size2)
		{
			double[][] newArray = new double[size1][];
			for (int array1 = 0; array1 < size1; array1++)
			{
				newArray[array1] = new double[size2];
			}

			return newArray;
		}

		public static string[][] RectangularStringArray(int size1, int size2)
		{
			string[][] newArray = new string[size1][];
			for (int array1 = 0; array1 < size1; array1++)
			{
				newArray[array1] = new string[size2];
			}

			return newArray;
		}

		public static DateTime[][] RectangularDateTimeArray(int size1, int size2)
		{
			DateTime[][] newArray = new DateTime[size1][];
			for (int array1 = 0; array1 < size1; array1++)
			{
				newArray[array1] = new DateTime[size2];
			}

			return newArray;
		}
	}

}
