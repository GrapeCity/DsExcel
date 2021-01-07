using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GCExcelBenchMark
{
    class NPOIBenchmark
    {
		
		public static void TestSetRangeValues_Double(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
		{
			Console.WriteLine();
			Console.WriteLine(string.Format("NPOI benchmark for double values with {0} rows and {1} columns", rowCount, columnCount));

			double startMem = GetMemory();

			XSSFWorkbook workbook = new XSSFWorkbook();
			var worksheet = workbook.CreateSheet("poi");

			Random rand = new Random();
			DateTime start = DateTime.Now;

			for (int r = 0; r < rowCount; r++)
			{
				IRow row = worksheet.CreateRow(r);
				for (int c = 0; c < columnCount; c++)
				{
					row.CreateCell(c).SetCellValue(r+c);
				}
			}
			DateTime end = DateTime.Now;

			setTime= (end - start).TotalSeconds;
			Console.WriteLine(string.Format("NPOI set double values: {0:N3}s", setTime));

			start = DateTime.Now;
			for (int r = 0; r < rowCount; r++)
			{
				IRow row = worksheet.GetRow(r);
				for (int c = 0; c < columnCount; c++)
				{
					double d= row.GetCell(c).NumericCellValue;
				}
			}
			end = DateTime.Now;
			getTime = (end - start).TotalSeconds;
			Console.WriteLine(string.Format("NPOI get double values: {0:N3} ms", getTime));

			start = DateTime.Now;
			// Write the output to a file
			FileStream fileOut = new FileStream("../../output/poi-saved-doubles.xlsx", FileMode.Create, FileAccess.Write);
			workbook.Write(fileOut);
			fileOut.Close();

			// Closing the workbook
			workbook.Close();
			end = DateTime.Now;
			saveTime = (end - start).TotalSeconds;
			Console.WriteLine(string.Format("NPOI save doubles to Excel: {0:N3} ms", saveTime));


			double endMem = GetMemory();
			usedMem = (endMem - startMem) ;
			Console.WriteLine(string.Format("NPOI used memory: {0:N3} MB", usedMem));
		}

		public static void TestSetRangeValues_String(int rowCount, int columnCount,  ref double setTime,  ref double getTime,  ref double saveTime,  ref double usedMem)
		{
			Console.WriteLine();
			//JAVA TO C# CONVERTER TODO TASK: The following line has a Java format specifier which cannot be directly translated to .NET:
			//ORIGINAL LINE: System.out.println(String.format("NPOI benchmark for string values with {0} rows and {1} columns", rowCount, columnCount));
			Console.WriteLine(string.Format("NPOI benchmark for string values with {0} rows and {1} columns", rowCount, columnCount));

			double startMem = GetMemory();

			XSSFWorkbook workbook = new XSSFWorkbook();
			ISheet worksheet = workbook.CreateSheet("poi");


			Random random = new Random();
			string AlphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

			DateTime start = DateTime.Now;

			for (int r = 0; r < rowCount; r++)
			{
				IRow row = worksheet.CreateRow(r);
				for (int c = 0; c < columnCount; c++)
				{
					row.CreateCell(c).SetCellValue(AlphaNumericString[random.Next(25)].ToString());
				}
			}
			DateTime end = DateTime.Now;

			setTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI set string values: {0:N3}s", setTime));

			start = DateTime.Now;
			for (int r = 0; r < rowCount; r++)
			{
				IRow row = worksheet.GetRow(r);
				for (int c = 0; c < columnCount; c++)
				{
					string s=row.GetCell(c).StringCellValue;
				}
			}
			end = DateTime.Now;

			getTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI get string values: {0:N3} ms", getTime));

			start = DateTime.Now;
			// Write the output to a file
			FileStream fileOut = new FileStream("../../output/poi-saved-string.xlsx", FileMode.Create, FileAccess.Write);
			workbook.Write(fileOut);
			fileOut.Close();

			// Closing the workbook
			workbook.Close();
			end = DateTime.Now;
			saveTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI save string to Excel: {0:N3} ms", saveTime));


			double endMem = GetMemory();
			usedMem = (endMem - startMem);
			Console.WriteLine(string.Format("NPOI used memory: {0:N3} MB", usedMem));
		}

        public static void TestSetRangeValues_Date(int rowCount, int columnCount,  ref double setTime,  ref double getTime,  ref double saveTime,  ref double usedMem)
		{
			Console.WriteLine();
			//JAVA TO C# CONVERTER TODO TASK: The following line has a Java format specifier which cannot be directly translated to .NET:
			//ORIGINAL LINE: System.out.println(String.format("NPOI benchmark for date values with {0} rows and {1} columns", rowCount, columnCount));
			Console.WriteLine(string.Format("NPOI benchmark for date values with {0} rows and {1} columns", rowCount, columnCount));

			double startMem = GetMemory();

			XSSFWorkbook workbook = new XSSFWorkbook();
			ISheet worksheet = workbook.CreateSheet("poi");

			DateTime start = DateTime.Now;

			for (int r = 0; r < rowCount; r++)
			{
				IRow row = worksheet.CreateRow(r);
				for (int c = 0; c < columnCount; c++)
				{
					row.CreateCell(c).SetCellValue(DateTime.Now);
				}
			}
			DateTime end = DateTime.Now;

			setTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI set date values: {0:N3}s", setTime));

			start = DateTime.Now;
			for (int r = 0; r < rowCount; r++)
			{
				IRow row = worksheet.GetRow(r);
				for (int c = 0; c < columnCount; c++)
				{
					DateTime d= row.GetCell(c).DateCellValue;
				}
			}
			end = DateTime.Now;

			getTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI get date values: {0:N3} ms", getTime));

			start = DateTime.Now;
			// Write the output to a file
			FileStream fileOut = new FileStream("../../output/poi-saved-doubles.xlsx", FileMode.Create, FileAccess.Write);
			workbook.Write(fileOut);
			fileOut.Close();

			// Closing the workbook
			workbook.Close();
			end = DateTime.Now;
			saveTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI save date to Excel: {0:N3} ms", saveTime));


			double endMem = GetMemory();
			usedMem = (endMem - startMem) / 1024 / 1024;
			//        System.out.println(String.format("NPOI used memory: %.1f MB", usedMem.value));
		}

		public static void TestSetRangeFormulas(int rowCount, int columnCount,  ref double setTime,  ref double calcTime,  ref double saveTime,  ref double usedMem)
		{
			Console.WriteLine();
			Console.WriteLine(string.Format("NPOI benchmark for formulas values with {0} rows and {1} columns", rowCount, columnCount));

			double startMem = GetMemory();

			XSSFWorkbook workbook = new XSSFWorkbook();
			ISheet worksheet = workbook.CreateSheet("poi");

			Random rand = new Random();


			for (int r = 0; r < rowCount; r++)
			{
				IRow row = worksheet.CreateRow(r);
				for (int c = 0; c < 2; c++)
				{
					row.CreateCell(c).SetCellValue(r + c);
				}
			}
			
			DateTime start = DateTime.Now;

            for (int r = 0; r < rowCount; r++)
            {
                IRow row = worksheet.GetRow(r);
                for (int c = 2; c < columnCount + 2; c++)
                {
                    ICell cell = row.CreateCell(c);
                    CellReference reference1 = new CellReference(r, c - 2);
                    CellReference reference2 = new CellReference(r, c - 1);
                    cell.CellFormula = string.Format("SUM({0}, {1})", reference1.FormatAsString(), reference2.FormatAsString());
                }
            }

            DateTime end = DateTime.Now;
			setTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI set formulas values: {0:N3}s", setTime));

			start = DateTime.Now;
			workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
			end = DateTime.Now;

			calcTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI calculate formulas: {0:N3} ms", calcTime));

			start = DateTime.Now;
			// Write the output to a file
			FileStream fileOut = new FileStream("../../output/poi-saved-formulas.xlsx", FileMode.Create, FileAccess.Write);
			workbook.Write(fileOut);
			fileOut.Close();

			// Closing the workbook
			workbook.Close();
			end = DateTime.Now;
			saveTime = (end - start).TotalSeconds;
			Console.WriteLine(string.Format("NPOI save formulas to Excel: {0:N3} ms", saveTime));


			double endMem = GetMemory();
			usedMem = (endMem - startMem);
			//        System.out.println(String.format("NPOI used memory: %.1f MB", usedMem.value));
		}

		public static void TestBigExcelFile(int rowCount, int columnCount,  ref double openTime,  ref double calcTime,  ref double saveTime,  ref double usedMem)
		{
			Console.WriteLine();
			Console.WriteLine(string.Format("NPOI benchmark for test-performance.xlsx which is 20.5MB with a lot of values, formulas and styles"));

			double startMem = GetMemory();


			DateTime start = DateTime.Now;
			XSSFWorkbook workbook = new XSSFWorkbook("../../files/test-performance.xlsx");
			DateTime end = DateTime.Now;

			openTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI open big Excel: {0:N3}s", openTime));

			start = DateTime.Now;
			workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
			calcTime = (end - start) .TotalSeconds;
			end = DateTime.Now;
			calcTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI calculate formulas for big Excel: {0:N3} ms", calcTime));

			start = DateTime.Now;
			// Write the output to a file
			FileStream fileOut = new FileStream("../../output/poi-saved-test-performance.xlsx", FileMode.Create, FileAccess.Write);
			workbook.Write(fileOut);                                                                                      
			fileOut.Close();

			// Closing the workbook
			workbook.Close();
			end = DateTime.Now;
			saveTime = (end - start) .TotalSeconds;
			Console.WriteLine(string.Format("NPOI save back to big Excel: {0:N3} ms", saveTime));


			double endMem = GetMemory();
			usedMem = (endMem - startMem);
			Console.WriteLine(string.Format("NPOI used memory: {0:N3} MB", usedMem));
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
}
