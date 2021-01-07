using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GCExcelBenchMark
{
    public class AsposeBenchmark {

        public static void TestSetRangeValues_Double(int rowCount,
                                                     int columnCount,
                                                     ref double setTime,
                                                     ref double getTime,
                                                     ref double saveTime,
                                                     ref double usedMem)  {
            Console.WriteLine();
            Console.WriteLine(String.Format("Aspose.Cells benchmark for double values with {0} rows and {1} columns", rowCount, columnCount));

            double startMem = GetMemory();

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            DateTime start = DateTime.Now;

            Double[,] values = new Double[rowCount,columnCount];

            for (int i = 0; i < rowCount; i++) {
                for (int j = 0; j < columnCount; j++) {
                    values[i,j] = (double)(i + j);
                }
            }

            worksheet.Cells.CreateRange("A1:AC" + rowCount).Value=values;
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells set double values: {0:N3} s", setTime));

            start = DateTime.Now;
            Object tmpValues = worksheet.Cells.CreateRange("A1:AC" + rowCount).Value;
            end = DateTime.Now;

            getTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells get double values: {0:N3} s", getTime));

            start = DateTime.Now;
            workbook.Save("../../output/aspose-saved-doubles.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells save doubles to Excel: {0:N3} s", saveTime));

            double endMem = GetMemory();
            usedMem = (endMem - startMem) ;
            Console.WriteLine(String.Format("Aspose.Cells used memory: {0:N3} MB", usedMem));

        }
        public static void TestSetRangeValues_String(int rowCount,
                                                 int columnCount,
                                                 ref double setTime,
                                                 ref double getTime,
                                                 ref double saveTime,
                                                 ref double usedMem)
        {
            Console.WriteLine();
            Console.WriteLine(String.Format("Aspose.Cells benchmark for string values with {0} rows and {1} columns", rowCount, columnCount));

            double startMem = GetMemory();

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            Random random = new Random();
            String AlphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            DateTime start = DateTime.Now;

            String[,] values = new String[rowCount,columnCount];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    values[i,j] = AlphaNumericString[random.Next(25)].ToString();
                }
            }

            worksheet.Cells.CreateRange("A1:AC" + rowCount).Value=values;
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells set string values: {0:N3} s", setTime));

            start = DateTime.Now;
            Object tmpValues = worksheet.Cells.CreateRange("A1:AC" + rowCount).Value;
            end = DateTime.Now;

            getTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells get string values: {0:N3} s", getTime));

            start = DateTime.Now;
            workbook.Save("../../output/aspose-saved-string.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells save string to Excel: {0:N3} s", saveTime));

            double endMem = GetMemory();
            usedMem = (endMem - startMem) ;
            Console.WriteLine(String.Format("Aspose.Cells used memory: {0:N3} MB", usedMem));

        }

        public static void TestSetRangeValues_Date(int rowCount,
                                                     int columnCount,
                                                     ref double setTime,
                                                     ref double getTime,
                                                     ref double saveTime,
                                                     ref double usedMem)
        {
            Console.WriteLine();
            Console.WriteLine(String.Format("Aspose.Cells benchmark for date values with {0} rows and {1} columns", rowCount, columnCount));

            double startMem = GetMemory();

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            DateTime start = DateTime.Now;

            DateTime[,] values = new DateTime[rowCount,columnCount];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    values[i,j] = DateTime.Now;
                }
            }

            worksheet.Cells.CreateRange("A1:AC" + rowCount).Value=values;
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells set date values: {0:N3} s", setTime));

            start = DateTime.Now;
            Object tmpValues = worksheet.Cells.CreateRange("A1:AC" + rowCount).Value;
            end = DateTime.Now;

            getTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells get date values: {0:N3} s", getTime));

            start = DateTime.Now;
            workbook.Save("../../output/aspose-saved-date.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells save date to Excel: {0:N3} s", saveTime));

            double endMem = GetMemory();
            usedMem = (endMem - startMem) ;
            //        Console.WriteLine(String.Format("Aspose.Cells used memory: {0:N3} MB", usedMem));

        }

        public static void TestSetRangeFormulas(int rowCount,
                                                     int columnCount,
                                                     ref double setTime,
                                                     ref double calcTime,
                                                     ref double saveTime,
                                                     ref double usedMem)
        {
            Console.WriteLine();
            Console.WriteLine(String.Format("Aspose.Cells benchmark for formulas with {0} rows and {1} columns", rowCount, columnCount));

            double startMem = GetMemory();

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            Double[,] values = new Double[rowCount,columnCount];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    values[i,j] = (double)(i + j);
                }
            }
            worksheet.Cells.CreateRange("A1:B" + rowCount).Value=values;

            DateTime start = DateTime.Now;
            Cells cells = worksheet.Cells;
            //for (int c = 2; c < 3; c++)
            //{
            //    string formulastring = String.Format("SUM({0}1, {1}1)", CellsHelper.ColumnIndexToName(c - 2), CellsHelper.ColumnIndexToName(c - 1));
                
            //}
            string formulastring = String.Format("SUM({0}1, {1}1)", CellsHelper.ColumnIndexToName(2 - 2), CellsHelper.ColumnIndexToName(2 - 1));

            cells[0, 2].SetSharedFormula(formulastring, rowCount, 30);
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells set formulas: {0:N3} s", setTime));

            start = DateTime.Now;
            workbook.CalculateFormula();
            end = DateTime.Now;

            calcTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells calculate formulas: {0:N3} s", calcTime));

            start = DateTime.Now;
            workbook.Save("../../output/aspose-saved-formulas.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells save formulas to Excel: {0:N3} s", saveTime));

            double endMem = GetMemory();
            usedMem = (endMem - startMem) ;
            //      Console.WriteLine(String.Format("Aspose.Cells used memory: {0:N3} MB", usedMem));

        }

        public static void TestBigExcelFile(int rowCount,
                                                int columnCount,
                                                ref double openTime,
                                                ref double calcTime,
                                                ref double saveTime,
                                                ref double usedMem)
        {
            Console.WriteLine();
            Console.WriteLine(String.Format("Aspose.Cells benchmark for test-performance.xlsx which is 20.5MB with a lot of values, formulas and styles"));

            double startMem = GetMemory();


            DateTime start = DateTime.Now;
            Workbook workbook = new Workbook("../../files/test-performance.xlsx");
            DateTime end = DateTime.Now;

            openTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells open big Excel file: {0:N3} s", openTime));

            start = DateTime.Now;
            workbook.CalculateFormula();
            end = DateTime.Now;

            calcTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells calculate formulas: {0:N3} s", calcTime));

            start = DateTime.Now;
            workbook.Save("../../output/aspose-saved-test-performance.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(String.Format("Aspose.Cells save formulas to Excel: {0:N3} s", saveTime));

            double endMem = GetMemory();
            usedMem = (endMem - startMem) ;
            Console.WriteLine(String.Format("Aspose.Cells used memory: {0:N3} MB", usedMem));

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
