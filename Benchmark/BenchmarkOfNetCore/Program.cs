using GrapeCity.Documents.Excel;
using System;
using System.Diagnostics;
using System.IO;

namespace Benchmark
{
    class Program
    {
        private const string InputFilePath = @"Files\Input";
        private const string OutFilePath = @"Files\Output";

        static void Main(string[] args)
        {
            var files = Directory.GetFiles(InputFilePath);
            if (files == null || files.Length == 0)
            {
                string fullInputFilePath = Path.Combine(Directory.GetCurrentDirectory(), InputFilePath);
                Console.WriteLine("Please put a file in \"" + fullInputFilePath + "\"");
                Console.WriteLine("Press any key to quit...");
                Console.ReadLine();
                return;
            }
            if (files.Length > 1)
            {
                Console.WriteLine("There are more than one file, but we only test one.");
            }
            var inputFile = files[0];
            var fileName = Path.GetFileName(inputFile);
            Console.WriteLine("Benchmark for SpreadServices");
            Console.WriteLine();

            Console.WriteLine("FileName: \"" + fileName + "\"");
            Console.WriteLine();

            Workbook workbook = new Workbook();
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            workbook.Open(inputFile);
            stopwatch.Stop();
            Console.WriteLine("Open time:      " + (stopwatch.ElapsedMilliseconds / 1000d).ToString("0.###") + "s");

            long memorySize = GC.GetTotalMemory(true);
            Console.WriteLine("Used Memory:            " + (memorySize / 1024d / 1024d).ToString("##.###") + "M");
            Console.WriteLine();
            stopwatch.Restart();
            workbook.Dirty();
            workbook.Calculate();
            stopwatch.Stop();
            Console.WriteLine("Calclate time   " + (stopwatch.ElapsedMilliseconds / 1000d).ToString("0.###") + "s");

            memorySize = GC.GetTotalMemory(true);
            Console.WriteLine("Used Memory:            " + (memorySize / 1024d / 1024d).ToString("##.###") + "M");
            Console.WriteLine();

            if (!Directory.Exists(OutFilePath))
            {
                Directory.CreateDirectory(OutFilePath);
            }

            stopwatch.Restart();
            workbook.Save(Path.Combine(OutFilePath, fileName), null, new SaveOptions() { IsCompactMode = true });
            stopwatch.Stop();
            Console.WriteLine("Save time       " + (stopwatch.ElapsedMilliseconds / 1000d).ToString("0.###") + "s");

            memorySize = GC.GetTotalMemory(true);
            Console.WriteLine("Used Memory:            " + (memorySize / 1024d / 1024d).ToString("##.###") + "M");
            Console.WriteLine();

			// Prevent the GC collect the workbook before we show the memory size.
            workbook.Worksheets[0].Cells[0, 0].Value = 1;
			
            Console.WriteLine("Press any key to quit...");
            Console.ReadLine();
        }
    }
}
