using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GCExcelBenchMark
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
  
            //也可以在此直接执行测试
            //int row = 100000, col = 30;
            //double setTime = 0, getTime = 0, saveTime = 0, usedMem = 0;
            //GcExcelBenchmark.TestSetRangeValues_Double(row, col,ref setTime, ref getTime, ref setTime, ref usedMem);

        }
    }
}
