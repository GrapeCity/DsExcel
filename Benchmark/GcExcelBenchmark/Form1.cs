using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GCExcelBenchMark
{
    public partial class Form1 : Form
    {
        int time_as = 0, time_npoi=0, time_gc=0;
        int row = 100000, col = 30;
        public Form1()
        {
            InitializeComponent();
            grid.CellPainting += grid_CellPainting;
            grid.Rows.Add("常规文档", "设置 (SetTime)", "", "", "");
            grid.Rows.Add("常规文档", "获取 (GetTime)", "", "", "");
            grid.Rows.Add("常规文档", "保存 (SaveTime)", "", "", "");
            grid.Rows.Add("常规文档", "内存占用 (UsedMem)", "", "", "");
            grid.Rows.Add("大型文档", "打开 (OpenTime)", "", "", "");
            grid.Rows.Add("大型文档", "计算 (CalcTime)", "", "", "");
            grid.Rows.Add("大型文档", "保存 (SaveTime)", "", "", "");
            grid.Rows.Add("大型文档", "内存占用 (UsedMem)", "", "", "");
            grid.Rows.Add("公式计算", "设置 (SetTime)", "", "", "");
            grid.Rows.Add("公式计算", "计算 (CalcTime)", "", "", "");
            grid.Rows.Add("公式计算", "保存 (SaveTime)", "", "", "");

            timer1.Start();
        }



        private void timer1_Tick(object sender, EventArgs e)
        {
            if (worker_aspose.IsBusy)
            {
                time_as++;
                timelabel_as.Text = String.Format("共用时{0:F1}秒", time_as / 10.0);
            }
            if (worker_gcexcel.IsBusy)
            {
                time_gc++;
                timelabel_gc.Text = String.Format("共用时{0:F1}秒", time_gc / 10.0);
            }
            if (worker_npoi.IsBusy)
            {
                time_npoi++;
                timelabel_npoi.Text = String.Format("共用时{0:F1}秒", time_npoi / 10.0);
            }

        }

        private void worker_npoi_DoWork(object sender, DoWorkEventArgs e)
        {
            GC.Collect();
            double setTime = 0, getTime = 0, saveTime = 0, usedMem = 0;
            NPOIBenchmark.TestSetRangeValues_Double(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            grid.Rows[0].Cells[3].Value = string.Format("{0:F3}s", setTime);
            grid.Rows[1].Cells[3].Value = string.Format("{0:F3}s", getTime);
            grid.Rows[2].Cells[3].Value = string.Format("{0:F3}s", saveTime);
            grid.Rows[3].Cells[3].Value = string.Format("{0}MB", usedMem);
            GC.Collect();
            NPOIBenchmark.TestBigExcelFile(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            grid.Rows[4].Cells[3].Value = string.Format("{0:F3}s", setTime);
            grid.Rows[5].Cells[3].Value = string.Format("{0:F3}s", getTime);
            grid.Rows[6].Cells[3].Value = string.Format("{0:F3}s", saveTime);
            grid.Rows[7].Cells[3].Value = string.Format("{0}MB", usedMem);
            GC.Collect();
            NPOIBenchmark .TestSetRangeFormulas(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            grid.Rows[8]. Cells[3].Value = string.Format("{0:F3}s", setTime);
            grid.Rows[9]. Cells[3].Value = string.Format("{0:F3}s", getTime);
            grid.Rows[10].Cells[3].Value = string.Format("{0:F3}s", saveTime);
            
        }
        private void worker_npoi_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            run_npoi.Enabled = true;
        }
        private void run_npoi_Click(object sender, EventArgs e)
        {
            time_npoi = 0;
            run_npoi.Enabled = false;
            worker_npoi.RunWorkerAsync();
        }

        private void worker_gcexcel_DoWork(object sender, DoWorkEventArgs e)
        {
            
            double setTime = 0, getTime = 0, saveTime = 0, usedMem = 0;
            GC.Collect();
            GcExcelBenchmark.TestSetRangeValues_Double(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            grid.Rows[0].Cells[2].Value = string.Format("{0:F3}s", setTime);
            grid.Rows[1].Cells[2].Value = string.Format("{0:F3}s", getTime);
            grid.Rows[2].Cells[2].Value = string.Format("{0:F3}s", saveTime);
            grid.Rows[3].Cells[2].Value = string.Format("{0}MB" , usedMem);
           
            GC.Collect();
            GcExcelBenchmark.TestBigExcelFile(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            grid.Rows[4].Cells[2].Value = string.Format("{0:F3}s", setTime);
            grid.Rows[5].Cells[2].Value = string.Format("{0:F3}s", getTime);
            grid.Rows[6].Cells[2].Value = string.Format("{0:F3}s", saveTime);
            grid.Rows[7].Cells[2].Value = string.Format("{0}MB", usedMem);
            
            GC.Collect();
            GcExcelBenchmark.TestSetRangeFormulas(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            grid.Rows[8].Cells[2].Value = string.Format("{0:F3}s", setTime);
            grid.Rows[9].Cells[2].Value = string.Format("{0:F3}s", getTime);
            grid.Rows[10].Cells[2].Value = string.Format("{0:F3}s", saveTime);

        }
        private void worker_gcexcel_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            run_gcexcel.Enabled = true;
        }
        private void run_gcexcel_Click(object sender, EventArgs e)
        {
            time_gc = 0;
            run_gcexcel.Enabled = false;
            worker_gcexcel.RunWorkerAsync();
        }

        private void worker_aspose_DoWork(object sender, DoWorkEventArgs e)
        {
            double setTime = 0, getTime = 0, saveTime = 0, usedMem = 0;
            GC.Collect();
            AsposeBenchmark.TestSetRangeValues_Double(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            grid.Rows[0].Cells[4].Value = string.Format("{0:F3}s", setTime);
            grid.Rows[1].Cells[4].Value = string.Format("{0:F3}s", getTime);
            grid.Rows[2].Cells[4].Value = string.Format("{0:F3}s", saveTime);
            grid.Rows[3].Cells[4].Value = string.Format("{0}MB", usedMem);
            GC.Collect();
            AsposeBenchmark.TestBigExcelFile(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            grid.Rows[4].Cells[4].Value = string.Format("{0:F3}s", setTime);
            grid.Rows[5].Cells[4].Value = string.Format("{0:F3}s", getTime);
            grid.Rows[6].Cells[4].Value = string.Format("{0:F3}s", saveTime);
            grid.Rows[7].Cells[4].Value = string.Format("{0}MB", usedMem);
            GC.Collect();
            AsposeBenchmark.TestSetRangeFormulas(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            grid.Rows[8].Cells[4].Value = string.Format("{0:F3}s", setTime);
            grid.Rows[9].Cells[4].Value = string.Format("{0:F3}s", getTime);
            grid.Rows[10].Cells[4].Value = string.Format("{0:F3}s", saveTime);

            
        }
        private void worker_aspose_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            run_aspose.Enabled = true;
        }
        private void run_aspose_Click(object sender, EventArgs e)
        {
            time_as = 0;
            run_aspose.Enabled = false;
            worker_aspose.RunWorkerAsync();
        }


        private void grid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // 对第1列相同单元格进行合并
            if (e.ColumnIndex == 0 && e.RowIndex != -1)
            {
                using
                    (
                    Brush gridBrush = new SolidBrush(this.grid.GridColor),
                    backColorBrush = new SolidBrush(e.CellStyle.BackColor)
                    )
                {
                    using (Pen gridLinePen = new Pen(gridBrush))
                    {
                        // 清除单元格
                        e.Graphics.FillRectangle(backColorBrush, e.CellBounds);
                        // 画 Grid 边线（仅画单元格的底边线和右边线）
                        //   如果下一行和当前行的数据不同，则在当前的单元格画一条底边线
                        if ((e.RowIndex < grid.Rows.Count - 1 &&
                        grid.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value.ToString() !=
                        e.Value.ToString() || e.RowIndex == grid.Rows.Count - 1))
                        {
                            e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);
                        }

                        // 画右边线
                        e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1,
                            e.CellBounds.Top, e.CellBounds.Right - 1,
                            e.CellBounds.Bottom);
                        // 画（填写）单元格内容，相同的内容的单元格只填写第一个
                        if (e.Value != null)
                        {
                            if (e.RowIndex > 0 &&
                            grid.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString() ==
                            e.Value.ToString())
                            { }
                            else
                            {
                                e.Graphics.DrawString((String)e.Value, e.CellStyle.Font,
                                    Brushes.Black, e.CellBounds.X + 2,
                                    e.CellBounds.Y + 5, StringFormat.GenericDefault);
                            }
                        }
                        //e.Handled=true;这一句非常重要，必须加上，要不所画的内容就被后面的Painting事件刷新不见了！！！
                        e.Handled = true;
                    }

                }
            }
        }


    }
}
