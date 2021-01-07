
namespace GCExcelBenchMark
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.run_npoi = new System.Windows.Forms.Button();
            this.grid = new System.Windows.Forms.DataGridView();
            this.timelabel_as = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.worker_npoi = new System.ComponentModel.BackgroundWorker();
            this.worker_gcexcel = new System.ComponentModel.BackgroundWorker();
            this.run_gcexcel = new System.Windows.Forms.Button();
            this.run_aspose = new System.Windows.Forms.Button();
            this.worker_aspose = new System.ComponentModel.BackgroundWorker();
            this.timelabel_gc = new System.Windows.Forms.Label();
            this.timelabel_npoi = new System.Windows.Forms.Label();
            this.Aspose = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Npoi = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Gcexcel = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Action = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Group = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            this.SuspendLayout();
            // 
            // run_npoi
            // 
            this.run_npoi.Location = new System.Drawing.Point(668, 79);
            this.run_npoi.Name = "run_npoi";
            this.run_npoi.Size = new System.Drawing.Size(81, 32);
            this.run_npoi.TabIndex = 0;
            this.run_npoi.Text = "NPOI测试";
            this.run_npoi.UseVisualStyleBackColor = true;
            this.run_npoi.Click += new System.EventHandler(this.run_npoi_Click);
            // 
            // grid
            // 
            this.grid.AllowUserToAddRows = false;
            this.grid.AllowUserToDeleteRows = false;
            this.grid.AllowUserToResizeColumns = false;
            this.grid.AllowUserToResizeRows = false;
            this.grid.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Teal;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("黑体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.grid.ColumnHeadersHeight = 40;
            this.grid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Group,
            this.Action,
            this.Gcexcel,
            this.Npoi,
            this.Aspose});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微软雅黑", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.grid.DefaultCellStyle = dataGridViewCellStyle2;
            this.grid.Location = new System.Drawing.Point(12, 12);
            this.grid.Name = "grid";
            this.grid.ReadOnly = true;
            this.grid.RowHeadersVisible = false;
            this.grid.RowTemplate.Height = 30;
            this.grid.Size = new System.Drawing.Size(645, 500);
            this.grid.TabIndex = 1;
            // 
            // timelabel_as
            // 
            this.timelabel_as.AutoSize = true;
            this.timelabel_as.Location = new System.Drawing.Point(754, 145);
            this.timelabel_as.Name = "timelabel_as";
            this.timelabel_as.Size = new System.Drawing.Size(35, 12);
            this.timelabel_as.TabIndex = 2;
            this.timelabel_as.Text = "0.000";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // worker_npoi
            // 
            this.worker_npoi.DoWork += new System.ComponentModel.DoWorkEventHandler(this.worker_npoi_DoWork);
            this.worker_npoi.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.worker_npoi_RunWorkerCompleted);
            // 
            // worker_gcexcel
            // 
            this.worker_gcexcel.DoWork += new System.ComponentModel.DoWorkEventHandler(this.worker_gcexcel_DoWork);
            this.worker_gcexcel.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.worker_gcexcel_RunWorkerCompleted);
            // 
            // run_gcexcel
            // 
            this.run_gcexcel.Location = new System.Drawing.Point(668, 22);
            this.run_gcexcel.Name = "run_gcexcel";
            this.run_gcexcel.Size = new System.Drawing.Size(81, 32);
            this.run_gcexcel.TabIndex = 3;
            this.run_gcexcel.Text = "GcExcel测试";
            this.run_gcexcel.UseVisualStyleBackColor = true;
            this.run_gcexcel.Click += new System.EventHandler(this.run_gcexcel_Click);
            // 
            // run_aspose
            // 
            this.run_aspose.Location = new System.Drawing.Point(668, 135);
            this.run_aspose.Name = "run_aspose";
            this.run_aspose.Size = new System.Drawing.Size(81, 32);
            this.run_aspose.TabIndex = 4;
            this.run_aspose.Text = "Aspose测试";
            this.run_aspose.UseVisualStyleBackColor = true;
            this.run_aspose.Click += new System.EventHandler(this.run_aspose_Click);
            // 
            // worker_aspose
            // 
            this.worker_aspose.DoWork += new System.ComponentModel.DoWorkEventHandler(this.worker_aspose_DoWork);
            this.worker_aspose.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.worker_aspose_RunWorkerCompleted);
            // 
            // timelabel_gc
            // 
            this.timelabel_gc.AutoSize = true;
            this.timelabel_gc.Location = new System.Drawing.Point(754, 32);
            this.timelabel_gc.Name = "timelabel_gc";
            this.timelabel_gc.Size = new System.Drawing.Size(35, 12);
            this.timelabel_gc.TabIndex = 5;
            this.timelabel_gc.Text = "0.000";
            // 
            // timelabel_npoi
            // 
            this.timelabel_npoi.AutoSize = true;
            this.timelabel_npoi.Location = new System.Drawing.Point(754, 89);
            this.timelabel_npoi.Name = "timelabel_npoi";
            this.timelabel_npoi.Size = new System.Drawing.Size(35, 12);
            this.timelabel_npoi.TabIndex = 6;
            this.timelabel_npoi.Text = "0.000";
            // 
            // Aspose
            // 
            this.Aspose.HeaderText = "第三方组件";
            this.Aspose.Name = "Aspose";
            this.Aspose.ReadOnly = true;
            this.Aspose.Width = 120;
            // 
            // Npoi
            // 
            this.Npoi.HeaderText = "NPOI";
            this.Npoi.Name = "Npoi";
            this.Npoi.ReadOnly = true;
            this.Npoi.Width = 120;
            // 
            // Gcexcel
            // 
            this.Gcexcel.HeaderText = "GcExcel";
            this.Gcexcel.Name = "Gcexcel";
            this.Gcexcel.ReadOnly = true;
            this.Gcexcel.Width = 120;
            // 
            // Action
            // 
            this.Action.Frozen = true;
            this.Action.HeaderText = "操作";
            this.Action.Name = "Action";
            this.Action.ReadOnly = true;
            this.Action.Width = 150;
            // 
            // Group
            // 
            this.Group.Frozen = true;
            this.Group.HeaderText = "";
            this.Group.Name = "Group";
            this.Group.ReadOnly = true;
            this.Group.Width = 130;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(834, 521);
            this.Controls.Add(this.timelabel_npoi);
            this.Controls.Add(this.timelabel_gc);
            this.Controls.Add(this.run_aspose);
            this.Controls.Add(this.run_gcexcel);
            this.Controls.Add(this.timelabel_as);
            this.Controls.Add(this.grid);
            this.Controls.Add(this.run_npoi);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button run_npoi;
        private System.Windows.Forms.DataGridView grid;
        private System.Windows.Forms.Label timelabel_as;
        private System.Windows.Forms.Timer timer1;
        private System.ComponentModel.BackgroundWorker worker_npoi;
        private System.ComponentModel.BackgroundWorker worker_gcexcel;
        private System.Windows.Forms.Button run_gcexcel;
        private System.Windows.Forms.Button run_aspose;
        private System.ComponentModel.BackgroundWorker worker_aspose;
        private System.Windows.Forms.Label timelabel_gc;
        private System.Windows.Forms.Label timelabel_npoi;
        private System.Windows.Forms.DataGridViewTextBoxColumn Group;
        private System.Windows.Forms.DataGridViewTextBoxColumn Action;
        private System.Windows.Forms.DataGridViewTextBoxColumn Gcexcel;
        private System.Windows.Forms.DataGridViewTextBoxColumn Npoi;
        private System.Windows.Forms.DataGridViewTextBoxColumn Aspose;
    }
}

