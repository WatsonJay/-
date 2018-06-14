using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 易之星车贷数据处理
{
    public partial class ProcessForm : Form
    {
        private BackgroundWorker backgroundWorker1; //ProcessForm 窗体事件(进度条窗体)

        public ProcessForm(BackgroundWorker backgroundWorker1,int value)
        {
            InitializeComponent();
            progressBar1.Maximum = value+1;
            this.backgroundWorker1 = backgroundWorker1;
            this.backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
        }

        void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Close();//执行完之后，直接关闭页面
        }

        void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Value = e.ProgressPercentage;
        }

        private void cancle_Click(object sender, EventArgs e)
        {
            this.backgroundWorker1.CancelAsync();
            this.cancle.Enabled = false;
            this.Close();
        }
    }
}
