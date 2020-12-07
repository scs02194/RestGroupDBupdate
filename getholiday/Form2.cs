using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace getholiday
{
    public partial class Form2 : Form
    {
        BackgroundWorker worker = new BackgroundWorker();
        public Form2()
        {
            InitializeComponent();
        }
        void plusdot()
        {
            while (true)
            {
                if(worker.CancellationPending == true)
                {
                    break;
                }
                else
                {
                    if (label1.Text.Length == 14)
                    {
                        label1.Text = "Loading";
                    }
                    else
                    {
                        label1.Text += ".";
                        Thread.Sleep(1000);
                    }
                }
            }
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += (sender1, args) => plusdot();
             worker.RunWorkerAsync();

        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            worker.CancelAsync();
        }
    }
}
