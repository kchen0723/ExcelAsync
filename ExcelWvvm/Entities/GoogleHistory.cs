using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelWvvm.Interfaces;
using System.ComponentModel;

namespace ExcelWvvm.Entities
{
    public class GoogleHistory : IGoogleHistory
    {
        public event EventHandler OnRetrievedData;

        public string SecurityId { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        private BackgroundWorker asyncWorker = null;

        public void ExecuteAsync()
        {
            this.asyncWorker = new BackgroundWorker();
            this.asyncWorker.WorkerReportsProgress = true;
            this.asyncWorker.WorkerSupportsCancellation = true;
            this.asyncWorker.DoWork += AsyncWorker_DoWork;
            this.asyncWorker.RunWorkerCompleted += AsyncWorker_RunWorkerCompleted;
            if (this.asyncWorker.IsBusy == false)
            {
                this.asyncWorker.RunWorkerAsync();
            }
        }

        public void CancelExecute()
        {
            if (this.asyncWorker != null)
            {
                this.asyncWorker.CancelAsync();
            }
        }

        private void AsyncWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (this.OnRetrievedData != null)
            {
                this.OnRetrievedData(this, EventArgs.Empty);
            }
        }

        private void AsyncWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            object[,]  result = GoogleHistoryManager.GoogleHistory(this);
            if (this.asyncWorker != null && this.asyncWorker.CancellationPending == false)
            {
                if (ExcelHandler.WriteToRangeHandler != null)
                {
                    ExcelHandler.WriteToRangeHandler(result);
                }
            }
        }
    }
}
