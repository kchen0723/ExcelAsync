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
        public Action<object, object> OnRetrievedDataHandler { get; set; }

        public string SecurityId { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string InstanceId { get; set; }
        public string RangeName { get; set; }
        public string SheetId { get; set; }

        private BackgroundWorker asyncWorker = null;

        public void ExecuteAsync()
        {
            this.asyncWorker = new BackgroundWorker();
            this.asyncWorker.WorkerReportsProgress = true;
            this.asyncWorker.WorkerSupportsCancellation = true;
            this.asyncWorker.DoWork += AsyncWorker_DoWork;
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

        private void AsyncWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            object[,]  result = GoogleHistoryManager.GoogleHistory(this);
            if (this.OnRetrievedDataHandler != null && this.asyncWorker != null && this.asyncWorker.CancellationPending == false )
            {
                this.OnRetrievedDataHandler(this, result);
            }
        }
    }
}
