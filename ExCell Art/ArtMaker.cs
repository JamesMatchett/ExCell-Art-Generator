using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExCell_Art
{
    class ArtMaker
    {
      public BackgroundWorker bw = new BackgroundWorker();
        string ImagePath;
        string OutputPath;
        int PercComplete;

        public Excel.Application xlApp = new Excel.Application();

        public ArtMaker(string _ImagePath, string _OutputPath)
        {
            ImagePath = _ImagePath;
            OutputPath = _OutputPath;
        }

        public void start()
        {
            this.bw = new BackgroundWorker();
            this.bw.WorkerReportsProgress = true;
            this.bw.DoWork += bw_makeArt;
            this.bw.RunWorkerAsync();
            this.bw.WorkerSupportsCancellation = true;
        }

        public void stop()
        {
            bw.CancelAsync();
        }

        public void bw_makeArt(object sender, EventArgs e)
        {
            //read input image
            Bitmap bm = new Bitmap(ImagePath);
            //create a new excel document
           
            
            
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
            Excel._Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            xlWorksheet.Unprotect();
            xlWorksheet.StandardWidth = 20 / 7.25;
            
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //i = across, j = up, image coordinates start from bottom left corner whereas excel starts from top left
            for(int j = 0; j < bm.Height - 1; j++) 
            {
                for(int i = 0; i < bm.Width - 1; i++)
                {
                    
                    xlRange.Cells[j+1, i+1].Interior.Color = System.Drawing.ColorTranslator.ToOle(bm.GetPixel(i,j));
                }
                PercComplete = Convert.ToInt32(Math.Floor(((decimal)j / (bm.Height-1)) * 100));
                
                Debug.WriteLine("{0} out of {1}",j*bm.Width,bm.Width*bm.Height);
                
                Debug.WriteLine("amount is {0}", PercComplete);
                bw.ReportProgress(PercComplete);
            }
            

            xlWorkbook.SaveAs(OutputPath);
            xlApp.Quit();
            bw.ReportProgress(100);


        }

    }
}
