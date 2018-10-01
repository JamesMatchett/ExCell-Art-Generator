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
   public class ArtMaker
    {
      public BackgroundWorker bw = new BackgroundWorker();
      
        string ImagePath;
        string OutputPath;
        int PercComplete;

        bool Quitting = false;

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
            Quitting = true;
            bw.CancelAsync();
            bw.Dispose();
            
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

            ColorSetter cs = new ColorSetter(xlRange);

            //assign it here rather than in the for loop to avoid multithreading calamities
            var Width = bm.Width;
            var Height = bm.Height;

            decimal totalPixels = (bm.Width - 1) * (bm.Height - 1);
            decimal pixelCounter = 0;
            
            //i = across, j = up, image coordinates start from bottom left corner whereas excel starts from top left
            Parallel.For(0, Height - 1, (j,loopState) =>
            {
                
                var bmClone_ = bm.Clone();
                Bitmap bmClone = ((Bitmap)(bmClone_));
                for (int i = 0; i < Width - 1; i++)
                {
                    //make sure a cancel hasn't been requested
                    if (!Quitting)
                    {
                        try
                        {
                            cs.setColor(i, j, ref bmClone);
                           
                            pixelCounter++;
                            
                            
                        } catch
                        {
                            
                        }
                    }
                    else
                    {
                        //stops all threads from executing for clean exit when cancel is called
                        loopState.Stop();
                        break;
                    }
                }

                if (Quitting)
                {
                    loopState.Stop();
                }

                string msg = pixelCounter + "/" + totalPixels;
                Debug.WriteLine(msg);
                decimal progress = ((pixelCounter / totalPixels)*100);
                int progressInt = Convert.ToInt32(progress);
                bw.ReportProgress(progressInt);
            });

           
            if (!Quitting)
            {
                xlWorkbook.SaveAs(OutputPath);
                bw.ReportProgress(101);
            }

            xlRange = cs.Finish();
            xlWorkbook.Close(0);
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

        }

       

        private class ColorSetter
        {
            Excel.Range _xlRange;
            public ColorSetter(Excel.Range xlRange)
            {
                _xlRange = xlRange;
            }

            public void setColor(int i, int j, ref Bitmap bmClone)
            {
                var x = _xlRange.Cells[j + 1, i + 1];
                x.Interior.Color = ColorTranslator.ToOle(bmClone.GetPixel(i, j));
            }

            public Excel.Range Finish()
            {
                return _xlRange;
            }
        }
    }
}
