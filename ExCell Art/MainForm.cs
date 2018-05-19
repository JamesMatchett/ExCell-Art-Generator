using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExCell_Art.OpenFileExtensionExtension;

namespace ExCell_Art
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        public List<ArtMaker> BWList = new List<ArtMaker>();

        public string ImagePath;
        public string OutputPath;

        private void ImageButton_Click(object sender, EventArgs e)
        {
            
            openFileDialog.Title = "ExCell Art generator";
            openFileDialog.FileName = "";
            openFileDialog.Filter = "Image Files(*.BMP; *.JPG; *.GIF; *.PNG;)| *.BMP; *.JPG; *.GIF; *.PNG;";
            openFileDialog.CheckFileExists = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.Multiselect = true;
            openFileDialog.InitialDirectory = Environment.GetFolderPath(System.Environment.SpecialFolder.MyPictures);
            openFileDialog.FileOk += OpenFileDialog_FileOk;
            openFileDialog.ShowDialog();
            
        }

        private void OpenFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            MessageBox.Show("File chosen");
            InputName.Text = openFileDialog.FileName;
            ImagePath = openFileDialog.FileName;
        }

        private void pathButton_Click(object sender, EventArgs e)
        {
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(System.Environment.SpecialFolder.MyPictures);
            saveFileDialog.CheckPathExists = true;
            saveFileDialog.Filter = "Excel |*.xlsx";
            saveFileDialog.FileOk += SaveFileDialog_FileOk;
            saveFileDialog.AddExtension = true;
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.ShowDialog();
        }

       private void SaveFileDialog_FileOk(object sender, EventArgs e)
        {
            OutputPathLabel.Text = saveFileDialog.FileName;
            OutputPath = saveFileDialog.FileName;
        }

        private void Start_Click(object sender, EventArgs e)
        {
            //check that both an input image and output path have been chosen
            if(!(ImagePath == null) && !(OutputPath == null))
            {
                ArtMaker A = new ArtMaker(ImagePath, OutputPath);
                Cancel.Enabled = true;
                A.start();
                A.bw.ProgressChanged += Bw_ProgressChanged;
                BWList.Add(A);
                Start.Enabled = false;
               
            } else
            {
                if(ImagePath == null)
                {
                    MessageBox.Show("Please choose a valid input image");
                }
                else
                {
                    MessageBox.Show("Please choose a valid output path");
                }
            }

        }

        //for estimated time to completion counter
        DateTime LastPercentageTime = DateTime.Now;
        decimal LastPercentageValue = 0;
        List<decimal> AverageTimeFor1Percent = new List<decimal>();

        private void Bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //if not initiliased, or percentage has not increased since last execution
            if(LastPercentageValue == 0 || e.ProgressPercentage == 0 || LastPercentageValue == e.ProgressPercentage)
            {
                LastPercentageTime = DateTime.Now;
                LastPercentageValue = e.ProgressPercentage;
            }
            else
            {
                decimal PercentageDifference = e.ProgressPercentage - LastPercentageValue;
                TimeSpan TimeDifference = DateTime.Now - LastPercentageTime;

                //time for 1% = (1/percentageDifference * Time Difference)
                decimal TimeFor1Percent = (1 / PercentageDifference) * Convert.ToDecimal(TimeDifference.TotalSeconds);
                AverageTimeFor1Percent.Add(TimeFor1Percent);
                
                if(AverageTimeFor1Percent.Count > 5)
                {
                    AverageTimeFor1Percent.Remove(AverageTimeFor1Percent.First());
                }

                //now multiply time for 1% by the percentage remianing e.g. 45% done would be 55% percent to go
                decimal SecondsRemaining = (AverageTimeFor1Percent.Average() * (100 - e.ProgressPercentage));
                SecondsRemaining = Math.Floor(SecondsRemaining);

                if (SecondsRemaining > 0)
                {
                    //update labels and refresh values
                    TimeRemainingLabel.Text = ("Time Remaining: " + SecondsRemaining + " seconds");
                    LastPercentageTime = DateTime.Now;
                    LastPercentageValue = e.ProgressPercentage;
                }
            }


            int percentage = Convert.ToInt32(e.ProgressPercentage);

            //101 signifies finished, stop multithread sending multiple "finished" signal
            if (percentage == 101)
            {
                TimeRemainingLabel.Text = ("Time Remaining: " + 0 + " seconds");
                progressBar.Value = 100;
                MessageBox.Show("Finished!");
                foreach (ArtMaker A in BWList)
                {
                    A.stop();
                }
                progressBar.Value = 0;
                Start.Enabled = true;
            }
            else
            {
                progressBar.Value = percentage;
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Cancel.Enabled = false;
           foreach(ArtMaker A in BWList)
            {
                A.stop();
            }
           //empty list of all cancelled bw's
            BWList = new List<ArtMaker>();
            MessageBox.Show("Cancelled!");
            progressBar.Value = 0;
            Start.Enabled = true;
        }

        private void TimeRemainingLabel_Click(object sender, EventArgs e)
        {

        }
    }
}
