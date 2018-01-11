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
                A.start();
                A.bw.ProgressChanged += Bw_ProgressChanged;

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

        private void Bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }
    }
}
