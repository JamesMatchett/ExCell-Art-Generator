using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExCell_Art
{
    public static class OpenFileExtensionExtension
    {
        /// <summary>
        /// Returns a string array of all extensions of files selected in the OpenFileDialog
        /// </summary>
        /// <param name="O"></param>
        /// <returns>string[]</returns>
        public static string[] Extensions(this OpenFileDialog O)
        {
            
            string[] fileNames = O.FileNames;
            string[] ext = new string[fileNames.GetUpperBound(0)];
            int x = fileNames.GetUpperBound(0);
            if (x == 0)
            {
                ext[0] = fileNames[0].Substring(fileNames[0].LastIndexOf("."), fileNames[0].Length - 1);
            }
            for (int i = 0; i <= fileNames.GetUpperBound(0); i++)
            {
                string t = (fileNames[i].Substring(fileNames[i].LastIndexOf("."), fileNames[i].Length - 1)); 
                //ext[i] = (fileNames[i].Substring(fileNames[i].LastIndexOf("."), fileNames[i].Length - 1));
            }
            return ext;
        }
    }
}
