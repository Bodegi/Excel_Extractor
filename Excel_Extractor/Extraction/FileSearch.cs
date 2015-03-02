using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Extraction
{
    public class FileSearch
    {
        public FileSearch(string dir, string output)
        {
            Excel final = new Excel();
            List<string> lstFilesFound = new List<string>();
            bool isExcel = false;
            foreach(string d in Directory.GetDirectories(dir))
            {
                foreach(string f in Directory.GetFiles(d))
                {
                    isExcel = extension(f);
                    if(isExcel == true)
                    {

                    }
                }
            }
        }

        private bool extension(string file)
        {
            string ext = "";
            int count = 0;
            count = file.IndexOf(".");
            ext = file.Substring(count).ToUpper();
            if (ext == ".XLS")
            {
                return true;
            }
            if (ext == ".XLSX")
            {
                return true;
            }
            return false;
        }
    }
}
