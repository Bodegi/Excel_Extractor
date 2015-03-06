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
            List<string> visited = new List<string>();
            traversal(dir, visited);
        }

        public static void traversal(string dir, List<string> visited)
        {
            bool isExcel = false;
            foreach(string f in Directory.GetFiles(dir))
            {
                isExcel = extension(f);
                if(isExcel == true)
                {
                    Excel.TabCheck(f);
                }
            }
            visited.Add(dir);
            foreach(string d in Directory.GetDirectories(dir))
            {
                if(!visited.Contains(d))
                {
                    traversal(d, visited);
                }
            }
        }

        public static bool extension(string file)
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
