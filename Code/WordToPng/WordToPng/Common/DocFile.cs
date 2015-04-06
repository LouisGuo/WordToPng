using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPng.Common
{
    class DocFile
    {
        public static List<String> GetFileList(string fullPath)
        {
            List<string> fullPathList = new List<string>();
            if (File.Exists(fullPath) && (fullPath.ToLower().EndsWith(".docx") || fullPath.ToLower().EndsWith(".doc")) && !fullPath.Contains("~$"))
            {
                fullPathList.Add(fullPath);
            }
            else if (Directory.Exists(fullPath))
            {
                string[] files = Directory.GetFiles(fullPath, "*.doc", SearchOption.AllDirectories);
                if (files != null)
                {
                    for (int j = 0; j < files.Length; j++)
                    {
                        if (!files[j].Contains("~$"))
                            fullPathList.Add(files[j]);
                    }
                }
            }
            return fullPathList;
        }
    }
}
