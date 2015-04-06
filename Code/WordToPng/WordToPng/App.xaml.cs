using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using WordToPng.Common;

namespace WordToPng
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        public static List<string> fullPathList = new List<string>();
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            for (int i = 0; i < e.Args.Length; i++)
            {
                fullPathList.AddRange(DocFile.GetFileList(e.Args[i]));
            }
        }
    }
}
