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
        public static string itemName = "WordToPng";
        public string itemPath = AppDomain.CurrentDomain.DomainManager.EntryAssembly.Location + " %1";//当前exe文件夹路径
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            if(e.Args.Length>0)
            {
                for (int i = 0; i < e.Args.Length; i++)
                {
                    fullPathList.AddRange(DocFile.GetFileList(e.Args[i]));
                }
            }
            else
            {
                //RegistryKeyHelper.DeleteContextMenu(itemName);
                //RegistryKeyHelper.AddFileContextMenuItem(itemName, itemPath);
            }
        }

        private void Application_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.Message);
        }
    }
}
