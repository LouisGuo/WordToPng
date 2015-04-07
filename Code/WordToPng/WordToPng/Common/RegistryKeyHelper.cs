using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPng.Common
{

    class RegistryKeyHelper
    {
        /// <summary>
        /// 注册右键菜单
        /// </summary>
        /// <param name="itemName">右键菜单名称</param>
        /// <param name="assoCreatedProgramFullPath">程序所在路径</param>
        public static void AddFileContextMenuItem(string itemName, string assoCreatedProgramFullPath)
        {
            //注册到所有文件
            RegistryKey shellKey = Registry.ClassesRoot.OpenSubKey(@"*\shell", true);
            //注册到所有目录
            //RegistryKey shellKey = Registry.ClassesRoot.OpenSubKey(@"directory\shell", true);
            //注册到文件夹
            //RegistryKey shellKey = Registry.ClassesRoot.OpenSubKey("directory", true).OpenSubKey("shell", true);
            if (shellKey == null)
            {
                shellKey = Registry.ClassesRoot.CreateSubKey(@"*\shell");
            }
            RegistryKey rightCommondKey = shellKey.OpenSubKey(itemName);
            if (rightCommondKey == null)
            {
                rightCommondKey = shellKey.CreateSubKey(itemName);
            }

            RegistryKey assoCreatedProgramKey = rightCommondKey.CreateSubKey("command");
            assoCreatedProgramKey.SetValue(string.Empty, assoCreatedProgramFullPath);

            assoCreatedProgramKey.Close();
            rightCommondKey.Close();
            shellKey.Close();


            //注册到所有文件
            //RegistryKey shellKey = Registry.ClassesRoot.OpenSubKey(@"*\shell", true);
            //注册到所有目录
             shellKey = Registry.ClassesRoot.OpenSubKey(@"directory\shell", true);
            //注册到文件夹
            //RegistryKey shellKey = Registry.ClassesRoot.OpenSubKey("directory", true).OpenSubKey("shell", true);
            if (shellKey == null)
            {
                shellKey = Registry.ClassesRoot.CreateSubKey(@"*\shell");
            }
            rightCommondKey = shellKey.OpenSubKey(itemName);
            if (rightCommondKey == null)
            {
                rightCommondKey = shellKey.CreateSubKey(itemName);
            }

            assoCreatedProgramKey = rightCommondKey.CreateSubKey("command");
            assoCreatedProgramKey.SetValue(string.Empty, assoCreatedProgramFullPath);

            assoCreatedProgramKey.Close();
            rightCommondKey.Close();
            shellKey.Close();

        }

        /// <summary>
        /// 删除菜单
        /// </summary>
        /// <param name="itemName"></param>
        public static void DeleteContextMenu(string itemName)
        {
            RegistryKey shellKey = Registry.ClassesRoot.OpenSubKey(@"*\shell", true);
                RegistryKey rightCommondKey = shellKey.OpenSubKey(itemName, true);
                if (rightCommondKey != null)
                {
                    rightCommondKey.DeleteSubKeyTree("");
                }


                shellKey = Registry.ClassesRoot.OpenSubKey(@"directory\shell", true);
                rightCommondKey = shellKey.OpenSubKey(itemName, true);
                if (rightCommondKey != null)
                {
                    rightCommondKey.DeleteSubKeyTree("");
                }
        }
    }
}
