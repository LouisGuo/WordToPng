using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WordToPng.Common;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordToPng
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private ImgHelper imgH = new ImgHelper();

        private List<string> fullPathList = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            this.Topmost = true;
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            
            if(App.fullPathList.Count>0)
            {
                StartRotate();
                this.fullPathList.AddRange(App.fullPathList);
                t = new Thread(doWork);
                t.Start();  
            }
        }

        private void ThisToImg(string from)
        {
            int lastIndex = from.LastIndexOf("\\");
            string path = from.Substring(0, lastIndex + 1);
            string fileName = from.Substring(lastIndex + 1, from.Length - lastIndex - 6);

            imgH.ToImg(from, path + fileName + ".png");
        }

        private void StartRotate()
        {
            RotateTransform rtf = new RotateTransform();
            rtf.CenterX = this.Width / 2;
            rtf.CenterY = this.Height / 2;

            this.pic.RenderTransform = rtf;

            DoubleAnimation dbAscending = new DoubleAnimation(0, -360, new Duration(TimeSpan.FromSeconds(1)));
            dbAscending.RepeatBehavior = RepeatBehavior.Forever;

            this.pic.RenderTransform.BeginAnimation(RotateTransform.AngleProperty, dbAscending);
        }

        private void EndRotate()
        {
            if (this.pic.RenderTransform != null)
            {
                this.pic.RenderTransform = null;
            }
        }



        delegate void MyDelegate(int now,int sum);  
        Thread t;  
  
  
        // 要长时间做的工作  
        void doWork()  
        {  
            MyDelegate d = new MyDelegate(setValue);
            for (int i = 0; i < fullPathList.Count;i++ )
            {
                this.Dispatcher.Invoke(d, i,fullPathList.Count);
                ThisToImg(fullPathList[i]);
            }
            this.Dispatcher.Invoke(d, fullPathList.Count, fullPathList.Count);
            fullPathList.Clear();
        }  
  
        // 更新用户界面  
        void setValue(int now,int sum)      
        {  
            label1.Content = now+"/"+sum; 
 
            if(now==sum)
            {
                EndRotate();
                if(App.fullPathList.Count>0)
                {
                    this.Close();
                }
            }


        }  
  







        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            base.DragMove();
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            this.backPic.Visibility = Visibility.Hidden;
            if (fullPathList.Count != 0)
            {
                MessageBox.Show("Please wait Until the ongoing work is done");
                return;
            }
            StartRotate();
            //添加要转换的文件列表
            string fullPath = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

            for (int i = 1; fullPath != null; i++)
            {
                fullPathList.AddRange(DocFile.GetFileList(fullPath));
                try
                {
                    fullPath = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(i).ToString();
                }
                catch
                {
                    fullPath = null;
                }

            }


            t = new Thread(doWork);
            t.Start();  
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            this.backPic.Visibility = Visibility.Visible;
            this.label1.Content = "";
            if(e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Link;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key==Key.Delete)
            {
                //RegistryKeyHelper.DeleteContextMenu(App.itemName);
                this.Close();
            }
        }


       

    }
}
