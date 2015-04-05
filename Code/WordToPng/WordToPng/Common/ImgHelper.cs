using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;


namespace WordToPng.Common
{
    public class ImgHelper
    {
        public int cutTimes = 1;
        public int cutTimesCount = 0;

        private List<StartEnd> CutRange(List<StartEnd> fromList, List<StartEnd> cutList)
        {
            List<StartEnd> resultList = new List<StartEnd>();
            for (int i = 0; i < fromList.Count; i++)
            {
                bool SAVE = true;
                for (int j = 0; j < cutList.Count; j++)
                {
                    if (fromList[i].Start >= cutList[j].Start && fromList[i].End <= cutList[j].End)
                    {
                        SAVE = false;
                        if(fromList[i].Start == cutList[j].Start)
                        {
                            resultList.Add(cutList[j]);
                        }
                    }
                }
                if (SAVE)
                    resultList.Add(fromList[i]);
            }
            return resultList;
        }




        public void ToImg(string From,string To)
        {
            Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Word.Document doc = null;
            object unknow = Type.Missing;
            app.Visible = false;
            object file = From;
            doc = app.Documents.Open(ref file,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow);


            cutTimes = 1;
            cutTimesCount = 0;

            const int MAX_height = 3000;
            Range range = doc.Range();
            Range range2 = doc.Range();


            double zoom = 0.33;
            const int imgWidth = 1188;
            

            Image imgTemp = Metafile.FromStream(new MemoryStream(range.EnhMetaFileBits));

            if (MAX_height < imgTemp.Height)
            {
                Paragraphs paragraphs = range.Paragraphs;
                Tables tables = range.Tables;

                List<StartEnd> paragraphList = new List<StartEnd>();
                List<StartEnd> tableList = new List<StartEnd>();

                for(int i=0;i<paragraphs.Count;i++)
                {
                    Paragraph paragraph=paragraphs[i+1];
                    StartEnd startEnd = new StartEnd();
                    startEnd.Start = paragraph.Range.Start;
                    startEnd.End = paragraph.Range.End;

                    paragraphList.Add(startEnd);
                }
                for(int i=0;i<tables.Count;i++)
                {
                    Table table=tables[i+1];
                    StartEnd startEnd = new StartEnd();
                    startEnd.Start = table.Range.Start;
                    startEnd.End = table.Range.End;

                    tableList.Add(startEnd);
                }

                List<StartEnd> resultList = CutRange(paragraphList,tableList);

                List<StartEnd> finalImgRangeList = new List<StartEnd>();
                for (int i = 0; i < resultList.Count;i++ )
                {
                    StartEnd startendImg = new StartEnd();
                    startendImg.Start = resultList[i].Start;
                    startendImg.End = resultList[i].End;
                    for(int j=i;j<resultList.Count;j++)
                    {
                        range2.SetRange((int)resultList[i].Start,(int)resultList[j].End);
                        Image img = Metafile.FromStream(new MemoryStream(range2.EnhMetaFileBits));
                        if(img.Height<MAX_height)
                        {
                            startendImg.End = resultList[j].End;
                            if (j == resultList.Count - 1)
                                i = j;
                        }
                        else
                        {
                            if(i==j)
                            {
                                MessageBox.Show("请确定没有超过一页的段落或表格");
                                //doc.Application.ActiveWindow.ScrollIntoView(range2);

                            }
                            i = j-1;
                            break;
                        }
                    }
                    finalImgRangeList.Add(startendImg);
                }
                cutTimes = finalImgRangeList.Count;


                //timeResult = timeResult + "Time2:" + DateTime.Now.ToString() + "\n";


                int allImgHeight = 0;
                int allImgWidth = 0;
                for (int i = 0; i < finalImgRangeList.Count; i++)
                {
                    range2.SetRange((int)finalImgRangeList[i].Start,(int)finalImgRangeList[i].End);
                    Image img = Metafile.FromStream(new MemoryStream(range2.EnhMetaFileBits));
                    if (img.Width > allImgWidth)
                        allImgWidth = img.Width;

                    allImgHeight += img.Height;
                }
                //zoom = (double)imgWidth / (double)allImgWidth;
                System.Drawing.Bitmap bmp = new Bitmap(imgWidth, (int)(allImgHeight * zoom));
                
                System.Drawing.Graphics gx = System.Drawing.Graphics.FromImage(bmp); // 创建Graphics对象 
                gx.InterpolationMode = InterpolationMode.HighQualityBicubic;
                // 指定高质量、低速度呈现。  
                gx.SmoothingMode = SmoothingMode.HighQuality;
                gx.CompositingQuality = CompositingQuality.HighQuality;

                gx.CompositingMode = CompositingMode.SourceOver;
                gx.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                int startPosition = 0;
                double oldZoom = zoom;
                for (int i = 0; i < finalImgRangeList.Count; i++)
                {
                    range2.SetRange((int)finalImgRangeList[i].Start, (int)finalImgRangeList[i].End);
                    Image img = Metafile.FromStream(new MemoryStream(range2.EnhMetaFileBits));

                    if ((double)imgWidth / (double)img.Width < zoom)
                        zoom = (double)imgWidth / (double)img.Width;

                    gx.FillRectangle(new SolidBrush(System.Drawing.Color.White), 0, startPosition, (int)(img.Width * zoom), (int)(img.Height * zoom));
                    gx.DrawImage(img, new System.Drawing.Rectangle(0, startPosition, (int)(img.Width * zoom), (int)(img.Height * zoom)));

                    startPosition += (int)(img.Height * zoom);
                    zoom = oldZoom;

                    cutTimesCount = i + 1;
                }

                //bmp = KiSharpen(bmp,(float)0.3);
                bmp.Save(To, System.Drawing.Imaging.ImageFormat.Png);
                
            }
            else
            {
                //zoom = (double)imgWidth / (double)imgTemp.Width;
                System.Drawing.Bitmap bmp = new Bitmap(imgWidth, (int)(imgTemp.Height * zoom));
                System.Drawing.Graphics gx = System.Drawing.Graphics.FromImage(bmp); // 创建Graphics对象 
                gx.InterpolationMode = InterpolationMode.HighQualityBicubic;
                // 指定高质量、低速度呈现。  
                gx.SmoothingMode = SmoothingMode.HighQuality;
                gx.CompositingQuality = CompositingQuality.HighQuality;

                gx.CompositingMode = CompositingMode.SourceOver;
                gx.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                gx.FillRectangle(new SolidBrush(System.Drawing.Color.White), 0, 0, (int)(imgTemp.Width * zoom), (int)(imgTemp.Height * zoom));
                gx.DrawImage(imgTemp, new System.Drawing.Rectangle(0, 0, (int)(imgTemp.Width * zoom), (int)(imgTemp.Height * zoom)));
                //imgTemp.Save(Globals.ThisAddIn.exerciseJsonPath + paperName + "\\" + imgName, System.Drawing.Imaging.ImageFormat.Png);
                bmp.Save(To, System.Drawing.Imaging.ImageFormat.Png);
                
            }
            //timeResult = timeResult + "Time3:" + DateTime.Now.ToString() + "\n";
            doc.Close();
            
        }

    }
}
