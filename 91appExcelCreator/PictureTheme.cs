using Ionic.Zip;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;

namespace _91appExcelCreator
{
    internal class PictureTheme
    {
        public Color BackgroundColor { get; set; }
        public string Words { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public Font FontCounter { get; set; }
        public ImageFormat ImageFormat { get; set; }
        public Color GetRandomColor()
        {
            var random = new Random();
            var r = random.Next(0, 255);
            Thread.Sleep(5);
            var g = random.Next(0, 255);
            Thread.Sleep(5);
            var b = random.Next(0, 255);
            return Color.FromArgb(r, g, b);
        }

        public void CreateZip(int count)
        {
            using (var zip = new ZipFile("Test"))
            {
                for (var i = 1; i <= count; i++)
                {
                    zip.AddFile(this.locate + i + ".jpg", string.Empty);
                }
                zip.Save(this.locate + DateTime.Now.ToString("yyyy-MM-dd,HH-mm-ss") + @"_test.zip");
            }
        }

        public string locate { get; set; }
    }
}