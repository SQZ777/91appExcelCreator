using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace _91appExcelCreator
{
    class PictureTheme
    {
        public Color BackgroundColor { get; set; }
        public string Words { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public Font FontCounter { get; set; }
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
    }
}
