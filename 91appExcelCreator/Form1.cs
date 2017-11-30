using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace _91appExcelCreator
{
    public partial class Form1 : Form
    {
        private readonly PictureTheme _pictureTheme = new PictureTheme();

        public Form1()
        {
            InitializeComponent();
        }

        private static void AddWorkSheet(Excel._Workbook workbook, Excel._Application excelApp, string sheetName, bool needCreate)
        {
            excelApp.Worksheets.Add();
            var wSheet = (Excel._Worksheet)workbook.Worksheets[1];
            wSheet.Name = sheetName;
            wSheet.Activate();
            if (needCreate)
            {
                excelApp.Cells[1, 1] = "交期";
                excelApp.Cells[2, 1] = "一般";
                excelApp.Cells[3, 1] = "預購";
                excelApp.Cells[4, 1] = "訂製";
            }
        }

        /// <summary>
        /// 繪製文字的Function
        /// </summary>
        /// <param name="graphics">要被繪製的圖像</param>
        /// <param name="drawString">要繪製上去的文字</param>
        /// <param name="font">繪製上去的文字字型</param>
        /// <param name="color">繪製上去的文字顏色</param>
        /// <param name="positionX">文字的左上角X位置</param>
        /// <param name="positionY">文字的左上角Y位置</param>
        private static void Drawing(Graphics graphics, string drawString, Font font, Color color, int positionX, int positionY)
        {
            graphics.DrawString(drawString, font, new SolidBrush(color), positionX, positionY);
        }

        private static void InitialExcelTitles(Excel._Application excelApp)
        {
            excelApp.Cells[1, 1] = "商品品類";
            excelApp.Cells[1, 2] = "商店類別";
            excelApp.Cells[1, 3] = "商品名稱";
            excelApp.Cells[1, 4] = "數量";
            excelApp.Cells[1, 5] = "建議售價";
            excelApp.Cells[1, 6] = "售價";
            excelApp.Cells[1, 7] = "成本";
            excelApp.Cells[1, 8] = "一次最高購買量";
            excelApp.Cells[1, 9] = "銷售開始日期";
            excelApp.Cells[1, 10] = "銷售結束日期";
            excelApp.Cells[1, 11] = "交期";
            excelApp.Cells[1, 12] = "預定出貨日期";
            excelApp.Cells[1, 13] = "付款完成後幾天出貨";
            excelApp.Cells[1, 14] = "配送方式";
            excelApp.Cells[1, 15] = "付款方式";
            excelApp.Cells[1, 16] = "商品選項";
            excelApp.Cells[1, 17] = "商品選項一";
            excelApp.Cells[1, 18] = "商品選項二";
            excelApp.Cells[1, 19] = "商品料號";
            excelApp.Cells[1, 20] = "商品選項圖檔";
            excelApp.Cells[1, 21] = "商品規格";

            excelApp.Cells[1, 22] = "商品圖檔一";
            excelApp.Cells[1, 23] = "商品圖檔二";
            excelApp.Cells[1, 24] = "商品圖檔三";
            excelApp.Cells[1, 25] = "商品圖檔四";
            excelApp.Cells[1, 26] = "商品圖檔五";
            excelApp.Cells[1, 27] = "商品圖檔六";
            excelApp.Cells[1, 28] = "商品圖檔七";
            excelApp.Cells[1, 29] = "商品圖檔八";
            excelApp.Cells[1, 30] = "商品圖檔九";
            excelApp.Cells[1, 31] = "商品圖檔十";

            excelApp.Cells[1, 32] = "銷售重點";
            excelApp.Cells[1, 33] = "商品特色";
            excelApp.Cells[1, 34] = "詳細說明";
            excelApp.Cells[1, 35] = "商店名稱";
            excelApp.Cells[1, 36] = "SEOTitle";
            excelApp.Cells[1, 37] = "SEOKeyword";
            excelApp.Cells[1, 38] = "SEODescription";
            excelApp.Cells[1, 39] = "溫層類別";
            excelApp.Cells[1, 40] = "商品材積(長x寬x高)";
            excelApp.Cells[1, 41] = "商品重量";
        }

        private static Excel._Worksheet SetFirstWorkSheet(Excel._Workbook workbook, Excel.Application excelApp)
        {
            var firstWorkSheet = (Excel._Worksheet)workbook.Worksheets[1];
            firstWorkSheet.Name = "商品資料";
            firstWorkSheet.Activate();
            InitialExcelTitles(excelApp);
            var range = firstWorkSheet.Range[firstWorkSheet.Cells[1, 1], firstWorkSheet.Cells[1, 41]];
            range.Select();
            range.Font.Color = ColorTranslator.ToOle(Color.White);
            range.Interior.Color = ColorTranslator.ToOle(Color.DimGray);
            return firstWorkSheet;
        }

        private void AutoFitExcelContent(Excel._Worksheet workSheet)
        {
            var range = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[int.Parse(amountOfData.Text), 41]];
            range.Select();
            range.Columns.AutoFit();
        }

        private void CheckAndCreateFolder()
        {
            var pathString = Path.Combine(_pictureTheme.locate);
            Directory.CreateDirectory(pathString);
        }

        /// <summary>
        /// 創建圖片的Function
        /// </summary>
        private void CreateExampleImg()
        {
            var exampleImage = new Bitmap(_pictureTheme.Width, _pictureTheme.Height,
                PixelFormat.Format32bppArgb);
            var graphics = Graphics.FromImage(exampleImage);
            graphics.FillRectangle(new SolidBrush(_pictureTheme.BackgroundColor),
                new Rectangle(0, 0, _pictureTheme.Width, _pictureTheme.Height));
            CreateImg(graphics, 0);
            example.Image = exampleImage;
        }

        /// <summary>
        /// 創建圖片
        /// </summary>
        /// <param name="graphics">要繪製的圖像</param>
        /// <param name="number">文字後的數字</param>
        private void CreateImg(Graphics graphics, int number)
        {
            var stringWidth = (int)graphics.MeasureString(pictureWords.Text + number, _pictureTheme.FontCounter).Width / 2;
            var stringHeight = (int)graphics.MeasureString(pictureWords.Text + number, _pictureTheme.FontCounter).Height / 2;
            if (randomColor.Checked)
            {
                _pictureTheme.BackgroundColor = _pictureTheme.GetRandomColor();
            }
            graphics.Clear(_pictureTheme.BackgroundColor);

            DrawCharacter(graphics, number, stringWidth, stringHeight);
        }

        private void CreateImgBtn_Click(object sender, EventArgs e)
        {
            var bitmap = new Bitmap(_pictureTheme.Width, _pictureTheme.Height,
                PixelFormat.Format32bppArgb);
            var graphics = Graphics.FromImage(bitmap);
            graphics.FillRectangle(new SolidBrush(_pictureTheme.BackgroundColor),
                new Rectangle(0, 0, _pictureTheme.Width, _pictureTheme.Height));

            CheckAndCreateFolder();
            try
            {
                for (int i = 1; i <= int.Parse(amountOfData.Text); i++)
                {
                    CreateImg(graphics, i);
                    bitmap.Save(_pictureTheme.locate + i + ProductImg1.Text, _pictureTheme.ImageFormat);
                }
                _pictureTheme.CreateZip(int.Parse(amountOfData.Text));
                MessageBox.Show(@"圖片建立完成");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
                throw;
            }
            finally
            {
                graphics?.Dispose();
                bitmap?.Dispose();
            }
        }

        private void DoExcel(Excel._Application excelApp, int amountValue)
        {
            for (var i = 2; i < amountValue + 2; i++)
            {
                excelApp.Cells[i, 1] = cbProductCategory.Text;
                excelApp.Cells[i, 2] = StoreClass.Text;
                excelApp.Cells[i, 3] = ProductName.Text + (i - 1);
                excelApp.Cells[i, 4] = Quantity.Text;
                excelApp.Cells[i, 5] = SugestPrice.Text;
                excelApp.Cells[i, 6] = Price.Text;
                excelApp.Cells[i, 7] = Cost.Text;
                excelApp.Cells[i, 8] = HighestBuyQuantity.Text;
                excelApp.Cells[i, 9] = StartDateTime.Text;
                excelApp.Cells[i, 10] = EndDateTime.Text;
                excelApp.Cells[i, 11] = Delivery.Text;
                excelApp.Cells[i, 12] = ExpectedShippingDay.Text;
                excelApp.Cells[i, 13] = AfterPayShippingDay.Text;
                excelApp.Cells[i, 14] = ShippingType.Text;
                excelApp.Cells[i, 15] = PayType.Text;
                excelApp.Cells[i, 16] = ProductOption.Text;
                excelApp.Cells[i, 17] = ProductOption1.Text;
                excelApp.Cells[i, 18] = ProductOption2.Text;
                excelApp.Cells[i, 19] = ProductNumber.Text;
                excelApp.Cells[i, 20] = ProductOptionImg.Text;
                excelApp.Cells[i, 21] = ProductSpec.Text;
                excelApp.Cells[i, 22] = (i - 1) + ProductImg1.Text;
                excelApp.Cells[i, 32] = SalePoint.Text;
                excelApp.Cells[i, 33] = ProductFeature.Text;
                excelApp.Cells[i, 34] = Detail.Text;
                excelApp.Cells[i, 35] = cbStoreName.Text;
                excelApp.Cells[i, 36] = SEOTitle.Text;
                excelApp.Cells[i, 37] = SEOKeyword.Text;
                excelApp.Cells[i, 38] = SEODescription.Text;
                excelApp.Cells[i, 39] = WarmLayerClass.Text;
                excelApp.Cells[i, 40] = Volume.Text;
                excelApp.Cells[i, 41] = Weight.Text;
                excelApp.Cells[1, 1] = Convert.ToString(1314520, 16);
            }
        }

        private void DrawCharacter(Graphics graphics, int number, int stringWidth, int stringHeight)
        {
            var middleWidth = (_pictureTheme.Width / 2) - stringWidth;
            var middleHeight = (_pictureTheme.Height / 2) - stringHeight;

            Drawing(graphics,
                pictureWords.Text.Insert(pictureWords.TextLength, number.ToString()),
                _pictureTheme.FontCounter,
                Color.DarkMagenta,
                middleWidth,
                middleHeight - 150);

            Drawing(graphics,
                pictureWords.Text.Insert(pictureWords.TextLength / 2,
                    number.ToString()),
                _pictureTheme.FontCounter,
                Color.Yellow,
                middleWidth,
                middleHeight);

            Drawing(graphics,
                pictureWords.Text.Insert(0, number.ToString()),
                _pictureTheme.FontCounter,
                Color.SeaGreen,
                middleWidth,
                middleHeight + 150);
        }

        private void fileLocate_TextChanged(object sender, EventArgs e)
        {
            _pictureTheme.locate = fileLocate.Text;
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            amountOfData.Focus();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _pictureTheme.BackgroundColor = Color.Black;
            _pictureTheme.Words = "測試專用";
            _pictureTheme.Width = 400;
            _pictureTheme.Height = 400;
            _pictureTheme.FontCounter = new Font("微軟正黑體", 48);
            _pictureTheme.locate = @"C:\Users\" + Environment.UserName + @"\Documents\Test\";
            _pictureTheme.ImageFormat = ImageFormat.Jpeg;
            fileLocate.Text = _pictureTheme.locate;
            CreateExampleImg();
        }

        private void OutputExcel_Click(object sender, EventArgs e)
        {
            var pathFile = @"C:\Users\" + Environment.UserName + @"\Documents\Test\" + "test.xlsx";

            var excelApp = new Excel.Application
            {
                Visible = true,
                DisplayAlerts = false
            };
            excelApp.Workbooks.Add(Type.Missing);
            var workbook = excelApp.Workbooks[1];
            try
            {
                var worksheet = SetFirstWorkSheet(workbook, excelApp);
                DoExcel(excelApp, int.Parse(amountOfData.Text));
                AutoFitExcelContent(worksheet);
                AddWorkSheet(workbook, excelApp, "Sheet3", false);
                worksheet.Move(workbook.Sheets[1]);
                AddWorkSheet(workbook, excelApp, "資料驗證, 請勿刪除", true);
                worksheet.Move(workbook.Sheets[1]);
                try
                {
                    workbook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MessageBox.Show("儲存文件於 " + Environment.NewLine + pathFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("產生報表時出錯！" + Environment.NewLine + ex.Message);
            }
            workbook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            GC.Collect();
            Console.Read();
        }

        private void PickColor_Click(object sender, EventArgs e)
        {
            var colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                _pictureTheme.BackgroundColor = colorDialog.Color;
                CreateExampleImg();
            }
        }

        private void PickFolder_Click(object sender, EventArgs e)
        {
            var browserDialog = new FolderBrowserDialog();
            if (browserDialog.ShowDialog() == DialogResult.OK)
            {
                fileLocate.Text = browserDialog.SelectedPath + @"\";
                _pictureTheme.locate = browserDialog.SelectedPath + @"\";
            }
        }

        private void pictureWords_TextChanged(object sender, EventArgs e)
        {
            CreateExampleImg();
        }

        private void imgWidth_ValueChanged(object sender, EventArgs e)
        {
            _pictureTheme.Width = (int)imgWidth.Value;
            CreateExampleImg();
        }

        private void imgHeight_ValueChanged(object sender, EventArgs e)
        {

            _pictureTheme.Height = (int)imgHeight.Value;
            CreateExampleImg();

        }
        private void radioButton1_CheckedChanged_1(object sender, EventArgs e)
        {
            _pictureTheme.ImageFormat = radioButton1.Checked ? ImageFormat.Jpeg : ImageFormat.Png;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

            _pictureTheme.ImageFormat = radioButton1.Checked ? ImageFormat.Png : ImageFormat.Jpeg;
        }
    }
}