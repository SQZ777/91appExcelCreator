﻿using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
namespace _91appExcelCreator
{
    public partial class Form1 : Form
    {
        private readonly PictureTheme _pictureTheme = new PictureTheme();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string pathFile = _pictureTheme.locate + DateTime.Now.ToString("yyyy-MM-dd,HH-mm-ss") + ".xlsx";
            var excelApp = new Excel.Application
            {
                Visible = true,
                DisplayAlerts = false
            };
            excelApp.Workbooks.Add(Type.Missing);
            Excel._Workbook wBook = excelApp.Workbooks[1];

            try
            {
                var worksheet = SetFirstWorkSheet(wBook, excelApp);
                DoExcel(excelApp, int.Parse(amountOfData.Text));
                AutoFitExcelContent(worksheet);
                AddWorkSheet(wBook, excelApp, "Sheet3", false);
                worksheet.Move(wBook.Sheets[1]);
                AddWorkSheet(wBook, excelApp, "資料驗證, 請勿刪除", true);
                worksheet.Move(wBook.Sheets[1]);
                try
                {
                    wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MessageBox.Show("儲存文件於 " + Environment.NewLine + pathFile);
                    Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
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
            wBook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            GC.Collect();
            Console.Read();
        }

        private static Excel._Worksheet SetFirstWorkSheet(Excel._Workbook wBook, Excel.Application excelApp)
        {
            var wSheet = (Excel._Worksheet)wBook.Worksheets[1];
            wSheet.Name = "商品資料";
            wSheet.Activate();
            InitialExcelTitles(excelApp);
            var wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 41]];
            wRange.Select();
            wRange.Font.Color = ColorTranslator.ToOle(Color.White);
            wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);
            return wSheet;
        }

        private static void AddWorkSheet(Excel._Workbook wBook, Excel._Application excelApp, string sheetName, bool needCreate)
        {
            excelApp.Worksheets.Add();
            var wSheet = (Excel._Worksheet)wBook.Worksheets[1];
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

        private void AutoFitExcelContent(Excel._Worksheet workSheet)
        {
            var wRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[int.Parse(amountOfData.Text), 41]];
            wRange.Select();
            wRange.Columns.AutoFit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _pictureTheme.BackgroundColor = Color.Black;
            _pictureTheme.Words = "測試專用";
            _pictureTheme.Width = 400;
            _pictureTheme.Height = 400;
            _pictureTheme.FontCounter = new Font("微軟正黑體", 48);
            _pictureTheme.locate = @"C:\Users\" + Environment.UserName + @"\Documents\Test\";
            fileLocate.Text = _pictureTheme.locate;
            CreateExampleImg();
        }

        private void CreateExampleImg()
        {
            var exampleImage = new Bitmap(_pictureTheme.Width, _pictureTheme.Height, PixelFormat.Format32bppArgb);
            var g = Graphics.FromImage(exampleImage);
            g.FillRectangle(new SolidBrush(_pictureTheme.BackgroundColor),
                new Rectangle(0, 0, _pictureTheme.Width, _pictureTheme.Height));
            CreateImg(g, 0);
            example.Image = exampleImage;
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            amountOfData.Focus();
        }

        private void CreateImgBtn_Click(object sender, EventArgs e)
        {
            var newBitmap = new Bitmap(_pictureTheme.Width, _pictureTheme.Height, PixelFormat.Format32bppArgb);
            var g = Graphics.FromImage(newBitmap);
            g.FillRectangle(new SolidBrush(_pictureTheme.BackgroundColor),
                new Rectangle(0, 0, _pictureTheme.Width, _pictureTheme.Height));
            try
            {
                CheckAndCreateFolder();
                try
                {
                    for (int i = 1; i <= int.Parse(amountOfData.Text); i++)
                    {
                        CreateImg(g, i);
                        newBitmap.Save(_pictureTheme.locate + i + ProductImg1.Text, ImageFormat.Jpeg);
                    }
                    _pictureTheme.CreateZip(int.Parse(amountOfData.Text));
                    MessageBox.Show(@"圖片建立完成");
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.ToString());
                    throw;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                g?.Dispose();
                newBitmap?.Dispose();
            }
        }

        private void CreateImg(Graphics g, int i)
        {
            var stringWidth = (int)g.MeasureString(pictureWords.Text + i, _pictureTheme.FontCounter).Width / 2;
            var stringHeight = (int)g.MeasureString(pictureWords.Text + i, _pictureTheme.FontCounter).Height / 2;
            if (randomColor.Checked)
            {
                _pictureTheme.BackgroundColor = _pictureTheme.GetRandomColor();
            }
            g.Clear(_pictureTheme.BackgroundColor);
            var middleWidth = (_pictureTheme.Width / 2) - stringWidth;
            var middleHeight = (_pictureTheme.Height / 2) - stringHeight;
            Drawing(g, pictureWords.Text.Insert(pictureWords.TextLength, i.ToString()), _pictureTheme.FontCounter, Color.DarkMagenta, middleWidth, middleHeight - 150);
            Drawing(g, pictureWords.Text.Insert(pictureWords.TextLength / 2, i.ToString()), _pictureTheme.FontCounter, Color.Yellow, middleWidth, middleHeight);
            Drawing(g, pictureWords.Text.Insert(0, i.ToString()), _pictureTheme.FontCounter, Color.SeaGreen, middleWidth, middleHeight + 150);
        }

        private static void Drawing(Graphics g, string drawString, Font font, Color color, int positionX, int positionY)
        {
            g.DrawString(drawString, font, new SolidBrush(color), positionX, positionY);
        }

        private void PickColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                _pictureTheme.BackgroundColor = colorDialog1.Color;
                CreateExampleImg();
            }
        }

        private void CheckAndCreateFolder()
        {
            var pathString = Path.Combine(_pictureTheme.locate);
            Directory.CreateDirectory(pathString);
        }

        private void pictureWords_TextChanged(object sender, EventArgs e)
        {
            CreateExampleImg();
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
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var browserDialog = new FolderBrowserDialog();
            if (browserDialog.ShowDialog() == DialogResult.OK)
            {
                fileLocate.Text = browserDialog.SelectedPath + @"\";
                _pictureTheme.locate = browserDialog.SelectedPath + @"\";
            }
        }
    }
}