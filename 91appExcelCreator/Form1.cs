using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Windows.Forms;

namespace _91appExcelCreator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void ProductCategory_Enter(object sender, EventArgs e)
        {
            var defaultValue = "SEO0118";
            PlaceHolderSetting(ProductCategory, defaultValue, false);
        }

        private void ProductCategory_Leave(object sender, EventArgs e)
        {

            var defaultValue = "SEO0118";
            PlaceHolderSetting(ProductCategory, defaultValue, true);
        }

        private void StoreClass_Enter(object sender, EventArgs e)
        {

            var defaultValue = "巴拉巴拉";
            PlaceHolderSetting(StoreClass, defaultValue, false);
        }

        private void StoreClass_Leave(object sender, EventArgs e)
        {
            var defaultValue = "巴拉巴拉";
            PlaceHolderSetting(StoreClass, defaultValue, true);
        }

        private static void PlaceHolderSetting(Control textBox, string defaultValue, bool leave)
        {
            if (textBox.Text.Equals(defaultValue) && !leave)
            {
                textBox.Text = string.Empty;
                return;
            }
            textBox.Text = defaultValue;
        }

        private void output_Click(object sender, EventArgs e)
        {
            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名
            string pathFile = @"D:\test";

            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();

            // 讓Excel文件可見
            excelApp.Visible = true;

            // 停用警告訊息
            excelApp.DisplayAlerts = false;

            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);

            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];

            // 設定活頁簿焦點
            wBook.Activate();

            try
            {
                // 引用第一個工作表
                wSheet = (Excel._Worksheet)wBook.Worksheets[1];

                // 命名工作表的名稱
                wSheet.Name = "工作表測試";

                // 設定工作表焦點
                wSheet.Activate();

                //excelApp.Cells[1, 1] = "Excel測試";

                // 設定第1列資料
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
                excelApp.Cells[1, 23] = "銷售重點";
                excelApp.Cells[1, 24] = "商品特色";
                excelApp.Cells[1, 25] = "詳細說明";
                excelApp.Cells[1, 26] = "商店名稱";
                excelApp.Cells[1, 27] = "SEOTitle";
                excelApp.Cells[1, 28] = "SEOKeyword";
                excelApp.Cells[1, 29] = "SEODescription";
                excelApp.Cells[1, 30] = "溫層類別";
                excelApp.Cells[1, 31] = "商品材積";
                excelApp.Cells[1, 32] = "商品重量";


                // 設定第1列顏色
                wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 32]];
                wRange.Select();
                wRange.Font.Color = ColorTranslator.ToOle(Color.White);
                wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);

                // 設定第2列資料
                excelApp.Cells[2, 1] = "AA";
                excelApp.Cells[2, 2] = "10";

                // 自動調整欄寬
                wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[5, 2]];
                wRange.Select();
                wRange.Columns.AutoFit();

                try
                {
                    //另存活頁簿
                    wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
            }

            //關閉活頁簿
            wBook.Close(false, Type.Missing, Type.Missing);

            //關閉Excel
            excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            Console.Read();
        }
    }
}
