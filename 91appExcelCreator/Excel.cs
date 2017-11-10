namespace _91appExcelCreator
{
    public class Excel
    {
        private Form1 _form1;

        public Excel(Form1 form1)
        {
            _form1 = form1;
        }

        private void DoExcel(Excel._Application excelApp, int amountValue)
        {
            for (var i = 2; i < amountValue + 2; i++)
            {
                excelApp.Cells[i, 1] = _form1.ProductCategory.Text;
                excelApp.Cells[i, 2] = _form1.StoreClass.Text;
                excelApp.Cells[i, 3] = _form1.ProductName.Text + (i - 1);
                excelApp.Cells[i, 4] = _form1.Quantity.Text;
                excelApp.Cells[i, 5] = _form1.SugestPrice.Text;
                excelApp.Cells[i, 6] = _form1.Price.Text;
                excelApp.Cells[i, 7] = _form1.Cost.Text;
                excelApp.Cells[i, 8] = _form1.HighestBuyQuantity.Text;
                excelApp.Cells[i, 9] = _form1.StartDateTime.Text;
                excelApp.Cells[i, 10] = _form1.EndDateTime.Text;
                excelApp.Cells[i, 11] = _form1.Delivery.Text;
                excelApp.Cells[i, 12] = _form1.ExpectedShippingDay.Text;
                excelApp.Cells[i, 13] = _form1.AfterPayShippingDay.Text;
                excelApp.Cells[i, 14] = _form1.ShippingType.Text;
                excelApp.Cells[i, 15] = _form1.PayType.Text;
                excelApp.Cells[i, 16] = _form1.ProductOption.Text;
                excelApp.Cells[i, 17] = _form1.ProductOption1.Text;
                excelApp.Cells[i, 18] = _form1.ProductOption2.Text;
                excelApp.Cells[i, 19] = _form1.ProductNumber.Text;
                excelApp.Cells[i, 20] = _form1.ProductOptionImg.Text;
                excelApp.Cells[i, 21] = _form1.ProductSpec.Text;
                excelApp.Cells[i, 22] = (i - 1) + _form1.ProductImg1.Text;
                excelApp.Cells[i, 32] = _form1.SalePoint.Text;
                excelApp.Cells[i, 33] = _form1.ProductFeature.Text;
                excelApp.Cells[i, 34] = _form1.Detail.Text;
                excelApp.Cells[i, 35] = _form1.StoreName.Text;
                excelApp.Cells[i, 36] = _form1.SEOTitle.Text;
                excelApp.Cells[i, 37] = _form1.SEOKeyword.Text;
                excelApp.Cells[i, 38] = _form1.SEODescription.Text;
                excelApp.Cells[i, 39] = _form1.WarmLayerClass.Text;
                excelApp.Cells[i, 40] = _form1.Volume.Text;
                excelApp.Cells[i, 41] = _form1.Weight.Text;
            }
        }
    }
}