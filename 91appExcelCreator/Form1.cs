using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            var defaultValue = "巴拉巴拉";
            placeHolderSetting(ProductCategory, defaultValue, false);
        }

        private void placeHolderSetting(TextBox sender, string defaultValue, bool leave)
        {
            if (sender.Text.Equals(defaultValue) && !leave)
            {
                sender.Text = string.Empty;
                return;
            }
            sender.Text = defaultValue;
        }

        private void StoreClass_Enter(object sender, EventArgs e)
        {
            if (StoreClass.Text.Equals("巴拉巴拉"))
            {
                StoreClass.Text = string.Empty;
            }
        }


        private void ProductCategory_Leave(object sender, EventArgs e)
        {
            if (ProductCategory.Text.Equals(string.Empty))
            {
                ProductCategory.Text = @"SEO0118";
            }
        }

        private void StoreClass_Leave(object sender, EventArgs e)
        {
            var defaultValue = "SEO0118";
            placeHolderSetting(StoreClass, defaultValue, true);
        }
    }
}
