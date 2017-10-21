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
            
        }
    }
}
