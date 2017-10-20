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
            if (ProductCategory.Text.Equals("SEO0118"))
            {
                ProductCategory.Text = string.Empty;
            }
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
            if (StoreClass.Text.Equals(string.Empty))
            {
                StoreClass.Text = @"巴拉巴拉";
            }
        }
    }
}
