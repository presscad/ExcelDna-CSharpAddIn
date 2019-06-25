using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpAddIn
{
    public partial class frmSpotLight : Form
    {
        public frmSpotLight()
        {
            InitializeComponent();

            btnRow.BackColor = Color.FromArgb(RibbonController.row_clr & 0x0000ff, (RibbonController.row_clr & 0x00ff00) >> 8, (RibbonController.row_clr & 0xff0000) >> 16);
            btnCol.BackColor = Color.FromArgb(RibbonController.col_clr & 0x0000ff, (RibbonController.col_clr & 0x00ff00) >> 8, (RibbonController.col_clr & 0xff0000) >> 16);
            ndTransparent.Value = Convert.ToDecimal(RibbonController.clr_transparent) * 100;
        }

        private void btnRow_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.AllowFullOpen = true;
            colorDialog.FullOpen = true;
            colorDialog.ShowHelp = true;
            colorDialog.Color = Color.Black;//初始化颜色
            colorDialog.ShowDialog();
            Color clr = colorDialog.Color;
            btnRow.BackColor = clr;
            RibbonController.row_clr = (int)(((uint)clr.B << 16) | (ushort)(((ushort)clr.G << 8) | clr.R));
        }

        private void btnCol_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.AllowFullOpen = true;
            colorDialog.FullOpen = true;
            colorDialog.ShowHelp = true;
            colorDialog.Color = Color.Black;//初始化颜色
            colorDialog.ShowDialog();
            Color clr = colorDialog.Color;
            btnCol.BackColor = clr;
            RibbonController.row_clr = (int)(((uint)clr.B << 16) | (ushort)(((ushort)clr.G << 8) | clr.R));
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            RibbonController.clr_transparent = Decimal.ToSingle(ndTransparent.Value)/100.0f;
        }
    }
}
