using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CSharpAddIn
{
    public partial class frmReg : Form
    {
        public frmReg()
        {
            InitializeComponent();
        }

        static DotNet.Utilities.SoftReg softReg = new DotNet.Utilities.SoftReg();
        static DotNet.Utilities.CheckReg ckReg = new DotNet.Utilities.CheckReg();
 
        private void btnCode_Click(object sender, EventArgs e)
        {
            txtCode.Text = softReg.GetMachineNum();
        }

        private void btnReg_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtReg.Text == softReg.GetRegisterNum(txtCode.Text))
                {
                    MessageBox.Show("ExcelDna 注册成功！重启Excel后生效！", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RegistryKey retkey = Registry.CurrentUser.OpenSubKey("Software", true).CreateSubKey("ExcelDna").CreateSubKey("Register.INI").CreateSubKey(txtReg.Text);
                    retkey.SetValue("UserName", "Rsoft");
                    this.Close();
                }
                else
                {
                    MessageBox.Show("注册码错误！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtReg.SelectAll();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void FrmReg_Load(object sender, EventArgs e)
        {
            if (ckReg.GetIsReg())
            {
                txtCode.ReadOnly = true;
                txtReg.ReadOnly = true;
                this.Text = "ExcelDna 已注册";
                MessageBox.Show("ExcelDna 已注册！", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                this.Text = "ExcelDna 未注册";

                int count = 0;
                if (ckReg.GetUseInfo(ref count))
                {
                    int k = 5 - count;
                    MessageBox.Show("ExcelDna 未注册！试用次数还剩下" + k.ToString() + "次！", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void txtReg_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            txtReg.Text = softReg.GetRegisterNum(txtCode.Text);
        }
    }
}
