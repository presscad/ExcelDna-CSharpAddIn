using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Core = Microsoft.Office.Core;
using Vbe = Microsoft.Vbe.Interop;
using ExcelDna.Integration;
using System.Drawing;
using System.Data.SqlClient;
using ExcelDna.Integration.CustomUI;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using RDotNet;
using Log4Net;
using ExcelDna.Logging;
using ZXing;


namespace CSharpAddIn
{
    [ComVisible(true)]
    public class CTPControl : UserControl
    {
        private GroupBox groupBox1;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private TextBox txtCount;
        private TextBox txtMean;
        private TextBox txtSd;
        private TextBox txtSeed;
        private Button bnRNorm;
        private GroupBox groupBox2;
        private DateTimePicker dtPicker1;
        private ComboBox cbDtfmt1;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label label8;
        private Label label9;
        private Label label10;
        private TextBox txtCount_t;
        private TextBox txtMean_t;
        private TextBox txtSd_t;
        private TextBox txtSeed_t;
        private Button bnRNorm_t;
        private GroupBox groupBox3;
        private DateTimePicker dtPicker2;
        private ComboBox cbDtfmt2;
        private Label label11;
        private Label label12;
        private Button bnDateTime;



        public CTPControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtCount = new System.Windows.Forms.TextBox();
            this.txtMean = new System.Windows.Forms.TextBox();
            this.txtSd = new System.Windows.Forms.TextBox();
            this.txtSeed = new System.Windows.Forms.TextBox();
            this.bnRNorm = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dtPicker1 = new System.Windows.Forms.DateTimePicker();
            this.cbDtfmt1 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.txtCount_t = new System.Windows.Forms.TextBox();
            this.txtMean_t = new System.Windows.Forms.TextBox();
            this.txtSd_t = new System.Windows.Forms.TextBox();
            this.txtSeed_t = new System.Windows.Forms.TextBox();
            this.bnRNorm_t = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dtPicker2 = new System.Windows.Forms.DateTimePicker();
            this.cbDtfmt2 = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.bnDateTime = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtCount);
            this.groupBox1.Controls.Add(this.txtMean);
            this.groupBox1.Controls.Add(this.txtSd);
            this.groupBox1.Controls.Add(this.txtSeed);
            this.groupBox1.Controls.Add(this.bnRNorm);
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(180, 155);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "R.NET 随机数";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "个  数";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "平均值";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 75);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "均方差";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 101);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 8;
            this.label4.Text = "种  子";
            // 
            // txtCount
            // 
            this.txtCount.Location = new System.Drawing.Point(65, 19);
            this.txtCount.Name = "txtCount";
            this.txtCount.Size = new System.Drawing.Size(100, 21);
            this.txtCount.TabIndex = 2;
            this.txtCount.Text = "10";
            // 
            // txtMean
            // 
            this.txtMean.Location = new System.Drawing.Point(65, 45);
            this.txtMean.Name = "txtMean";
            this.txtMean.Size = new System.Drawing.Size(100, 21);
            this.txtMean.TabIndex = 3;
            this.txtMean.Text = "1";
            // 
            // txtSd
            // 
            this.txtSd.Location = new System.Drawing.Point(65, 71);
            this.txtSd.Name = "txtSd";
            this.txtSd.Size = new System.Drawing.Size(100, 21);
            this.txtSd.TabIndex = 4;
            this.txtSd.Text = "0.5";
            // 
            // txtSeed
            // 
            this.txtSeed.Location = new System.Drawing.Point(65, 97);
            this.txtSeed.Name = "txtSeed";
            this.txtSeed.Size = new System.Drawing.Size(100, 21);
            this.txtSeed.TabIndex = 7;
            this.txtSeed.Text = "123";
            // 
            // bnRNorm
            // 
            this.bnRNorm.Location = new System.Drawing.Point(43, 124);
            this.bnRNorm.Name = "bnRNorm";
            this.bnRNorm.Size = new System.Drawing.Size(93, 23);
            this.bnRNorm.TabIndex = 1;
            this.bnRNorm.Text = "生成随机数";
            this.bnRNorm.UseVisualStyleBackColor = true;
            this.bnRNorm.Click += new System.EventHandler(this.btnRNorm_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dtPicker1);
            this.groupBox2.Controls.Add(this.cbDtfmt1);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.txtCount_t);
            this.groupBox2.Controls.Add(this.txtMean_t);
            this.groupBox2.Controls.Add(this.txtSd_t);
            this.groupBox2.Controls.Add(this.txtSeed_t);
            this.groupBox2.Controls.Add(this.bnRNorm_t);
            this.groupBox2.Location = new System.Drawing.Point(3, 170);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(180, 217);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "R.NET 随机时间";
            // 
            // dtPicker1
            // 
            this.dtPicker1.CustomFormat = "HH:mm:ss";
            this.dtPicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtPicker1.Location = new System.Drawing.Point(65, 131);
            this.dtPicker1.Name = "dtPicker1";
            this.dtPicker1.ShowUpDown = true;
            this.dtPicker1.Size = new System.Drawing.Size(100, 21);
            this.dtPicker1.TabIndex = 11;
            // 
            // cbDtfmt1
            // 
            this.cbDtfmt1.FormattingEnabled = true;
            this.cbDtfmt1.Items.AddRange(new object[] {
            "yyyy/MM/dd HH:mm:ss",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy/MM/dd HH:mm",
            "yyyy-MM-dd HH:mm",
            "MM/dd HH:mm:ss",
            "MM-dd HH:mm:ss",
            "MM/dd HH:mm",
            "MM-dd HH:mm",
            "yyyy/MM/dd",
            "yyyy-MM-dd",
            "MM/dd",
            "MM-dd",
            "HH:mm:ss",
            "HH:mm"});
            this.cbDtfmt1.Location = new System.Drawing.Point(65, 157);
            this.cbDtfmt1.Name = "cbDtfmt1";
            this.cbDtfmt1.Size = new System.Drawing.Size(100, 20);
            this.cbDtfmt1.TabIndex = 13;
            this.cbDtfmt1.Text = "HH:mm";
            this.cbDtfmt1.SelectedIndexChanged += new System.EventHandler(this.cbDtfmt1_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(17, 23);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 12);
            this.label5.TabIndex = 0;
            this.label5.Text = "个  数";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(17, 49);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(41, 12);
            this.label6.TabIndex = 5;
            this.label6.Text = "平均值";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(17, 75);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(41, 12);
            this.label7.TabIndex = 6;
            this.label7.Text = "均方差";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(17, 101);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(41, 12);
            this.label8.TabIndex = 8;
            this.label8.Text = "种  子";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(5, 134);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(53, 12);
            this.label9.TabIndex = 10;
            this.label9.Text = "基准时间";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(5, 161);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(53, 12);
            this.label10.TabIndex = 12;
            this.label10.Text = "时间格式";
            // 
            // txtCount_t
            // 
            this.txtCount_t.Location = new System.Drawing.Point(65, 19);
            this.txtCount_t.Name = "txtCount_t";
            this.txtCount_t.Size = new System.Drawing.Size(100, 21);
            this.txtCount_t.TabIndex = 2;
            this.txtCount_t.Text = "10";
            // 
            // txtMean_t
            // 
            this.txtMean_t.Location = new System.Drawing.Point(65, 45);
            this.txtMean_t.Name = "txtMean_t";
            this.txtMean_t.Size = new System.Drawing.Size(100, 21);
            this.txtMean_t.TabIndex = 3;
            this.txtMean_t.Text = "1";
            // 
            // txtSd_t
            // 
            this.txtSd_t.Location = new System.Drawing.Point(65, 71);
            this.txtSd_t.Name = "txtSd_t";
            this.txtSd_t.Size = new System.Drawing.Size(100, 21);
            this.txtSd_t.TabIndex = 4;
            this.txtSd_t.Text = "0.5";
            // 
            // txtSeed_t
            // 
            this.txtSeed_t.Location = new System.Drawing.Point(65, 97);
            this.txtSeed_t.Name = "txtSeed_t";
            this.txtSeed_t.Size = new System.Drawing.Size(100, 21);
            this.txtSeed_t.TabIndex = 7;
            this.txtSeed_t.Text = "123";
            // 
            // bnRNorm_t
            // 
            this.bnRNorm_t.Location = new System.Drawing.Point(43, 185);
            this.bnRNorm_t.Name = "bnRNorm_t";
            this.bnRNorm_t.Size = new System.Drawing.Size(93, 23);
            this.bnRNorm_t.TabIndex = 1;
            this.bnRNorm_t.Text = "生成随机时间";
            this.bnRNorm_t.UseVisualStyleBackColor = true;
            this.bnRNorm_t.Click += new System.EventHandler(this.btnRNorm_t_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dtPicker2);
            this.groupBox3.Controls.Add(this.cbDtfmt2);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.bnDateTime);
            this.groupBox3.Location = new System.Drawing.Point(6, 393);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(180, 107);
            this.groupBox3.TabIndex = 14;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "时间界面选择";
            // 
            // dtPicker2
            // 
            this.dtPicker2.CustomFormat = "yyyy-MM-dd";
            this.dtPicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtPicker2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.dtPicker2.Location = new System.Drawing.Point(65, 20);
            this.dtPicker2.Name = "dtPicker2";
            this.dtPicker2.Size = new System.Drawing.Size(100, 21);
            this.dtPicker2.TabIndex = 11;
            // 
            // cbDtfmt2
            // 
            this.cbDtfmt2.FormattingEnabled = true;
            this.cbDtfmt2.Items.AddRange(new object[] {
            "yyyy/MM/dd HH:mm:ss",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy/MM/dd HH:mm",
            "yyyy-MM-dd HH:mm",
            "MM/dd HH:mm:ss",
            "MM-dd HH:mm:ss",
            "MM/dd HH:mm",
            "MM-dd HH:mm",
            "yyyy/MM/dd",
            "yyyy-MM-dd",
            "MM/dd",
            "MM-dd",
            "HH:mm:ss",
            "HH:mm"});
            this.cbDtfmt2.Location = new System.Drawing.Point(65, 48);
            this.cbDtfmt2.Name = "cbDtfmt2";
            this.cbDtfmt2.Size = new System.Drawing.Size(100, 20);
            this.cbDtfmt2.TabIndex = 13;
            this.cbDtfmt2.Text = "yyyy-MM-dd";
            this.cbDtfmt2.SelectedIndexChanged += new System.EventHandler(this.cbDtfmt2_SelectedIndexChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(5, 23);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(53, 12);
            this.label11.TabIndex = 10;
            this.label11.Text = "时间选择";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(5, 52);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(53, 12);
            this.label12.TabIndex = 12;
            this.label12.Text = "时间格式";
            // 
            // bnDateTime
            // 
            this.bnDateTime.Location = new System.Drawing.Point(40, 74);
            this.bnDateTime.Name = "button1";
            this.bnDateTime.Size = new System.Drawing.Size(93, 23);
            this.bnDateTime.TabIndex = 16;
            this.bnDateTime.Text = "生成指定时间";
            this.bnDateTime.UseVisualStyleBackColor = true;
            this.bnDateTime.Click += new System.EventHandler(this.bnDateTime_Click);
            // 
            // CTPControl
            // 
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "CTPControl";
            this.Size = new System.Drawing.Size(186, 509);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        private void btnRNorm_Click(object sender, EventArgs e)
        {
            REngine.SetEnvironmentVariables();
            REngine engine = REngine.GetInstance();
            engine.Initialize();
            engine.Evaluate(string.Format("set.seed({0})", Convert.ToInt32(txtSeed.Text)));
            double[] dbls = engine.Evaluate(string.Format("rnorm({0},{1},{2})", Convert.ToInt32(txtCount.Text), Convert.ToDouble(txtMean.Text), Convert.ToDouble(txtSd.Text))).AsNumeric().ToArray();
            ArrayResizerFuncs.ResizeRows(dbls);
        }

        private void btnRNorm_t_Click(object sender, EventArgs e)
        {
            try
            {
                REngine.SetEnvironmentVariables();
                REngine engine = REngine.GetInstance();
                engine.Initialize();
                engine.Evaluate(string.Format("set.seed({0})", Convert.ToInt32(txtSeed.Text)));
                double[] dbls = engine.Evaluate(string.Format("rnorm({0},{1},{2})", Convert.ToInt32(txtCount_t.Text), Convert.ToDouble(txtMean_t.Text), Convert.ToDouble(txtSd_t.Text))).AsNumeric().ToArray();
                string[] tms = new string[dbls.Length];
                DateTime dt = dtPicker1.Value;
                int i = 0;
                foreach (double dbl in dbls)
                {
                    int _h = Convert.ToInt32(dbl);
                    int _m = Convert.ToInt32((dbl - _h) * 60);
                    int _s = Convert.ToInt32(((dbl - _h) * 60 - _m) * 60);

                    dt += new TimeSpan(_h, _m, _s);
                    tms[i++] = dt.ToString(cbDtfmt1.Text);
                }

                ArrayResizerFuncs.ResizeRows(tms);
                LogHelper.LogInfo("RNorm DateTime ArrayResizer Successful!");
            }
            catch (Exception ex)
            {
                LogHelper.LogError("RNorm DateTime ArrayResizer Faild!", ex);
            }
        }

        private void cbDtfmt1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dtPicker1.CustomFormat = cbDtfmt1.Text;
            if (cbDtfmt1.SelectedIndex > 11)
                dtPicker1.ShowUpDown = true;
            else
                dtPicker1.ShowUpDown = false;
        }

        private void bnDateTime_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;
            rg.Value = dtPicker2.Value.ToString(cbDtfmt2.Text);
        }

        private void cbDtfmt2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dtPicker2.CustomFormat = cbDtfmt2.Text;
            if (cbDtfmt2.SelectedIndex > 11)
                dtPicker2.ShowUpDown = true;
            else
                dtPicker2.ShowUpDown = false;
        }
    }

    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        private static IRibbonUI ui_ribbon;
        private static string con_str;
        private static bool mark_cell;
        private static bool row_col;
        private static bool hide_show;
        private static bool open_close;
        private static int merge_split;
        private static int encode_format;

        private static REngine engine;
        private static List<int> sp_ids;

        public static int row_clr;
        public static int col_clr;
        public static float clr_transparent;


        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;    //最左坐标
            public int Top;     //最上坐标
            public int Right;   //最右坐标
            public int Bottom;  //最下坐标
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;    //最左坐标
            public int Y;     //最上坐标
        }

        [DllImport("user32.dll")]
        private static extern int GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        private static extern int GetClientRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32")]
        private static extern int ClientToScreen(IntPtr hWnd, ref POINT lpPoint);
        [DllImport("user32")]
        private static extern int ScreenToClient(IntPtr hWnd, ref POINT lpPoint);
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        public override string GetCustomUI(string RibbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='ui_onLoad' loadImage='LoadImage'>
                <ribbon  startFromScratch='false'>
                    <tabs>
                        <tab id='tb1' label='Excel-DNA 工具' insertAfterMso='TabHome'>
                            <group id='gp1' label='基本功能'>
                                <button id='bn101' label='转至顶端' image='Back' size='large' onAction='ui_bnBack'/>
                                <splitButton id='sb110' size='large'>
                                    <button id='bn110' getLabel='ui_getMergeSplitLabel' getImage='ui_getMergeSplitImage' onAction='ui_bnMergeSplit'/>
                                    <menu id='mu110'>
                                        <button id='bn111' label='合并选定单元格' imageMso='MergeCells' onAction='ui_bnMergeCells'/>
                                        <button id='bn112' label='拆分选定单元格' imageMso='SplitCells' onAction='ui_bnSplitCells'/>
                                        <button id='bn113' label='拆分所有单元格' imageMso='SplitCells' onAction='ui_bnSplitAll'/>
                                    </menu>
                                </splitButton>
                                <separator id='sep120'/>
                                <button id='bn121' label='清空' imageMso='Clear' onAction='ui_bnClear'/>
                                <button id='bn122' label='保护' imageMso='SheetProtect' onAction='ui_bnLock'/>
                                <button id='bn123' label='解锁' imageMso='ReviewProtectWorkbook' onAction='ui_bnUnLock'/>
                                <separator id='sep130'/>
                                <splitButton id='sb130' size='large'>
                                    <button id='bn130' getLabel='ui_getMarkLabel' getImage='ui_getMarkImage' onAction='ui_bnMarkCell'/>
                                    <menu id='mu130'>
                                        <button id='bn131' label='激活标记' imageMso='TableSelect' onAction='ui_bnEnMark'/>
                                        <button id='bn132' label='取消标记' imageMso='TableSelectCell' onAction='ui_bnUnMark'/>
                                    </menu>
                                </splitButton>
                                <splitButton id='sb140' size='large'>
                                    <button id='bn140' getLabel='ui_getRowColLabel' getImage='ui_getRowColImage' onAction='ui_bnRowColSet'/>
                                    <menu id='mu140'>
                                        <button id='bn141' label='隐藏活动列' imageMso='TableColumnsInsertLeft' onAction='ui_bnHideCol'/>
                                        <button id='bn142' label='显示隐藏列' imageMso='TableColumnSelect' onAction='ui_bnShowCol'/>
                                        <button id='bn143' label='隐藏活动行' imageMso='TableRowsInsertAboveWord' onAction='ui_bnHideRow'/>
                                        <button id='bn144' label='显示隐藏行' imageMso='TableRowSelect' onAction='ui_bnShowRow'/>
                                    </menu>
                                </splitButton>
                            </group>
                            <group id='gp2' label='数据连接'>
                                <button id='bn201' label='数据连接' image='DataLink' size='large' onAction='ui_bnDataLink'/>
                                <button id='bn202' label='数据信息' imageMso='DatabaseSqlServer' size='large' onAction='ui_bnDataInfo'/>
                                <button id='bn203' label='字段信息' imageMso='PivotMoveField' size='large' onAction='ui_bnFieldInfo'/>
                                <separator id='sep210'/>
                                <button id='bn211' label='数据查询' imageMso='FindDialog' size='large' onAction='ui_bnDataFind'/>
                                <button id='bn212' label='数据过滤' imageMso='FilterToggleFilter' size='large' onAction='ui_bnDataFilter'/>
                            </group>
                            <group id='gp3' label='RDotNet工具'>
                                <button id='bn301' label='实例R系统' image='InitR' size='large' onAction='ui_bnInitR'/>
                                <button id='bn302' label='运行R脚本' image='Script' size='large' onAction='ui_bnExecScript'/>
                                <separator id='sep310'/>
                                <button id='bn311' label='浮点数组' imageMso='FunctionWizard' onAction='ui_bnTestDoubles'/>
                                <button id='bn312' label='数字数组' imageMso='FieldCodes' onAction='ui_bnTestNumeric'/>
                                <button id='bn313' label='字符数组' imageMso='ContentControlText' onAction='ui_bnTestCharacter'/>
                                <separator id='sep320'/>
                                <button id='bn321' label='加载CSV数据' imageMso='ChartShowData' size='large' onAction='ui_bnLoadCsvData'/>
                                <separator id='sep330'/>
                                <splitButton id='sb330' size='large'>
                                    <button id='bn330' label='生成f(x)图' imageMso='ChartLines' onAction='ui_bnPlotCurve'/>
                                    <menu id='mu330'>
                                        <button id='bn331' label='散点图' imageMso='ChartTrendline' onAction='ui_bnPlotPoint'/>
                                        <button id='bn332' label='直方图' imageMso='ChartAxes' onAction='ui_bnPlotBar'/>
                                        <button id='bn333' label='堆面图' imageMso='ChartDataLabel' onAction='ui_bnPlotArea'/>
                                        <button id='bn334' label='三维图' imageMso='ChartWall' onAction='ui_bnPlotPersp'/>
                                        <button id='bn335' label='等线图' imageMso='ChartTrendline' onAction='ui_bnPlotContour'/>
                                        <button id='bn336' label='直线图' imageMso='ChartLines' onAction='ui_bnPlotLines'/>
                                    </menu>
                                </splitButton>
                            </group>
                            <group id='gp4' label='阅读模式'>
                                <splitButton id='sb400' size='large'>
                                    <button id='bn400' getLabel='ui_getSpotLightLabel' image='Light' onAction='ui_bnSpotLight'/>
                                    <menu id='mu400'>
                                        <button id='bn401' label='打开聚光灯' image='Open' onAction='ui_bnOpenSpotLight'/>
                                        <button id='bn402' label='关闭聚光灯' image='Close' onAction='ui_bnCloseSpotLight'/>
                                        <button id='bn403' label='颜色设置' image='Set' onAction='ui_bnSpotLightColor'/>
                                    </menu>
                                </splitButton>
                                <toggleButton id='bn404' getLabel='ui_getCTPLabel' image='TaskPane' size='large' onAction='ui_bnSetCTP'/>
                            </group>
                            <group id='gp5' label='关于'>
                                <menu id='mu500' label='Excel 消息' image='Process' size='large'>
                                    <menuSeparator id='msep501' title='Excel Sheet Events'/>
                                    <toggleButton id='bn501' label='SheetChange' image='True' onAction='ui_bnSheetChange'/>
                                    <toggleButton id='bn502' label='SheetSelectionChange' image='False' onAction='ui_bnSheetSelectionChange'/>
                                    <toggleButton id='bn503' label='SheetActivate' image='True' onAction='ui_bnSheetActivate'/>
                                    <toggleButton id='bn504' label='SheetDeactivate' image='False' onAction='ui_bnSheetDeactivate'/>
                                    <toggleButton id='bn505' label='SheetBeforeDoubleClick' image='True' onAction='ui_bnSheetBeforeDoubleClick'/>
                                    <toggleButton id='bn506' label='SheetBeforeRightClick' image='False' onAction='ui_bnSheetBeforeRightClick'/>
                                    <toggleButton id='bn507' label='SheetCalculate' image='True' onAction='ui_bnSheetCalculate'/>
                                    <menuSeparator id='msep508' title='Excel Book Events'/>
                                    <toggleButton id='bn508' label='WorkbookAddinInstall' image='True' onAction='xl_WorkbookAddinInstall'/>
                                    <toggleButton id='bn509' label='WorkbookAddinUninstall' image='False' onAction='xl_WorkbookAddinUninstall'/>
                                    <toggleButton id='bn510' label='WorkbookActivate' image='True' onAction='ui_bnWorkbookActivate'/>
                                    <toggleButton id='bn511' label='WorkbookDeactivate' image='False' onAction='ui_bnWorkbookDeactivate'/>
                                    <toggleButton id='bn512' label='WorkbookOpen' image='True' onAction='xl_WorkbookOpen'/>
                                    <toggleButton id='bn513' label='WorkbookNewSheet' image='False' onAction='ui_bnWorkbookNewSheet'/>
                                    <toggleButton id='bn514' label='WorkbookBeforeSave' image='True' onAction='ui_bnWorkbookBeforeSave'/>
                                    <toggleButton id='bn515' label='WorkbookBeforeClose' image='False' onAction='ui_bnWorkbookBeforeClose'/>
                                    <toggleButton id='bn516' label='WorkbookRowsetComplete' image='True' onAction='ui_bnWorkbookRowsetComplete'/>
                                    <menuSeparator id='msep517' title='Excel Window Events'/>
                                    <toggleButton id='bn517' label='WindowActivate' image='True' onAction='xl_WindowActivate'/>
                                    <toggleButton id='bn518' label='WindowDeactivate' image='False' onAction='xl_WindowDeactivate'/>
                                </menu>
                                <button id='bn520' label='万年日历' image='Calendar' size='large' onAction='ui_bnCalendar'/>
                                <separator id='sep520'/>
                                <splitButton id='sb530' size='large'>
                                    <button id='bn530' getLabel='ui_getZxingLabel' image='ZxingFMT' onAction='ui_bnZxing'/>
                                    <menu id='mu530'>
                                        <button id='bn531' label='标准条形码' image='ZxingEAN' onAction='ui_bnEAN13'/>
                                        <button id='bn532' label='条形二维码' image='ZxingPDF' onAction='ui_bnPDF417'/>
                                        <button id='bn533' label='标准二维码' image='ZxingQRD' onAction='ui_bnQRCode'/>
                                        <button id='bn534' label='图标二维码' image='ZxingMID' onAction='ui_bnQRCodeMid'/>
                                    </menu>
                                </splitButton>
                                <separator id='sep530'/>
                                <button id='bn535' label='退出Excel' image='Exit' size='large' onAction='ui_bnExitExcel'/>
                                <splitButton id='sb540' size='large'>
                                    <button id='bn540' label='DNA  帮助' imageMso='Help' onAction='ui_bnHelp'/>
                                    <menu id='mu540'>
                                        <button id='bn541' label='主页...' imageMso='MsnLogo' onAction='RunTagMacro' tag='RunNet'/>
                                        <button id='bn542' label='注册...' imageMso='FormRegionMenu' onAction='ui_bnReg'/>
                                        <button id='bn543' label='工具栏' imageMso='LinkBarCustom' onAction='ui_bnAddBar'/>
                                    </menu>
                                </splitButton>
                            </group>
                        </tab>
                    </tabs>
                </ribbon>
            </customUI>";
        }

        public void ui_onLoad(IRibbonUI ribbon)
        {
            try
            {
                ui_ribbon = ribbon;
                mark_cell = true;
                open_close = true;
                merge_split = 0;
                encode_format = 0;
                sp_ids = new List<int>();
                for (int i = 0; i < 4; i++) sp_ids.Add(0);

                Color clr;
                clr = Color.FromArgb(0, 0, 255);
                row_clr = (int)(((uint)clr.B << 16) | (ushort)(((ushort)clr.G << 8) | clr.R));
                clr = Color.FromArgb(255, 0, 0);
                col_clr = (int)(((uint)clr.B << 16) | (ushort)(((ushort)clr.G << 8) | clr.R));
                clr_transparent = 0.6f;

                //Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
                //Excel.Workbook wb = xlApp.ActiveWorkbook;
                //if (wb == null) return;
                //StringBuilder sb;
                //Vbe.VBComponent xlModule;
                //Vbe.VBProject prj;
                //prj = wb.VBProject;
                //sb = new StringBuilder();

                //sb.Append("Sub" + "ui_SheetChange()" + "\n");
                //sb.Append("\t" + "msgbox \"" + "btn_Click()" + "\"\n");
                //sb.Append("End Sub");

                //xlModule = wb.VBProject.VBComponents.Add(Vbe.vbext_ComponentType.vbext_ct_StdModule);
                //xlModule.CodeModule.AddFromString(sb.ToString());

                ui_ribbon.Invalidate();
                LogHelper.LogInfo("Ribbon Load Successful!");
            }
            catch (Exception ex)
            {
                LogHelper.LogError("Ribbon Load Faild!", ex);
            }
        }

        public void ui_bnBack(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Range["A1:A1"];
            if (rg == null) return;
            rg.Select();
        }

        public string ui_getMergeSplitLabel(IRibbonControl control)
        {
            switch (merge_split)
            {
                case 1:
                    return "拆分选定单元格";
                case 2:
                    return "拆分所有单元格";
                default:
                    return "合并选定单元格";
            }
        }

        public string ui_getMergeSplitImage(IRibbonControl control)
        {
            if (merge_split == 0)
                return "MergeCells";
            else
                return "SplitCells";
        }

        public void ui_bnMergeSplit(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            switch (merge_split)
            {
                case 1:
                    {
                        if (rg.MergeCells == true)
                        {
                            string mergeAddress = rg.Address;
                            rg.UnMerge();
                            rg = ws.get_Range(mergeAddress);
                            rg.Value = rg[1].Value;
                        }
                    }
                    break;
                case 2:
                    {
                        xlApp.FindFormat.MergeCells = true;

                        while (true)
                        {
                            Excel.Range mergedCell = ws.UsedRange.Find("", Type.Missing, Type.Missing, Excel.XlLookAt.xlPart, Type.Missing, Excel.XlSearchDirection.xlNext, Type.Missing, Type.Missing, true);
                            if (mergedCell.MergeCells == false) break;
                            string mergeAddress = mergedCell.MergeArea.Address;
                            mergedCell.MergeArea.UnMerge();
                            rg = ws.get_Range(mergeAddress);
                            rg.Value = mergedCell.Value;
                        }

                        xlApp.FindFormat.Clear();
                    }
                    break;
                default:
                    rg.Merge();
                    break;
            }
        }

        public void ui_bnMergeCells(IRibbonControl control)
        {
            merge_split = 0;
            ui_ribbon.Invalidate();
        }

        public void ui_bnSplitCells(IRibbonControl control)
        {
            merge_split = 1;
            ui_ribbon.Invalidate();
        }

        public void ui_bnSplitAll(IRibbonControl control)
        {
            merge_split = 2;
            ui_ribbon.Invalidate();
        }

        public void ui_bnClear(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.UsedRange;
            if (rg == null) return;
            rg.Clear();
        }

        public void ui_bnLock(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;
            ws.Protect();
            rg.Style.Locked = true;
        }

        public void ui_bnUnLock(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;
            ws.Unprotect();
            rg.Style.Locked = false;
        }

        public string ui_getMarkLabel(IRibbonControl control)
        {
            return mark_cell ? "激活标记" : "取消标记";
        }

        public string ui_getMarkImage(IRibbonControl control)
        {
            return mark_cell ? "TableSelect" : "TableSelectCell";
        }

        public void ui_bnMarkCell(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;
            rg.Interior.Color = mark_cell ? Color.FromArgb(218, 218, 218) : Color.FromArgb(255, 255, 255);
        }

        public void ui_bnEnMark(IRibbonControl control)
        {
            mark_cell = true;
            ui_ribbon.Invalidate();
        }

        public void ui_bnUnMark(IRibbonControl control)
        {
            mark_cell = false;
            ui_ribbon.Invalidate();
        }

        public string ui_getRowColLabel(IRibbonControl control)
        {
            if (hide_show)
            {
                if (row_col)
                    return "隐藏活动行";
                else
                    return "隐藏活动列";
            }
            else
            {
                if (row_col)
                    return "显示隐藏行";
                else
                    return "显示隐藏列";
            }
        }

        public string ui_getRowColImage(IRibbonControl control)
        {
            if (hide_show)
            {
                if (row_col)
                    return "TableRowsInsertAboveWord";
                else
                    return "TableColumnsInsertLeft";
            }
            else
            {
                if (row_col)
                    return "TableRowSelect";
                else
                    return "TableColumnSelect";
            }
        }

        public void ui_bnRowColSet(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;

            if (hide_show)
            {
                Excel.Range rg = ws.Application.Selection;
                if (rg == null) return;

                if (row_col)
                    rg.EntireRow.Hidden = true;
                else
                    rg.EntireColumn.Hidden = true;
            }
            else
            {
                if (row_col)
                {
                    Excel.Range rg = ws.Application.Rows;
                    if (rg == null) return;
                    rg.EntireRow.Hidden = false;
                }
                else
                {
                    Excel.Range rg = ws.Application.Columns;
                    if (rg == null) return;
                    rg.EntireColumn.Hidden = false;
                }
            }
        }

        public void ui_bnHideCol(IRibbonControl control)
        {
            row_col = false;
            hide_show = true;
            ui_ribbon.Invalidate();
        }

        public void ui_bnShowCol(IRibbonControl control)
        {
            row_col = false;
            hide_show = false;
            ui_ribbon.Invalidate();
        }

        public void ui_bnHideRow(IRibbonControl control)
        {
            row_col = true;
            hide_show = true;
            ui_ribbon.Invalidate();
        }

        public void ui_bnShowRow(IRibbonControl control)
        {
            row_col = true;
            hide_show = false;
            ui_ribbon.Invalidate();
        }

        public void ui_bnDataLink(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;

            MSDASC.DataLinks dlg = new MSDASC.DataLinks();
            ADODB.Connection con = (ADODB.Connection)dlg.PromptNew();
            DotNet.Utilities.OleDbBase _con = new DotNet.Utilities.OleDbBase();
            con_str = con.ConnectionString;
            _con.ConStr = con_str;

            string[,] tbls = _con.GetTables();
            ArrayResizerFuncs.ResizeRange(tbls);

            //try
            //{
            //    MSDASC.DataLinks dlg = new MSDASC.DataLinks();
            //    ADODB.Connection _con = (ADODB.Connection)dlg.PromptNew();
            //    con_str = _con.ConnectionString;

            //    _con.Open();
            //    ADODB.Recordset _tbl = _con.OpenSchema(ADODB.SchemaEnum.adSchemaTables);
            //    _tbl.MoveFirst();

            //    int i = 0;
            //    while (!_tbl.EOF)
            //    {
            //        ADODB.Fields _fields = _tbl.Fields;

            //        ADODB.Field _field = _fields["TABLE_TYPE"];

            //        string val = _field.Value;

            //        if (String.Compare(val, "TABLE") == 0 || String.Compare(val, "VIEW") == 0)
            //        {
            //            _field = _fields["TABLE_NAME"];
            //            rg.get_Offset(i++, 0).Value = _field.Value;
            //        }

            //        _tbl.MoveNext();
            //    }

            //    _tbl.Close();
            //    _con.Close();
            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show(e.Message);
            //}
        }

        public void ui_bnDataInfo(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;

            DotNet.Utilities.OleDbBase _con = new DotNet.Utilities.OleDbBase();
            _con.ConStr = con_str;
            _con.Table = rg.Value;

            object[,] objs = _con.GetValues();
            ArrayResizerFuncs.ResizeRange(objs);
        }

        public void ui_bnFieldInfo(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;

            DotNet.Utilities.OleDbBase _con = new DotNet.Utilities.OleDbBase();
            _con.ConStr = con_str;
            _con.Table = rg.Value;

            string[,] objs = _con.GetFields();
            ArrayResizerFuncs.ResizeRange(objs);
        }

        public void ui_bnDataFind(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            if (rg.Count != 3)
            {
                LogDisplay.WriteLine("Error Select The Excel.Range Is Not 3 !");
                return;
            }

            Excel.Range rg1 = rg.Areas[3];

            DotNet.Utilities.OleDbBase _con = new DotNet.Utilities.OleDbBase();
            _con.ConStr = con_str;
            _con.Table = rg.Areas[1].Value;
            _con.Field = rg.Areas[2].Value;

            object[,] objs = _con.FindValues(Convert.ToString(rg1.Value));
            ArrayResizerFuncs.ResizeRange(objs);
        }

        public void ui_bnDataFilter(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            if (rg.Count != 3)
            {
                LogDisplay.WriteLine("Error Select The Excel.Range Is Not 3 !");
                return;
            }

            Excel.Range rg1 = rg.Areas[3];

            DotNet.Utilities.OleDbBase _con = new DotNet.Utilities.OleDbBase();
            _con.ConStr = con_str;
            _con.Table = rg.Areas[1].Value;
            _con.Field = rg.Areas[2].Value;

            object[] objs = _con.GetValue(Convert.ToInt32(rg1.Value));
            ArrayResizerFuncs.ResizeRows(objs);
        }

        public void ui_bnInitR(IRibbonControl control)
        {
            try
            {
                REngine.SetEnvironmentVariables();
                engine = REngine.GetInstance();
                engine.Initialize();

                LogDisplay.WriteLine("initializing RDotNet Successful !");
            }
            catch (Exception ex)
            {
                LogDisplay.WriteLine("Error initializing RDotNet: " + ex.Message);
            }
        }

        public void ui_bnExecScript(IRibbonControl control)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Filter = "R脚本文件(*.r)|*.r";
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                engine.Evaluate("source('" + fdlg.FileName + "')");
                LogDisplay.WriteLine("Exec RDotNet Script Successful !");
            }
            else { LogDisplay.WriteLine("Error Exec RDotNet Script !"); }
        }

        public void ui_bnTestDoubles(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;

            double[] dbls = engine.Evaluate("rnorm(10,50,1)").AsNumeric().ToArray();
            ArrayResizerFuncs.ResizeRows(dbls);

            LogDisplay.WriteLine("Engine Evaluate double[] Successful !");
        }

        public void ui_bnTestNumeric(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;

            // .NET Framework array to R vector.
            NumericVector group1 = engine.CreateNumericVector(new double[] { 30.02, 29.99, 30.11, 29.97, 30.01, 29.99 });
            if (group1.Length == 0)
            {
                LogDisplay.WriteLine("Error Engine Evaluate NumericVector !");
                return;
            }

            engine.SetSymbol("group1", group1);

            // Direct parsing from R script.
            NumericVector group2 = engine.Evaluate("group2 <- c(29.89, 29.93, 29.72, 29.98, 30.02, 29.98)").AsNumeric();

            // Test difference of mean and get the P-value.
            GenericVector testResult = engine.Evaluate("t.test(group1, group2)").AsList();
            if (testResult.Length == 0)
            {
                LogDisplay.WriteLine("Error Engine Evaluate GenericVector !");
                return;
            }

            double p = testResult["p.value"].AsNumeric()[0];

            rg.Value = string.Format("Group1: [{0}], Group2: [{1}], P-value = {2:0.000}", string.Join(", ", group1), string.Join(", ", group2), p);

            LogDisplay.WriteLine("Engine Evaluate NumericVector Successful !");
        }

        public void ui_bnTestCharacter(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;

            var _grid = engine.Evaluate("function(x, y) { expand.grid(x=x, y=y) }").AsFunction();
            var v1 = engine.CreateIntegerVector(new[] { 1, 2, 3 });
            var v2 = engine.CreateCharacterVector(new[] { "a", "b", "c" });
            var df = _grid.Invoke(new SymbolicExpression[] { v1, v2 }).AsDataFrame();
            engine.SetSymbol("cases", df);
            // Not correct: the 'y' column is a "factor". This returns "1", "1", "1", "2", "2", etc.
            CharacterVector _letter = engine.Evaluate("cases[,'y']").AsCharacter();
            // This returns something more intuitive for C#  Returns 'a','a','a','b','b','b','c' etc.
            _letter = engine.Evaluate("as.character(cases[,'y'])").AsCharacter();
            // In the manner of R.NET, try
            string[] strs = engine.Evaluate("cases[,'y']").AsFactor().GetFactors();
            ArrayResizerFuncs.ResizeRows(strs);

            LogDisplay.WriteLine("Engine Evaluate CharacterVector Successful !");
        }

        public void ui_bnLoadCsvData(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;

            DataFrame df = engine.Evaluate("df<-read.table(file.choose(),header=TRUE, sep = ',')").AsDataFrame();

            int rows = df.RowCount;
            int cols = df.ColumnCount;
            if (rows == 0 || cols == 0)
            {
                LogDisplay.WriteLine("Error Engine Evaluate DataFrame !");
                return;
            }

            object[,] objs = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    objs[i, j] = df[i, j];

            ArrayResizerFuncs.ResizeRange(objs);

            LogDisplay.WriteLine("Engine Evaluate DataFrame Successful !");
        }

        public void ui_bnPlotPoint(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            int row = rg.Rows.Count;
            int col = rg.Columns.Count;
            if (col != 4 || row < 2)
            {
                LogDisplay.WriteLine("Error Select The Excel.Range 4 Cols!");
                return;
            }

            double[] x = new double[row * 3];
            double[] y = new double[row * 3];
            for (int i = 0; i < row; i++)
            {
                x[i] = rg.Cells[i + 1, 1].Value;
                x[i + row] = rg.Cells[i + 1, 1].Value;
                x[i + row * 2] = rg.Cells[i + 1, 1].Value;
                y[i] = rg.Cells[i + 1, 2].Value;
                y[i + row] = rg.Cells[i + 1, 3].Value;
                y[i + row * 2] = rg.Cells[i + 1, 4].Value;
            }

            var vx = engine.CreateNumericVector(x);
            var vy = engine.CreateNumericVector(y);
            engine.SetSymbol("vx", vx);
            engine.SetSymbol("vy", vy);
            engine.Evaluate("require(ggplot2)");
            engine.Evaluate("library(ggplot2)");
            engine.Evaluate("require(gcookbook)");
            engine.Evaluate("library(gcookbook)");
            engine.Evaluate("set.seed(1234)");
            engine.Evaluate(string.Format("vz <- rep(c(5,10,20), each = {0})", row));
            engine.Evaluate("df <- data.frame(vx = vx, vy = vy, vz = vz)");

            string mapping = "aes(x = vx,y = vy, colour = vz)";
            string geom_point = "geom_point()+ geom_rug(position = 'jitter', size = 1)";
            string stat_smooth = "geom_point() + stat_smooth(method=lm)";
            engine.Evaluate(string.Format("ggplot(data = df, mapping = {0}) + {1} + {2}", mapping, geom_point, stat_smooth));

            //使用Logistic回归拟合一个二分类的样本
            //engine.Evaluate("require(MASS)");
            //engine.Evaluate("library(MASS)");
            //engine.Evaluate("b <- biopsy");
            //engine.Evaluate("b$classn[b$class=='benign'] <- 0");
            //engine.Evaluate("b$classn[b$class=='malignant'] <- 1");
            //engine.Evaluate("head(b)");
            //engine.Evaluate("ggplot(b, aes(x=V1, y=classn)) + geom_point(position=position_jitter(width=.3, height=.06), alpha=.4, shape=21, size=1.5) + geom_smooth(method='glm', method.args=list(family='binomial'))");
        }

        public void ui_bnPlotBar(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            int row = rg.Rows.Count;
            int col = rg.Columns.Count;
            if (col < 2 || row < 2)
            {
                LogDisplay.WriteLine("Error Select The Excel.Range 4 Cols!");
                return;
            }

            string[] x = new string[row * 3];
            double[] y = new double[row * 3];
            for (int i = 0; i < row; i++)
            {
                x[i] = rg.Cells[i + 1, 1].Value.ToString();
                x[i + row] = rg.Cells[i + 1, 1].Value.ToString();
                x[i + row * 2] = rg.Cells[i + 1, 1].Value.ToString();
                y[i] = rg.Cells[i + 1, 2].Value;
                y[i + row] = rg.Cells[i + 1, 3].Value;
                y[i + row * 2] = rg.Cells[i + 1, 4].Value;
            }

            var vx = engine.CreateCharacterVector(x);
            var vy = engine.CreateNumericVector(y);

            engine.Evaluate("require(ggplot2)");
            engine.Evaluate("library(ggplot2)");
            engine.SetSymbol("vx", vx);
            engine.SetSymbol("vy", vy);
            engine.Evaluate("set.seed(1234)");
            engine.Evaluate(string.Format("vz <- rep(c('A','B','C'), each = {0})", row));
            engine.Evaluate("df <- data.frame(vx = vx, vy = vy, vz = vz)");
            engine.Evaluate("col <- c('darkred','skyblue','purple')");
            engine.Evaluate("ggplot(data = df, mapping = aes(x = factor(vx), y = vy, fill = vz)) + geom_bar(stat = 'identity', colour= 'black', position = 'dodge') + scale_fill_manual(values = col) + xlab('vx')");
            //engine.Evaluate("lines(density(vy))");
        }

        public void ui_bnPlotArea(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            int row = rg.Rows.Count;
            int col = rg.Columns.Count;
            if (col != 4 || row < 2)
            {
                LogDisplay.WriteLine("Error Select The Excel.Range 4 Cols!");
                return;
            }

            double[] x = new double[row * 3];
            double[] y = new double[row * 3];
            for (int i = 0; i < row; i++)
            {
                x[i] = rg.Cells[i + 1, 1].Value;
                x[i + row] = rg.Cells[i + 1, 1].Value;
                x[i + row * 2] = rg.Cells[i + 1, 1].Value;
                y[i] = rg.Cells[i + 1, 2].Value;
                y[i + row] = rg.Cells[i + 1, 3].Value;
                y[i + row * 2] = rg.Cells[i + 1, 4].Value;
            }

            var vx = engine.CreateNumericVector(x);
            var vy = engine.CreateNumericVector(y);

            engine.Evaluate("require(ggplot2)");
            engine.Evaluate("library(ggplot2)");
            engine.Evaluate("require(dplyr)");
            engine.Evaluate("library(dplyr)");
            engine.SetSymbol("vx", vx);
            engine.SetSymbol("vy", vy);
            engine.Evaluate("set.seed(1234)");
            engine.Evaluate(string.Format("vz <- rep(c('A','B','C'), each={0})", row));
            engine.Evaluate("df <- data.frame(vx = vx, vy = vy, vz = vz)");
            engine.Evaluate("df_by_vz <- group_by(.data = df, vx)");
            engine.Evaluate("df_sum <- mutate(.data = df_by_vz, vy2 = vy/sum(vy))");

            string mapping = "aes(x = vx, y = vy2, fill = vz)";
            string geom_area = "geom_area(alpha=.6)";
            string geom_line = "geom_line(colour = 'black', size = 1, position = 'stack', alpha = 0.6)";
            string guides = "guides(fill = guide_legend(reverse=FALSE))";

            engine.Evaluate(string.Format("ggplot(data = df_sum, mapping = {0}) + {1} + {2} + {3}", mapping, geom_area, geom_line, guides));
        }

        public void ui_bnPlotPersp(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            int row = rg.Rows.Count;
            int col = rg.Columns.Count;
            if (col != 2 || row < 2)
            {
                LogDisplay.WriteLine("Error Select The Excel.Range 2 Cols!");
                return;
            }

            string fx = "23.86+5.525*b-2.5725*a-6.6413*b^2-5.1862*a^2";
            engine.Evaluate("vx <- 1:100").AsNumeric();
            engine.Evaluate("vy <- 5:105").AsNumeric();

            if (!string.IsNullOrEmpty(rg.Cells[1, 1].Value))
            {
                fx = rg.Cells[1, 1].Value;
                double[] x = new double[row - 1];
                double[] y = new double[row - 1];
                for (int i = 0; i < row - 1; i++)
                {
                    x[i] = rg.Cells[i + 2, 1].Value;
                    y[i] = rg.Cells[i + 2, 2].Value;
                }

                var vx = engine.CreateNumericVector(x);
                var vy = engine.CreateNumericVector(y);
                engine.SetSymbol("vx", vx);
                engine.SetSymbol("vy", vy);
            }

            engine.Evaluate(string.Format("f = function(a,b) {{{0}}}", fx));
            engine.Evaluate("vz = outer(vx,vy,f)");
            engine.Evaluate("persp(vx,vy,vz)");
        }

        public void ui_bnPlotContour(IRibbonControl control)
        {
            var x = engine.Evaluate("x <- 1:100").AsNumeric();
            var y = engine.Evaluate("y <- 5:105").AsNumeric();

            engine.Evaluate("f = function(a,b) {23.86+5.525*b-2.5725*a-6.6413*b^2-5.1862*a^2}"); //evaluate function
            engine.Evaluate("z = outer(x,y,f)");
            engine.Evaluate("contour(x,y,z,nlevels=10,main='the contour chart', xlab='X-axis', ylab='Y-axis')");

            engine.Evaluate("par(mar = rep(1, 4))");
            engine.Evaluate("x = 10 * (1:nrow(volcano))");
            engine.Evaluate("y = 10 * (1:ncol(volcano))");
            engine.Evaluate("image(x, y, volcano, col = terrain.colors(100), axes = FALSE)");
            engine.Evaluate("contour(x, y, volcano, levels = seq(90, 200, by = 5),add = TRUE, col = 'peru')");
            engine.Evaluate("box()");
        }

        public void ui_bnPlotLines(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            int row = rg.Rows.Count;
            int col = rg.Columns.Count;
            if (col != 4 || row < 2)
            {
                LogDisplay.WriteLine("Error Select The Excel.Range 4 Cols!");
                return;
            }

            double[] x = new double[row * 2];
            double[] y = new double[row * 2];
            for (int i = 0; i < row; i++)
            {
                x[i] = rg.Cells[i + 1, 1].Value;
                x[i + row] = rg.Cells[i + 1, 3].Value;
                y[i] = rg.Cells[i + 1, 2].Value;
                y[i + row] = rg.Cells[i + 1, 4].Value;
            }

            var vx = engine.CreateNumericVector(x);
            var vy = engine.CreateNumericVector(y);

            engine.Evaluate("require(ggplot2)");
            engine.Evaluate("library(ggplot2)");
            engine.SetSymbol("vx", vx);
            engine.SetSymbol("vy", vy);
            //engine.Evaluate("require(lubridate)");
            //engine.Evaluate("library(lubridate)");
            engine.Evaluate("set.seed(1234)");
            //engine.Evaluate("year <- rep(1990:2015, times = 2)");
            //engine.Evaluate("type <- rep(c('A','B'), each = 26)");
            engine.Evaluate(string.Format("vz <- rep(c('A','B'), each = {0})", row));
            //engine.Evaluate("value <- c(runif(26),runif(26, min = 1,max = 1.5))");
            engine.Evaluate("df <- data.frame(vx = vx, vy = vy, vz = vz)");
            //engine.Evaluate("ggplot(data = df, mapping = aes(x = year, y = value, colour = type)) + geom_line() + geom_point()");

            string mapping = "aes(x = vx, y = vy, colour = vz)";
            string geom = "geom_line() + geom_point()";
            string scale_linetype_manual = "scale_linetype_manual(values = c(1,2))";
            string scale_color_manual = "scale_color_manual(values = c('steelblue','darkred'))";
            string scale_shape_manual = "scale_shape_manual(values = c(21,23))";
            string scale_fill_manual = "scale_fill_manual(values = c('red','black'))";

            engine.Evaluate(string.Format("ggplot(data = df, mapping = {0}) + {1} + {2} + {3} + {4} + {5}", mapping,
                geom, scale_linetype_manual, scale_color_manual, scale_shape_manual, scale_fill_manual));
        }

        public void ui_bnPlotCurve(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            if (rg.Count != 3)
            {
                LogDisplay.WriteLine("Error Select The Excel.Range Is Not 3 !");
                return;
            }

            string fx = "exp(-abs(x))*sin(2*pi*x)";
            if (!string.IsNullOrEmpty(rg[1].Value))
                fx = rg[1].Value;

            double up = -5;
            if (!string.IsNullOrEmpty(rg[2].Value))
                up = rg[2].Value;

            double down = +5;
            if (!string.IsNullOrEmpty(rg[3].Value))
                down = rg[3].Value;

            engine.Evaluate(string.Format("f = function(x) {0}", fx));
            engine.Evaluate(string.Format("curve(f,{0},{1},main='dampened sine wave')", up, down));
        }

        public void ui_bnSpotLight(IRibbonControl control)
        {
            ui_bnSheetSelectionChange(control, open_close);
            //IntPtr hwnd_main = ExcelDnaUtil.WindowHandle;
            ////IntPtr hwnd_main = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "XLMAIN", null);
            //IntPtr hwnd_desk = hwnd_main != IntPtr.Zero ? FindWindowEx(hwnd_main, IntPtr.Zero, "XLDESK", null) : IntPtr.Zero;
            //IntPtr hwnd_sheet = hwnd_desk != IntPtr.Zero ? FindWindowEx(hwnd_desk, IntPtr.Zero, "Excel7", null) : IntPtr.Zero;
            //if (hwnd_sheet != IntPtr.Zero)
            //{
            //    RECT rc = new RECT();
            //    GetClientRect(hwnd_sheet, out rc);
            //    Graphics gra = Graphics.FromHwnd(hwnd_sheet);
            //    Pen pen = new Pen(Color.Red);
            //    pen.Width = 2;

            //    Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            //    Excel.Workbook wb = xlApp.ActiveWorkbook;
            //    if (wb == null) return;
            //    Excel.Worksheet ws = wb.ActiveSheet;
            //    Excel.Range rg = ws.Cells[1];

            //    int x = xlApp.ActiveWindow.PointsToScreenPixelsX(rg.Left);
            //    int y = xlApp.ActiveWindow.PointsToScreenPixelsY(rg.Top);
            //    POINT pt = new POINT();
            //    pt.X = x;
            //    pt.Y = y;
            //    gra.DrawLine(pen, pt.X, 0, pt.X, rc.Bottom);
            //    gra.DrawLine(pen, 0, pt.Y, rc.Right, pt.Y);

            //    pen.Dispose();
            //    gra.Dispose();
            //}
        }

        public string ui_getSpotLightLabel(IRibbonControl control)
        {
            if (open_close)
                return "打开聚光灯";
            else
                return "关闭聚光灯";
        }

        public void ui_bnOpenSpotLight(IRibbonControl control)
        {
            open_close = true;
            ui_ribbon.Invalidate();
        }

        public void ui_bnCloseSpotLight(IRibbonControl control)
        {
            open_close = false;
            ui_ribbon.Invalidate();
        }

        public void ui_bnSpotLightColor(IRibbonControl control)
        {
            frmSpotLight frm = new frmSpotLight();
            frm.ShowDialog();
        }

        public string ui_getCTPLabel(IRibbonControl control)
        {
            if (!CTPManager.Visible)
                return "显示窗格";
            else
                return "关闭窗格";
        }

        public void ui_bnSetCTP(IRibbonControl control, bool pressed)
        {
            if (pressed)
                CTPManager.OpenCTP();
            else
                CTPManager.CloseCTP();

            ui_ribbon.Invalidate();
        }

        private void xl_SheetChange(object Sh, Excel.Range Target)
        {
            LogDisplay.WriteLine("AppEvents_SheetChangeEvent");
            Excel.Worksheet ws = Sh as Excel.Worksheet;
            if (ws == null) return;
            Excel.Shapes sps = ws.Shapes;
            if (sps == null) return;
            foreach (Excel.Shape sp in sps) if (sp.Type == Core.MsoShapeType.msoAutoShape) sp.Delete();
            LogDisplay.WriteLine((ws.Name + "!" + Target.get_Address(Type.Missing, Type.Missing, Excel.XlReferenceStyle.xlA1) + " = " + Target.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault)));
        }
        public void ui_SheetChange(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.SheetChange += new Excel.AppEvents_SheetChangeEventHandler(xl_SheetChange);
            else
                xlApp.SheetChange -= xl_SheetChange;
        }

        private void xl_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            Excel.Worksheet ws = Sh as Excel.Worksheet;
            if (ws == null) return;
            Excel.Shapes sps = ws.Shapes;
            if (sps == null) return;

            foreach (int id in sp_ids)
            {
                foreach (Excel.Shape sp in sps)
                {
                    if (sp.ID == id)
                    {
                        sp.Delete();
                        break;
                    }
                }
            }

            Excel.Range rg = ws.Application.Selection;
            int row = rg.Row;
            int col = rg.Column;
            int rows = rg.Rows.Count;
            int cols = rg.Columns.Count;

            //新建一个下矩形
            Excel.Range rg1 = rg.Rows[rows];
            Excel.Range rg_1 = rg.Rows[rows];
            rg1 = rg1.get_Offset(1000, 0);
            rg_1 = rg_1.get_Offset(1, 0);
            rg1 = ws.get_Range(rg1, rg_1);
            Excel.Shape sp1 = sps.AddShape(Core.MsoAutoShapeType.msoShapeRectangle, rg1.Left, rg1.Top, rg1.Width, rg1.Height);
            sp_ids[0] = sp1.ID;
            sp1.Line.Visible = Core.MsoTriState.msoFalse;  //设置为无轮廓
            sp1.Fill.Visible = Core.MsoTriState.msoTrue;   //设置填充（fill）颜色，如果=msofalse，则为无颜色填充
            sp1.Fill.ForeColor.RGB = col_clr;  //设置颜色
            sp1.Fill.Transparency = clr_transparent;  //设置透明度，数字愈大透明度越高
            sp1.Fill.Solid();  //默认设置为纯色填充，此句可省略
            sp1.Locked = true;  //锁定矩形
            sp1.LockAspectRatio = Core.MsoTriState.msoFalse; //取消图片的"锁定纵横比",调整行高时图片会相应变化
            //sp1.Fill.TwoColorGradient(Core.MsoGradientStyle.msoGradientHorizontal, 2);  //双色渐变填充，数字代表渐变方向
            //sp1.Name = "列光线";

            //新建一个右矩形
            Excel.Range rg2 = rg.Columns[cols];
            Excel.Range rg_2 = rg.Columns[cols];
            rg2 = rg2.get_Offset(0, 100);
            rg_2 = rg_2.get_Offset(0, 1);
            rg2 = ws.get_Range(rg2, rg_2);
            Excel.Shape sp2 = sps.AddShape(Core.MsoAutoShapeType.msoShapeRectangle, rg2.Left, rg2.Top, rg2.Width, rg2.Height);
            sp_ids[1] = sp2.ID;
            sp2.Line.Visible = Core.MsoTriState.msoFalse;  //设置为无轮廓
            sp2.Fill.Visible = Core.MsoTriState.msoTrue;   //设置填充（fill）颜色，如果=msofalse，则为无颜色填充
            sp2.Fill.ForeColor.RGB = row_clr;    //设置颜色
            sp2.Fill.Transparency = clr_transparent;  //设置透明度，数字愈大透明度越高
            sp2.Fill.Solid();  //默认设置为纯色填充，此句可省略
            sp2.Locked = true;  //锁定矩形
            sp2.LockAspectRatio = Core.MsoTriState.msoFalse; //取消图片的"锁定纵横比",调整行高时图片会相应变化
            //sp2.Fill.TwoColorGradient(Core.MsoGradientStyle.msoGradientHorizontal, 2);  //双色渐变填充，数字代表渐变方向
            //sp2.Name = "行光线";

            if (row > 1)  //新建一个上矩形
            {
                rg1 = rg.Rows[1];
                rg_1 = rg.Rows[1];
                rg1 = rg1.get_Offset(1 - row, 0);
                rg_1 = rg_1.get_Offset(-1, 0);
                rg1 = ws.get_Range(rg1, rg_1);
                sp1 = sps.AddShape(Core.MsoAutoShapeType.msoShapeRectangle, rg1.Left, rg1.Top, rg1.Width, rg1.Height);
                sp_ids[2] = sp1.ID;
                sp1.Line.Visible = Core.MsoTriState.msoFalse;  //设置为无轮廓
                sp1.Fill.Visible = Core.MsoTriState.msoTrue;   //设置填充（fill）颜色，如果=msofalse，则为无颜色填充
                sp1.Fill.ForeColor.RGB = col_clr;  //设置颜色
                sp1.Fill.Transparency = clr_transparent;  //设置透明度，数字愈大透明度越高
                sp1.Fill.Solid();  //默认设置为纯色填充，此句可省略
                sp1.Locked = true;  //锁定矩形
                sp1.LockAspectRatio = Core.MsoTriState.msoFalse; //取消图片的"锁定纵横比",调整行高时图片会相应变化
                //sp1.Fill.TwoColorGradient(Core.MsoGradientStyle.msoGradientHorizontal, 2);  //双色渐变填充，数字代表渐变方向
                //sp1.Name = "列光线";
            }

            if (col > 1)  //新建一个左矩形
            {
                rg2 = rg.Columns[1];
                rg_2 = rg.Columns[1];
                rg2 = rg2.get_Offset(0, 1 - col);
                rg_2 = rg_2.get_Offset(0, -1);
                rg2 = ws.get_Range(rg2, rg_2);
                sp2 = sps.AddShape(Core.MsoAutoShapeType.msoShapeRectangle, rg2.Left, rg2.Top, rg2.Width, rg2.Height);
                sp_ids[3] = sp2.ID;
                sp2.Line.Visible = Core.MsoTriState.msoFalse;  //设置为无轮廓
                sp2.Fill.Visible = Core.MsoTriState.msoTrue;   //设置填充（fill）颜色，如果=msofalse，则为无颜色填充
                sp2.Fill.ForeColor.RGB = row_clr;    //设置颜色
                sp2.Fill.Transparency = clr_transparent;  //设置透明度，数字愈大透明度越高
                sp2.Fill.Solid();  //默认设置为纯色填充，此句可省略
                sp2.Locked = true;  //锁定矩形
                sp2.LockAspectRatio = Core.MsoTriState.msoFalse; //取消图片的"锁定纵横比",调整行高时图片会相应变化
                //sp2.Fill.TwoColorGradient(Core.MsoGradientStyle.msoGradientHorizontal, 2);  //双色渐变填充，数字代表渐变方向
                //sp2.Name = "行光线";
            }
        }
        public void ui_bnSheetSelectionChange(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(xl_SheetSelectionChange);
            else
                xlApp.SheetSelectionChange -= xl_SheetSelectionChange;

            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Shapes sps = ws.Shapes;
            if (sps == null) return;

            foreach (int id in sp_ids)
            {
                foreach (Excel.Shape sp in sps)
                {
                    if (sp.ID == id)
                    {
                        sp.Delete();
                        break;
                    }
                }
            }
        }

        private void xl_SheetActivate(object Sh)
        {
            LogDisplay.WriteLine("AppEvents_SheetActivateEvent");
        }
        public void ui_bnSheetActivate(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(xl_SheetActivate);
            else
                xlApp.SheetActivate -= xl_SheetActivate;
        }

        private void xl_SheetDeactivate(object Sh)
        {
            LogDisplay.WriteLine("AppEvents_SheetDeactivateEvent");
        }
        public void ui_bnSheetDeactivate(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.SheetDeactivate += new Excel.AppEvents_SheetDeactivateEventHandler(xl_SheetDeactivate);
            else
                xlApp.SheetDeactivate -= xl_SheetDeactivate;
        }

        private void xl_SheetBeforeDoubleClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            LogDisplay.WriteLine("AppEvents_SheetBeforeDoubleClickEvent");
        }
        public void ui_bnSheetBeforeDoubleClick(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.SheetBeforeDoubleClick += new Excel.AppEvents_SheetBeforeDoubleClickEventHandler(xl_SheetBeforeDoubleClick);
            else
                xlApp.SheetBeforeDoubleClick -= xl_SheetBeforeDoubleClick;
        }

        private void xl_SheetBeforeRightClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            LogDisplay.WriteLine("AppEvents_SheetBeforeRightClickEvent");
        }
        public void ui_bnSheetBeforeRightClick(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.SheetBeforeRightClick += new Excel.AppEvents_SheetBeforeRightClickEventHandler(xl_SheetBeforeRightClick);
            else
                xlApp.SheetBeforeRightClick -= xl_SheetBeforeRightClick;
        }

        private void xl_SheetCalculate(object Sh)
        {
            LogDisplay.WriteLine("AppEvents_SheetCalculateEvent");
        }
        public void ui_bnSheetCalculate(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.ActiveCell;
            if (rg == null) return;

            CommandBars bars = ExcelCommandBarUtil.GetCommandBars();
            for (int i = 1; i <= bars.Count; i++)
            {
                rg = rg.get_Offset(1, 0);
                CommandBar bar = bars[i];
                rg.Value = bar.Name;
                CommandBarControls ctls = bar.Controls;
                for (int j = 1; j <= ctls.Count; j++)
                {
                    CommandBarControl ctl = ctls[j];
                    MsoControlType type = ctl.Type;
                    if (type == MsoControlType.msoControlButton)
                    {
                        rg = rg.get_Offset(1, 0);
                        Excel.Range rg1 = rg.get_Offset(0, 1);
                        rg.Value = ctl.Caption;
                        rg1.Value = ((CommandBarButton)ctl).FaceId;
                        rg1 = rg1.get_Offset(0, 1);
                    }
                }
            }

            if (pressed)
                xlApp.SheetCalculate += new Excel.AppEvents_SheetCalculateEventHandler(xl_SheetCalculate);
            else
                xlApp.SheetCalculate -= xl_SheetCalculate;
        }

        private void xl_WorkbookAddinInstall(Excel.Workbook Wb)
        {
            LogDisplay.WriteLine("AppEvents_WorkbookAddinInstallEvent");
        }
        public void ui_bnWorkbookAddinInstall(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WorkbookAddinInstall += new Excel.AppEvents_WorkbookAddinInstallEventHandler(xl_WorkbookAddinInstall);
            else
                xlApp.WorkbookAddinInstall -= xl_WorkbookAddinInstall;
        }

        private void xl_WorkbookAddinUninstall(Excel.Workbook Wb)
        {
            LogDisplay.WriteLine("AppEvents_WorkbookAddinUninstallEvent");
        }
        public void ui_bnWorkbookAddinUninstall(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WorkbookAddinUninstall += new Excel.AppEvents_WorkbookAddinUninstallEventHandler(xl_WorkbookAddinUninstall);
            else
                xlApp.WorkbookAddinUninstall -= xl_WorkbookAddinUninstall;
        }

        private void xl_WorkbookActivate(Excel.Workbook Wb)
        {
            LogDisplay.WriteLine("AppEvents_WorkbookActivateEvent");
        }
        public void ui_bnWorkbookActivate(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(xl_WorkbookActivate);
            else
                xlApp.WorkbookActivate -= xl_WorkbookActivate;
        }

        private void xl_WorkbookDeactivate(Excel.Workbook Wb)
        {
            LogDisplay.WriteLine("AppEvents_WorkbookDeactivateEvent");
        }
        public void ui_bnWorkbookDeactivate(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WorkbookDeactivate += new Excel.AppEvents_WorkbookDeactivateEventHandler(xl_WorkbookDeactivate);
            else
                xlApp.WorkbookDeactivate -= xl_WorkbookDeactivate;
        }

        private void xl_WorkbookOpen(Excel.Workbook Wb)
        {
            LogDisplay.WriteLine("AppEvents_WorkbookOpenEvent");
        }
        public void ui_bnWorkbookOpen(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(xl_WorkbookOpen);
            else
                xlApp.WorkbookOpen -= xl_WorkbookOpen;
        }

        private void xl_WorkbookNewSheet(Excel.Workbook Wb, object Sh)
        {
            LogDisplay.WriteLine("AppEvents_WorkbookNewSheetEvent");
        }
        public void ui_bnWorkbookNewSheet(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WorkbookNewSheet += new Excel.AppEvents_WorkbookNewSheetEventHandler(xl_WorkbookNewSheet);
            else
                xlApp.WorkbookNewSheet -= xl_WorkbookNewSheet;
        }

        private void xl_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            LogDisplay.WriteLine("AppEvents_WorkbookBeforeSaveEvent");
        }
        public void ui_bnWorkbookBeforeSave(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(xl_WorkbookBeforeSave);
            else
                xlApp.WorkbookBeforeSave -= xl_WorkbookBeforeSave;
        }

        private void xl_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            LogDisplay.WriteLine("AppEvents_WorkbookBeforeCloseEvent");
        }
        public void ui_bnWorkbookBeforeClose(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(xl_WorkbookBeforeClose);
            else
                xlApp.WorkbookBeforeClose -= xl_WorkbookBeforeClose;
        }

        private void xl_WorkbookRowsetComplete(Excel.Workbook Wb, string Description, string Sheet, bool Success)
        {
            LogDisplay.WriteLine("AppEvents_WorkbookRowsetCompleteEvent");
        }
        public void ui_bnWorkbookRowsetComplete(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WorkbookRowsetComplete += new Excel.AppEvents_WorkbookRowsetCompleteEventHandler(xl_WorkbookRowsetComplete);
            else
                xlApp.WorkbookRowsetComplete -= xl_WorkbookRowsetComplete;
        }

        private void xl_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            LogDisplay.WriteLine("AppEvents_WindowActivateEvent");
        }
        public void ui_bnWindowActivate(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WindowActivate += new Excel.AppEvents_WindowActivateEventHandler(xl_WindowActivate);
            else
                xlApp.WindowActivate -= xl_WindowActivate;
        }

        private void xl_WindowDeactivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            LogDisplay.WriteLine("AppEvents_WindowDeactivateEvent");
        }
        public void ui_bnWindowDeactivate(IRibbonControl control, bool pressed)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            if (pressed)
                xlApp.WindowDeactivate += new Excel.AppEvents_WindowDeactivateEventHandler(xl_WindowDeactivate);
            else
                xlApp.WindowDeactivate -= xl_WindowDeactivate;
        }

        public void ui_bnCalendar(IRibbonControl control)
        {
            frmCalendar frm = new frmCalendar();
            frm.ShowDialog();
        }

        public void ui_bnZxing(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            Excel.Worksheet ws = wb.ActiveSheet;
            if (ws == null) return;
            Excel.Range rg = ws.Application.Selection;
            if (rg == null) return;

            try
            {
                LogHelper.LogInfo("Begin EnCode Info Successful!");

                int row = rg.Rows.Count;
                if (encode_format == 4)
                {
                    OpenFileDialog ofdlg = new OpenFileDialog();
                    ofdlg.Filter = "图片文件|*.bmp;*.jpg;*.png;*.ico";
                    if (ofdlg.ShowDialog() == DialogResult.OK)
                    {
                        string pngpath = ofdlg.FileName;
                        for (int i = 0; i < row; i++)
                        {
                            string enstr = rg.Cells[i + 1, 1].Text;
                            if (enstr.Length == 0) enstr = "6923450657713";
                            string file = System.Environment.CurrentDirectory + "\\qrcodemid_" + i.ToString("D4") + ".jpg";
                            ZXing.ZXingHelper.EnCode_QR_CODE_MID(enstr, file, pngpath, 300, 300, 2);
                            LogHelper.LogInfo("The " + i.ToString() + " QR_CODE_MID Has EnCoded! JPG Save Path:" + file);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < row; i++)
                    {
                        string enstr = rg.Cells[i + 1, 1].Text;
                        if (enstr.Length == 0) enstr = "6923450657713";

                        switch (encode_format)
                        {
                        case 1:
                            {
                                string file = System.Environment.CurrentDirectory + "\\ean13_" + i.ToString("D4") + ".jpg";
                                ZXing.ZXingHelper.EnCode_EAN_13(enstr, file, 300, 100, 2);
                                LogHelper.LogInfo("The " + i.ToString() + " EAN_13 Has EnCoded! JPG Save Path:" + file);
                            }
                            break;
                        case 2:
                            {
                                string file = System.Environment.CurrentDirectory + "\\pdf417_" + i.ToString("D4") + ".jpg";
                                ZXing.ZXingHelper.EnCode_PDF_417(enstr, file, 300, 100, 2);
                                LogHelper.LogInfo("The " + i.ToString() + " PDF_417 Has EnCoded! JPG Save Path:" + file);
                            }
                            break;
                        case 3:
                            {
                                string file = System.Environment.CurrentDirectory + "\\qrcode_" + i.ToString("D4") + ".jpg";
                                ZXing.ZXingHelper.EnCode_QR_CODE(enstr, file, 300, 300, 2);
                                LogHelper.LogInfo("The " + i.ToString() + " QR_CODE Has EnCoded! JPG Save Path:" + file);
                            }
                            break;
                        }
                    }
                }

                LogHelper.LogInfo("End EnCode Info Successful!");
                System.Diagnostics.Process.Start(System.Environment.CurrentDirectory);
            }
            catch (Exception ee)
            { MessageBox.Show(ee.Message); }
        }

        public string ui_getZxingLabel(IRibbonControl control)
        {
            switch (encode_format)
            {
                case 1:
                    return "普通条型码";
                case 2:
                    return "条形二维码";
                case 3:
                    return "普通二维码";
                case 4:
                    return "图标二维码";
                default:
                    return "编码类型";
            }
        }

        public void ui_bnEAN13(IRibbonControl control)
        {
            encode_format = 1;
            ui_ribbon.Invalidate();
        }

        public void ui_bnPDF417(IRibbonControl control)
        {
            encode_format = 2;
            ui_ribbon.Invalidate();
        }

        public void ui_bnQRCode(IRibbonControl control)
        {
            encode_format = 3;
            ui_ribbon.Invalidate();
        }

        public void ui_bnQRCodeMid(IRibbonControl control)
        {
            encode_format = 4;
            ui_ribbon.Invalidate();
        }

        public void ui_bnExitExcel(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            xlApp.Quit();
        }

        public static void RunNet()
        {
            System.Diagnostics.Process.Start("http://xidong.net/");
        }

        public void ui_bnReg(IRibbonControl control)
        {
            frmReg frm = new frmReg();
            frm.ShowDialog();
        }

        public void ui_bnHelp(IRibbonControl control)
        {
            string path = ExcelDnaUtil.XllPath;
            path = path.Substring(0, path.LastIndexOf('\\'));
            System.Diagnostics.Process.Start(path + "\\xlExcel\\help.pdf");
        }

        [ExcelCommand()]
        public static void xl_AddFaceId(int idx)
        {
            CommandBars bars = ExcelCommandBarUtil.GetCommandBars();
            CommandBar bar = bars.Add("FaceId", MsoBarPosition.msoBarFloating);

            CommandBarControls ctls = bars[3].Controls;

            for (int i = 1; i <= ctls.Count; i++)
            {
                CommandBarControl ctl = ctls[i];
                MsoControlType type = ctl.Type;
                if (type == MsoControlType.msoControlButton)
                {
                    CommandBarButton btn = bar.Controls.AddButton();
                    btn.Style = MsoButtonStyle.msoButtonIcon;
                    btn.FaceId = ((CommandBarButton)ctl).FaceId;
                    btn.TooltipText = ((CommandBarButton)ctl).FaceId.ToString();
                }
            }
        }
        public static void xl_DelFaceId()
        {
            CommandBars bars = ExcelCommandBarUtil.GetCommandBars();
            CommandBar bar = bars["FaceId"];
            if (bar != null) bar.Delete();
        }
        private void xl_ExcelCommandBar()
        {
            CommandBars bars = ExcelCommandBarUtil.GetCommandBars();

            //Create CommandBar
            CommandBar bar = bars.Add("Excel-DNA MenuBar", MsoBarPosition.msoBarLeft);
            CommandBarPopup pop = bar.Controls.AddPopup("Excel-DNA Popup");
            pop.Caption = "Excel-DNA MenuBar";
            CommandBarButton btn = pop.Controls.AddButton();
            btn.Caption = "AddWorksheet";
            btn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            btn.FaceId = 5919;
            btn.OnAction = "AddWorksheet";
            btn = pop.Controls.AddButton();
            btn.Caption = "TestXlCallCmd";
            btn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            btn.FaceId = 548;
            btn.OnAction = "TestXlCallCmd";
            btn = pop.Controls.AddButton();
            btn.Caption = "TestDataFrame";
            btn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            btn.FaceId = 459;
            btn.OnAction = "TestDataFrame";
            btn = pop.Controls.AddButton();
            btn.Caption = "TestROdbdcSql";
            btn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            btn.FaceId = 987;
            btn.OnAction = "TestROdbdcSql";

            //Get Cell CommandBar
            bar = bars["cell"];
            btn = bar.Controls.AddButton();
            btn.Caption = "ExcelVersion";
            btn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            btn.FaceId = 3631;
            btn.OnAction = "ExcelVersion";

            bar = bars.Add("Excel-DNA ToolBar", MsoBarPosition.msoBarFloating);
            CommandBarComboBox cbx = bar.Controls.AddComboBox();
            for (int i = 1; i <= bars.Count; i++) cbx.AddItem(bars[i].NameLocal);
            cbx.ListIndex = 1;
            cbx.Style = MsoComboStyle.msoComboLabel;
            cbx.Caption = "Bars";
            cbx.OnAction = "xl_AddFaceId";
            btn = bar.Controls.AddButton();
            btn.Style = MsoButtonStyle.msoButtonIcon;
            btn.FaceId = 3631;
            btn.TooltipText = "Add FaceId Button";
            btn.OnAction = "xl_AddFaceId";
            btn = bar.Controls.AddButton();
            btn.Style = MsoButtonStyle.msoButtonIcon;
            btn.FaceId = 47;
            btn.TooltipText = "Del FaceId Button";
            btn.OnAction = "xl_DelFaceId";
        }
        public void ui_bnAddBar(IRibbonControl control)
        {
            xl_ExcelCommandBar();
        }
    }

    [ComVisible(true)]
    public class CTPManager
    {
        static CustomTaskPane ctp;

        public static bool Visible { get { if (ctp != null) return ctp.Visible; return false; } }

        public static void OpenCTP()
        {
            if (ctp == null)
            {
                // Make a new one using ExcelDna.Integration.CustomUI.CustomTaskPaneFactory 
                ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(CTPControl), "ExcelDna Task Pane");
                ctp.Visible = true;
                ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                ctp.DockPositionStateChange += ctp_DockPositionStateChange;
                ctp.VisibleStateChange += ctp_VisibleStateChange;
            }
            else
            {
                // Just show it again
                ctp.Visible = true;
            }
        }

        public static void CloseCTP()
        {
            if (ctp != null)
            {
                // Could hide instead, by calling ctp.Visible = false;
                ctp.Delete();
                ctp = null;
            }
        }

        static void ctp_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            LogDisplay.WriteLine("CTP Visibility: {0}", CustomTaskPaneInst.Visible);
        }

        static void ctp_DockPositionStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            LogDisplay.WriteLine("CTP DockPosition: {0}", CustomTaskPaneInst.DockPosition.ToString());
        }
    }
}