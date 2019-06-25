namespace CSharpAddIn
{
    partial class frmSpotLight
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnRow = new System.Windows.Forms.Button();
            this.btnCol = new System.Windows.Forms.Button();
            this.ndTransparent = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.ndTransparent)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "行标颜色值";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "列标颜色值i";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "颜色透明度";
            // 
            // btnRow
            // 
            this.btnRow.Location = new System.Drawing.Point(99, 13);
            this.btnRow.Name = "btnRow";
            this.btnRow.Size = new System.Drawing.Size(86, 23);
            this.btnRow.TabIndex = 3;
            this.btnRow.Text = "颜色设置";
            this.btnRow.UseVisualStyleBackColor = true;
            this.btnRow.Click += new System.EventHandler(this.btnRow_Click);
            // 
            // btnCol
            // 
            this.btnCol.Location = new System.Drawing.Point(99, 45);
            this.btnCol.Name = "btnCol";
            this.btnCol.Size = new System.Drawing.Size(86, 23);
            this.btnCol.TabIndex = 4;
            this.btnCol.Text = "颜色设置";
            this.btnCol.UseVisualStyleBackColor = true;
            this.btnCol.Click += new System.EventHandler(this.btnCol_Click);
            // 
            // ndTransparent
            // 
            this.ndTransparent.Location = new System.Drawing.Point(99, 78);
            this.ndTransparent.Name = "ndTransparent";
            this.ndTransparent.Size = new System.Drawing.Size(86, 21);
            this.ndTransparent.TabIndex = 5;
            this.ndTransparent.Value = new decimal(new int[] {
            60,
            0,
            0,
            0});
            this.ndTransparent.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);
            // 
            // frmSpotLight
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(199, 109);
            this.Controls.Add(this.ndTransparent);
            this.Controls.Add(this.btnCol);
            this.Controls.Add(this.btnRow);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmSpotLight";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "聚光灯设置";
            ((System.ComponentModel.ISupportInitialize)(this.ndTransparent)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnRow;
        private System.Windows.Forms.Button btnCol;
        private System.Windows.Forms.NumericUpDown ndTransparent;
    }
}