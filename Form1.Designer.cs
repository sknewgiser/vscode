namespace ArcGISTools
{
    partial class Form1
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnSure = new System.Windows.Forms.Button();
            this.btnTemplatePath = new System.Windows.Forms.Button();
            this.btnSrcMxdPath = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.cbxExpention = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtResolution = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnSavePath = new System.Windows.Forms.Button();
            this.btnMxdPath = new System.Windows.Forms.Button();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.groupBox1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.groupBox2);
            this.splitContainer1.Size = new System.Drawing.Size(635, 431);
            this.splitContainer1.SplitterDistance = 169;
            this.splitContainer1.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.btnSure);
            this.groupBox1.Controls.Add(this.btnTemplatePath);
            this.groupBox1.Controls.Add(this.btnSrcMxdPath);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(635, 169);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "转换模板";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(181, 15);
            this.label2.TabIndex = 9;
            this.label2.Text = "请输入模板mxd文件夹路径";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(166, 15);
            this.label1.TabIndex = 10;
            this.label1.Text = "请输入源mxd文件夹路径";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(209, 82);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(323, 25);
            this.textBox2.TabIndex = 8;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(209, 26);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(323, 25);
            this.textBox1.TabIndex = 7;
            // 
            // btnSure
            // 
            this.btnSure.Location = new System.Drawing.Point(470, 128);
            this.btnSure.Name = "btnSure";
            this.btnSure.Size = new System.Drawing.Size(75, 30);
            this.btnSure.TabIndex = 4;
            this.btnSure.Text = "执行";
            this.btnSure.UseVisualStyleBackColor = true;
            this.btnSure.Click += new System.EventHandler(this.btnSure_Click);
            // 
            // btnTemplatePath
            // 
            this.btnTemplatePath.Location = new System.Drawing.Point(538, 82);
            this.btnTemplatePath.Name = "btnTemplatePath";
            this.btnTemplatePath.Size = new System.Drawing.Size(75, 25);
            this.btnTemplatePath.TabIndex = 5;
            this.btnTemplatePath.Text = "...";
            this.btnTemplatePath.UseVisualStyleBackColor = true;
            this.btnTemplatePath.Click += new System.EventHandler(this.btnTemplatePath_Click);
            // 
            // btnSrcMxdPath
            // 
            this.btnSrcMxdPath.Location = new System.Drawing.Point(538, 26);
            this.btnSrcMxdPath.Name = "btnSrcMxdPath";
            this.btnSrcMxdPath.Size = new System.Drawing.Size(75, 25);
            this.btnSrcMxdPath.TabIndex = 6;
            this.btnSrcMxdPath.Text = "...";
            this.btnSrcMxdPath.UseVisualStyleBackColor = true;
            this.btnSrcMxdPath.Click += new System.EventHandler(this.btnSrcMxdPath_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.cbxExpention);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.txtResolution);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.textBox3);
            this.groupBox2.Controls.Add(this.textBox4);
            this.groupBox2.Controls.Add(this.btnSave);
            this.groupBox2.Controls.Add(this.btnSavePath);
            this.groupBox2.Controls.Add(this.btnMxdPath);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(635, 258);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "导出地图";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(535, 136);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(31, 15);
            this.label7.TabIndex = 22;
            this.label7.Text = "dpi";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(374, 136);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 15);
            this.label6.TabIndex = 21;
            this.label6.Text = "分辨率";
            // 
            // cbxExpention
            // 
            this.cbxExpention.FormattingEnabled = true;
            this.cbxExpention.Items.AddRange(new object[] {
            "JPG",
            "PNG",
            "TIFF",
            "PDF",
            "BMP",
            "SVG"});
            this.cbxExpention.Location = new System.Drawing.Point(209, 130);
            this.cbxExpention.Name = "cbxExpention";
            this.cbxExpention.Size = new System.Drawing.Size(84, 23);
            this.cbxExpention.TabIndex = 20;
            this.cbxExpention.Text = "JPG";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(128, 139);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(67, 15);
            this.label5.TabIndex = 19;
            this.label5.Text = "保存类型";
            // 
            // txtResolution
            // 
            this.txtResolution.Location = new System.Drawing.Point(432, 128);
            this.txtResolution.Name = "txtResolution";
            this.txtResolution.Size = new System.Drawing.Size(100, 25);
            this.txtResolution.TabIndex = 18;
            this.txtResolution.Text = "300";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(87, 83);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 15);
            this.label3.TabIndex = 16;
            this.label3.Text = "请输入保存路径";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(49, 30);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(151, 15);
            this.label4.TabIndex = 17;
            this.label4.Text = "请输入mxd文件夹路径";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(209, 80);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(323, 25);
            this.textBox3.TabIndex = 15;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(209, 24);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(323, 25);
            this.textBox4.TabIndex = 14;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(470, 203);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 29);
            this.btnSave.TabIndex = 11;
            this.btnSave.Text = "导出";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnSavePath
            // 
            this.btnSavePath.Location = new System.Drawing.Point(538, 79);
            this.btnSavePath.Name = "btnSavePath";
            this.btnSavePath.Size = new System.Drawing.Size(75, 26);
            this.btnSavePath.TabIndex = 12;
            this.btnSavePath.Text = "...";
            this.btnSavePath.UseVisualStyleBackColor = true;
            this.btnSavePath.Click += new System.EventHandler(this.btnSavePath_Click);
            // 
            // btnMxdPath
            // 
            this.btnMxdPath.Location = new System.Drawing.Point(538, 24);
            this.btnMxdPath.Name = "btnMxdPath";
            this.btnMxdPath.Size = new System.Drawing.Size(75, 25);
            this.btnMxdPath.TabIndex = 13;
            this.btnMxdPath.Text = "...";
            this.btnMxdPath.UseVisualStyleBackColor = true;
            this.btnMxdPath.Click += new System.EventHandler(this.btnMxdPath_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(635, 431);
            this.Controls.Add(this.splitContainer1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnSure;
        private System.Windows.Forms.Button btnTemplatePath;
        private System.Windows.Forms.Button btnSrcMxdPath;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnSavePath;
        private System.Windows.Forms.Button btnMxdPath;
        private System.Windows.Forms.TextBox txtResolution;
        private System.Windows.Forms.ComboBox cbxExpention;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
    }
}

