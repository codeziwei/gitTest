﻿namespace WindowsFormsApplication1
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnLookfile = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnLookfile2 = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnshowColumnName = new System.Windows.Forms.Button();
            this.txtfileprefix = new System.Windows.Forms.TextBox();
            this.btnFormat = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtMessage = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtgys = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnLookfile
            // 
            this.btnLookfile.Location = new System.Drawing.Point(551, 25);
            this.btnLookfile.Name = "btnLookfile";
            this.btnLookfile.Size = new System.Drawing.Size(80, 24);
            this.btnLookfile.TabIndex = 1;
            this.btnLookfile.Text = "浏览..";
            this.btnLookfile.UseVisualStyleBackColor = true;
            this.btnLookfile.Click += new System.EventHandler(this.btnLookfile_Click);
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(11, 25);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(539, 21);
            this.textBox1.TabIndex = 2;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnLookfile2
            // 
            this.btnLookfile2.Location = new System.Drawing.Point(556, 23);
            this.btnLookfile2.Name = "btnLookfile2";
            this.btnLookfile2.Size = new System.Drawing.Size(75, 29);
            this.btnLookfile2.TabIndex = 1;
            this.btnLookfile2.Text = "浏览..";
            this.btnLookfile2.UseVisualStyleBackColor = true;
            this.btnLookfile2.Click += new System.EventHandler(this.btnLookfile2_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(11, 28);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(546, 21);
            this.textBox2.TabIndex = 2;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.btnLookfile);
            this.groupBox1.Location = new System.Drawing.Point(21, 20);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(679, 68);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "模版文件路径";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnshowColumnName);
            this.groupBox2.Controls.Add(this.txtgys);
            this.groupBox2.Controls.Add(this.txtfileprefix);
            this.groupBox2.Controls.Add(this.btnFormat);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.textBox2);
            this.groupBox2.Controls.Add(this.btnLookfile2);
            this.groupBox2.Location = new System.Drawing.Point(21, 112);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(679, 168);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "需要格式化的Excel文件所在文件夹";
            // 
            // btnshowColumnName
            // 
            this.btnshowColumnName.Location = new System.Drawing.Point(412, 70);
            this.btnshowColumnName.Name = "btnshowColumnName";
            this.btnshowColumnName.Size = new System.Drawing.Size(75, 31);
            this.btnshowColumnName.TabIndex = 6;
            this.btnshowColumnName.Text = "查看列名";
            this.btnshowColumnName.UseVisualStyleBackColor = true;
            this.btnshowColumnName.Click += new System.EventHandler(this.btnshowColumnName_Click);
            // 
            // txtfileprefix
            // 
            this.txtfileprefix.Location = new System.Drawing.Point(104, 71);
            this.txtfileprefix.Name = "txtfileprefix";
            this.txtfileprefix.Size = new System.Drawing.Size(100, 21);
            this.txtfileprefix.TabIndex = 4;
            this.txtfileprefix.Text = "格式化后_";
            // 
            // btnFormat
            // 
            this.btnFormat.Location = new System.Drawing.Point(556, 70);
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Size = new System.Drawing.Size(75, 31);
            this.btnFormat.TabIndex = 5;
            this.btnFormat.Text = "开始格式化";
            this.btnFormat.UseVisualStyleBackColor = true;
            this.btnFormat.Click += new System.EventHandler(this.btnFormat_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 74);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "格式化文件前缀";
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(21, 310);
            this.txtMessage.Multiline = true;
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtMessage.Size = new System.Drawing.Size(679, 339);
            this.txtMessage.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 108);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "供应商";
            // 
            // txtgys
            // 
            this.txtgys.Location = new System.Drawing.Point(104, 105);
            this.txtgys.Name = "txtgys";
            this.txtgys.Size = new System.Drawing.Size(100, 21);
            this.txtgys.TabIndex = 4;
            this.txtgys.Text = "供应商";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(734, 737);
            this.Controls.Add(this.txtMessage);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.MaximumSize = new System.Drawing.Size(750, 1600);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "中教育人-Excel格式化";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLookfile;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnLookfile2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnFormat;
        private System.Windows.Forms.TextBox txtMessage;
        private System.Windows.Forms.TextBox txtfileprefix;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnshowColumnName;
        private System.Windows.Forms.TextBox txtgys;
        private System.Windows.Forms.Label label2;
    }
}

