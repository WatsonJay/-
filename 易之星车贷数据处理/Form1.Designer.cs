namespace 易之星车贷数据处理
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.start_check = new System.Windows.Forms.Button();
            this.file_out = new System.Windows.Forms.Button();
            this.file_three = new System.Windows.Forms.Button();
            this.file_two = new System.Windows.Forms.Button();
            this.file_one = new System.Windows.Forms.Button();
            this.file_four_name = new System.Windows.Forms.TextBox();
            this.file_three_name = new System.Windows.Forms.TextBox();
            this.file_two_name = new System.Windows.Forms.TextBox();
            this.file_one_name = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            // 
            // start_check
            // 
            this.start_check.Location = new System.Drawing.Point(231, 189);
            this.start_check.Name = "start_check";
            this.start_check.Size = new System.Drawing.Size(163, 38);
            this.start_check.TabIndex = 25;
            this.start_check.Text = "数据校验";
            this.start_check.UseVisualStyleBackColor = true;
            this.start_check.Click += new System.EventHandler(this.start_check_Click);
            // 
            // file_out
            // 
            this.file_out.Location = new System.Drawing.Point(544, 138);
            this.file_out.Name = "file_out";
            this.file_out.Size = new System.Drawing.Size(75, 23);
            this.file_out.TabIndex = 24;
            this.file_out.Text = "浏览...";
            this.file_out.UseVisualStyleBackColor = true;
            this.file_out.Click += new System.EventHandler(this.file_out_Click);
            // 
            // file_three
            // 
            this.file_three.Location = new System.Drawing.Point(544, 102);
            this.file_three.Name = "file_three";
            this.file_three.Size = new System.Drawing.Size(75, 23);
            this.file_three.TabIndex = 23;
            this.file_three.Text = "浏览...";
            this.file_three.UseVisualStyleBackColor = true;
            this.file_three.Click += new System.EventHandler(this.file_three_Click);
            // 
            // file_two
            // 
            this.file_two.Location = new System.Drawing.Point(544, 67);
            this.file_two.Name = "file_two";
            this.file_two.Size = new System.Drawing.Size(75, 23);
            this.file_two.TabIndex = 22;
            this.file_two.Text = "浏览...";
            this.file_two.UseVisualStyleBackColor = true;
            this.file_two.Click += new System.EventHandler(this.file_two_Click);
            // 
            // file_one
            // 
            this.file_one.Location = new System.Drawing.Point(544, 31);
            this.file_one.Name = "file_one";
            this.file_one.Size = new System.Drawing.Size(75, 23);
            this.file_one.TabIndex = 21;
            this.file_one.Text = "浏览...";
            this.file_one.UseVisualStyleBackColor = true;
            this.file_one.Click += new System.EventHandler(this.file_one_Click);
            // 
            // file_four_name
            // 
            this.file_four_name.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.file_four_name.Location = new System.Drawing.Point(157, 137);
            this.file_four_name.Name = "file_four_name";
            this.file_four_name.Size = new System.Drawing.Size(372, 26);
            this.file_four_name.TabIndex = 20;
            // 
            // file_three_name
            // 
            this.file_three_name.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.file_three_name.Location = new System.Drawing.Point(157, 101);
            this.file_three_name.Name = "file_three_name";
            this.file_three_name.Size = new System.Drawing.Size(372, 26);
            this.file_three_name.TabIndex = 19;
            // 
            // file_two_name
            // 
            this.file_two_name.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.file_two_name.Location = new System.Drawing.Point(157, 66);
            this.file_two_name.Name = "file_two_name";
            this.file_two_name.Size = new System.Drawing.Size(372, 26);
            this.file_two_name.TabIndex = 18;
            // 
            // file_one_name
            // 
            this.file_one_name.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.file_one_name.Location = new System.Drawing.Point(157, 30);
            this.file_one_name.Name = "file_one_name";
            this.file_one_name.Size = new System.Drawing.Size(372, 26);
            this.file_one_name.TabIndex = 17;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(24, 139);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(112, 16);
            this.label4.TabIndex = 16;
            this.label4.Text = "表4：校验结果";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(24, 103);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 16);
            this.label3.TabIndex = 15;
            this.label3.Text = "表3：报账数据";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(24, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(128, 16);
            this.label2.TabIndex = 14;
            this.label2.Text = "表2：服务费标准";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(24, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 16);
            this.label1.TabIndex = 13;
            this.label1.Text = "表1：业务总表";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(643, 256);
            this.Controls.Add(this.start_check);
            this.Controls.Add(this.file_out);
            this.Controls.Add(this.file_three);
            this.Controls.Add(this.file_two);
            this.Controls.Add(this.file_one);
            this.Controls.Add(this.file_four_name);
            this.Controls.Add(this.file_three_name);
            this.Controls.Add(this.file_two_name);
            this.Controls.Add(this.file_one_name);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "报账数据校验";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button start_check;
        private System.Windows.Forms.Button file_out;
        private System.Windows.Forms.Button file_three;
        private System.Windows.Forms.Button file_two;
        private System.Windows.Forms.Button file_one;
        private System.Windows.Forms.TextBox file_four_name;
        private System.Windows.Forms.TextBox file_three_name;
        private System.Windows.Forms.TextBox file_two_name;
        private System.Windows.Forms.TextBox file_one_name;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;

    }
}

