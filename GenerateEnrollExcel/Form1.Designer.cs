namespace GenerateEnrollExcel
{
    partial class Frm_Generator
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_generator = new System.Windows.Forms.Button();
            this.ofD_file = new System.Windows.Forms.OpenFileDialog();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.btn_select = new System.Windows.Forms.Button();
            this.btnAddMobile = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_generator
            // 
            this.btn_generator.Location = new System.Drawing.Point(67, 80);
            this.btn_generator.Name = "btn_generator";
            this.btn_generator.Size = new System.Drawing.Size(75, 30);
            this.btn_generator.TabIndex = 0;
            this.btn_generator.Text = "生 成";
            this.btn_generator.UseVisualStyleBackColor = true;
            this.btn_generator.Click += new System.EventHandler(this.btn_generator_Click);
            // 
            // ofD_file
            // 
            this.ofD_file.FileName = "报名情况5.xlsx";
            this.ofD_file.Filter = "*.xls|*.xlsx";
            // 
            // txtFile
            // 
            this.txtFile.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtFile.Location = new System.Drawing.Point(67, 32);
            this.txtFile.Name = "txtFile";
            this.txtFile.ReadOnly = true;
            this.txtFile.Size = new System.Drawing.Size(312, 27);
            this.txtFile.TabIndex = 1;
            // 
            // btn_select
            // 
            this.btn_select.Location = new System.Drawing.Point(385, 31);
            this.btn_select.Name = "btn_select";
            this.btn_select.Size = new System.Drawing.Size(75, 30);
            this.btn_select.TabIndex = 2;
            this.btn_select.Text = "选择...";
            this.btn_select.UseVisualStyleBackColor = true;
            this.btn_select.Click += new System.EventHandler(this.btn_select_Click);
            // 
            // btnAddMobile
            // 
            this.btnAddMobile.Location = new System.Drawing.Point(304, 80);
            this.btnAddMobile.Name = "btnAddMobile";
            this.btnAddMobile.Size = new System.Drawing.Size(156, 30);
            this.btnAddMobile.TabIndex = 3;
            this.btnAddMobile.Text = "添加号码";
            this.btnAddMobile.UseVisualStyleBackColor = true;
            this.btnAddMobile.Click += new System.EventHandler(this.btnAddMobile_Click);
            // 
            // Frm_Generator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(535, 138);
            this.Controls.Add(this.btnAddMobile);
            this.Controls.Add(this.btn_select);
            this.Controls.Add(this.txtFile);
            this.Controls.Add(this.btn_generator);
            this.MaximizeBox = false;
            this.Name = "Frm_Generator";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "生成点名册";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_generator;
        private System.Windows.Forms.OpenFileDialog ofD_file;
        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.Button btn_select;
        private System.Windows.Forms.Button btnAddMobile;
    }
}

