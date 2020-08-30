namespace ExcelReader
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
            this.components = new System.ComponentModel.Container();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button_load_file = new System.Windows.Forms.Button();
            this.textBox_filePath = new System.Windows.Forms.TextBox();
            this.button_export_data = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button_load_file
            // 
            this.button_load_file.Location = new System.Drawing.Point(12, 21);
            this.button_load_file.Name = "button_load_file";
            this.button_load_file.Size = new System.Drawing.Size(115, 36);
            this.button_load_file.TabIndex = 0;
            this.button_load_file.Text = "导入文件";
            this.button_load_file.UseVisualStyleBackColor = true;
            this.button_load_file.Click += new System.EventHandler(this.button_load_file_Click);
            // 
            // textBox_filePath
            // 
            this.textBox_filePath.Location = new System.Drawing.Point(133, 29);
            this.textBox_filePath.Name = "textBox_filePath";
            this.textBox_filePath.Size = new System.Drawing.Size(655, 25);
            this.textBox_filePath.TabIndex = 1;
            // 
            // button_export_data
            // 
            this.button_export_data.Enabled = false;
            this.button_export_data.Location = new System.Drawing.Point(12, 306);
            this.button_export_data.Name = "button_export_data";
            this.button_export_data.Size = new System.Drawing.Size(115, 36);
            this.button_export_data.TabIndex = 3;
            this.button_export_data.Text = "导出数据";
            this.button_export_data.UseVisualStyleBackColor = true;
            this.button_export_data.Click += new System.EventHandler(this.button_export_data_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(133, 93);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(655, 23);
            this.progressBar1.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 101);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 15);
            this.label1.TabIndex = 5;
            this.label1.Text = "导入进度:";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "(*.csv)|*.csv|All Files(*.*)|*.*";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.button_export_data);
            this.Controls.Add(this.textBox_filePath);
            this.Controls.Add(this.button_load_file);
            this.Name = "Form1";
            this.Text = "ExcelReader";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button_load_file;
        private System.Windows.Forms.TextBox textBox_filePath;
        private System.Windows.Forms.Button button_export_data;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

