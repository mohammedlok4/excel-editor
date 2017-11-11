namespace ExcelTester
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
            this.tb_inputFile = new System.Windows.Forms.TextBox();
            this.btn_Browse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cb_removeFormulas = new System.Windows.Forms.CheckBox();
            this.btn_doWork = new System.Windows.Forms.Button();
            this.rb_logReports = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // tb_inputFile
            // 
            this.tb_inputFile.Location = new System.Drawing.Point(12, 26);
            this.tb_inputFile.Name = "tb_inputFile";
            this.tb_inputFile.Size = new System.Drawing.Size(546, 20);
            this.tb_inputFile.TabIndex = 0;
            // 
            // btn_Browse
            // 
            this.btn_Browse.Location = new System.Drawing.Point(564, 24);
            this.btn_Browse.Name = "btn_Browse";
            this.btn_Browse.Size = new System.Drawing.Size(75, 23);
            this.btn_Browse.TabIndex = 1;
            this.btn_Browse.Text = "Browse";
            this.btn_Browse.UseVisualStyleBackColor = true;
            this.btn_Browse.Click += new System.EventHandler(this.btn_Browse_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select input file:";
            // 
            // cb_removeFormulas
            // 
            this.cb_removeFormulas.AutoSize = true;
            this.cb_removeFormulas.Location = new System.Drawing.Point(12, 71);
            this.cb_removeFormulas.Name = "cb_removeFormulas";
            this.cb_removeFormulas.Size = new System.Drawing.Size(223, 17);
            this.cb_removeFormulas.TabIndex = 3;
            this.cb_removeFormulas.Text = "Remove formulas and replace with values";
            this.cb_removeFormulas.UseVisualStyleBackColor = true;
            // 
            // btn_doWork
            // 
            this.btn_doWork.Location = new System.Drawing.Point(562, 229);
            this.btn_doWork.Name = "btn_doWork";
            this.btn_doWork.Size = new System.Drawing.Size(75, 23);
            this.btn_doWork.TabIndex = 4;
            this.btn_doWork.Text = "Do work";
            this.btn_doWork.UseVisualStyleBackColor = true;
            this.btn_doWork.Click += new System.EventHandler(this.btn_doWork_Click);
            // 
            // rb_logReports
            // 
            this.rb_logReports.AutoSize = true;
            this.rb_logReports.Location = new System.Drawing.Point(465, 232);
            this.rb_logReports.Name = "rb_logReports";
            this.rb_logReports.Size = new System.Drawing.Size(78, 17);
            this.rb_logReports.TabIndex = 5;
            this.rb_logReports.TabStop = true;
            this.rb_logReports.Text = "Log reports";
            this.rb_logReports.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(649, 264);
            this.Controls.Add(this.rb_logReports);
            this.Controls.Add(this.btn_doWork);
            this.Controls.Add(this.cb_removeFormulas);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_Browse);
            this.Controls.Add(this.tb_inputFile);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tb_inputFile;
        private System.Windows.Forms.Button btn_Browse;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox cb_removeFormulas;
        private System.Windows.Forms.Button btn_doWork;
        private System.Windows.Forms.RadioButton rb_logReports;
    }
}

