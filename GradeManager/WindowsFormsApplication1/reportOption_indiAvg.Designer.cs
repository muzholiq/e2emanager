namespace WindowsFormsApplication1
{
    partial class reportOption_indiAvg
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
            this.textBox_OptAvgDurationStart = new System.Windows.Forms.TextBox();
            this.textBox_OptAvgDurationEnd = new System.Windows.Forms.TextBox();
            this.textBox_OptAvgMax = new System.Windows.Forms.TextBox();
            this.textBox_OptAvgMin = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.button_OptAvgSummit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox_OptAvgDurationStart
            // 
            this.textBox_OptAvgDurationStart.Location = new System.Drawing.Point(119, 20);
            this.textBox_OptAvgDurationStart.Name = "textBox_OptAvgDurationStart";
            this.textBox_OptAvgDurationStart.Size = new System.Drawing.Size(82, 21);
            this.textBox_OptAvgDurationStart.TabIndex = 0;
            // 
            // textBox_OptAvgDurationEnd
            // 
            this.textBox_OptAvgDurationEnd.Location = new System.Drawing.Point(247, 20);
            this.textBox_OptAvgDurationEnd.Name = "textBox_OptAvgDurationEnd";
            this.textBox_OptAvgDurationEnd.Size = new System.Drawing.Size(82, 21);
            this.textBox_OptAvgDurationEnd.TabIndex = 1;
            // 
            // textBox_OptAvgMax
            // 
            this.textBox_OptAvgMax.Location = new System.Drawing.Point(247, 49);
            this.textBox_OptAvgMax.Name = "textBox_OptAvgMax";
            this.textBox_OptAvgMax.Size = new System.Drawing.Size(82, 21);
            this.textBox_OptAvgMax.TabIndex = 3;
            // 
            // textBox_OptAvgMin
            // 
            this.textBox_OptAvgMin.Location = new System.Drawing.Point(119, 49);
            this.textBox_OptAvgMin.Name = "textBox_OptAvgMin";
            this.textBox_OptAvgMin.Size = new System.Drawing.Size(82, 21);
            this.textBox_OptAvgMin.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "기간";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(213, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(15, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "to";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(83, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "from";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(83, 52);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 12);
            this.label4.TabIndex = 8;
            this.label4.Text = "from";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(213, 52);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(15, 12);
            this.label5.TabIndex = 7;
            this.label5.Text = "to";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 49);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 12);
            this.label6.TabIndex = 9;
            this.label6.Text = "평균범위";
            // 
            // button_OptAvgSummit
            // 
            this.button_OptAvgSummit.Location = new System.Drawing.Point(345, 22);
            this.button_OptAvgSummit.Name = "button_OptAvgSummit";
            this.button_OptAvgSummit.Size = new System.Drawing.Size(102, 41);
            this.button_OptAvgSummit.TabIndex = 10;
            this.button_OptAvgSummit.Text = "summit";
            this.button_OptAvgSummit.UseVisualStyleBackColor = true;
            this.button_OptAvgSummit.Click += new System.EventHandler(this.button_OptAvgSummit_Click);
            // 
            // reportOption_indiAvg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(463, 91);
            this.Controls.Add(this.button_OptAvgSummit);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox_OptAvgMax);
            this.Controls.Add(this.textBox_OptAvgMin);
            this.Controls.Add(this.textBox_OptAvgDurationEnd);
            this.Controls.Add(this.textBox_OptAvgDurationStart);
            this.Name = "reportOption_indiAvg";
            this.Text = "reportOption";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_OptAvgDurationStart;
        private System.Windows.Forms.TextBox textBox_OptAvgDurationEnd;
        private System.Windows.Forms.TextBox textBox_OptAvgMax;
        private System.Windows.Forms.TextBox textBox_OptAvgMin;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button_OptAvgSummit;
    }
}