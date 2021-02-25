namespace TestStackWhiteDemo
{
    partial class Form2
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
            this.form2_btn_close = new System.Windows.Forms.Button();
            this.form2_lbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // form2_btn_close
            // 
            this.form2_btn_close.Location = new System.Drawing.Point(173, 173);
            this.form2_btn_close.Name = "form2_btn_close";
            this.form2_btn_close.Size = new System.Drawing.Size(75, 23);
            this.form2_btn_close.TabIndex = 0;
            this.form2_btn_close.Text = "Close";
            this.form2_btn_close.UseVisualStyleBackColor = true;
            this.form2_btn_close.Click += new System.EventHandler(this.form2_btn_close_Click);
            // 
            // form2_lbl
            // 
            this.form2_lbl.AutoSize = true;
            this.form2_lbl.Location = new System.Drawing.Point(85, 94);
            this.form2_lbl.Name = "form2_lbl";
            this.form2_lbl.Size = new System.Drawing.Size(55, 13);
            this.form2_lbl.TabIndex = 1;
            this.form2_lbl.Text = "Welcome ";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.form2_lbl);
            this.Controls.Add(this.form2_btn_close);
            this.Name = "Form2";
            this.Text = "Welcome";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button form2_btn_close;
        private System.Windows.Forms.Label form2_lbl;
    }
}