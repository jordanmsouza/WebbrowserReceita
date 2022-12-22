namespace WebBrowserReceita
{
    partial class MainForm
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
            this.btn_copiar = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_copiar
            // 
            this.btn_copiar.Location = new System.Drawing.Point(715, 4);
            this.btn_copiar.Name = "btn_copiar";
            this.btn_copiar.Size = new System.Drawing.Size(75, 23);
            this.btn_copiar.TabIndex = 0;
            this.btn_copiar.Text = "Copiar";
            this.btn_copiar.UseVisualStyleBackColor = true;
            this.btn_copiar.Click += new System.EventHandler(this.btn_copiar_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btn_copiar);
            this.panel1.Location = new System.Drawing.Point(0, 430);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(803, 39);
            this.panel1.TabIndex = 1;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 470);
            this.Controls.Add(this.panel1);
            this.Name = "MainForm";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_copiar;
        private System.Windows.Forms.Panel panel1;
    }
}

