namespace InteropC
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnWord = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.txtDato = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnWord
            // 
            this.btnWord.Location = new System.Drawing.Point(68, 127);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(75, 23);
            this.btnWord.TabIndex = 0;
            this.btnWord.Text = "button1";
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWord_Click_1);
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(160, 127);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(75, 23);
            this.btnExcel.TabIndex = 1;
            this.btnExcel.Text = "button2";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click_1);
            // 
            // txtDato
            // 
            this.txtDato.Location = new System.Drawing.Point(68, 70);
            this.txtDato.Name = "txtDato";
            this.txtDato.Size = new System.Drawing.Size(167, 20);
            this.txtDato.TabIndex = 2;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(314, 222);
            this.Controls.Add(this.txtDato);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnWord);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.TextBox txtDato;
    }
}

