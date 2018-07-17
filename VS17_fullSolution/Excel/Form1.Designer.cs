namespace Excel2
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnStart = new System.Windows.Forms.Button();
            this.tB_Columns = new System.Windows.Forms.TextBox();
            this.tB_Column_From = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnExcel = new System.Windows.Forms.Button();
            this.tB_Column_To = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStart.Location = new System.Drawing.Point(424, 40);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(160, 73);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "sort it!";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // tB_Columns
            // 
            this.tB_Columns.Location = new System.Drawing.Point(37, 34);
            this.tB_Columns.Name = "tB_Columns";
            this.tB_Columns.Size = new System.Drawing.Size(100, 20);
            this.tB_Columns.TabIndex = 1;
            this.tB_Columns.Text = "10";
            this.tB_Columns.TextChanged += new System.EventHandler(this.tB_Columns_TextChanged);
            // 
            // tB_Column_From
            // 
            this.tB_Column_From.Location = new System.Drawing.Point(40, 87);
            this.tB_Column_From.Name = "tB_Column_From";
            this.tB_Column_From.Size = new System.Drawing.Size(100, 20);
            this.tB_Column_From.TabIndex = 1;
            this.tB_Column_From.Text = "A";
            this.tB_Column_From.TextChanged += new System.EventHandler(this.tB_Column_From_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(37, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Rows";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(37, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(64, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "column from";
            // 
            // btnExcel
            // 
            this.btnExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExcel.Location = new System.Drawing.Point(258, 40);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(160, 73);
            this.btnExcel.TabIndex = 3;
            this.btnExcel.Text = "open excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // tB_Column_To
            // 
            this.tB_Column_To.Location = new System.Drawing.Point(143, 87);
            this.tB_Column_To.Name = "tB_Column_To";
            this.tB_Column_To.Size = new System.Drawing.Size(100, 20);
            this.tB_Column_To.TabIndex = 1;
            this.tB_Column_To.Text = "F";
            this.tB_Column_To.TextChanged += new System.EventHandler(this.tB_Column_To_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(140, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "column to";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(606, 248);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tB_Column_To);
            this.Controls.Add(this.tB_Column_From);
            this.Controls.Add(this.tB_Columns);
            this.Controls.Add(this.btnStart);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.TextBox tB_Columns;
        private System.Windows.Forms.TextBox tB_Column_From;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.TextBox tB_Column_To;
        private System.Windows.Forms.Label label3;
    }
}

