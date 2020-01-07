namespace WindowsFormsApp5
{
    partial class Students
    {
        /// <summary> 
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary> 
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBoxFamaly = new System.Windows.Forms.TextBox();
            this.textBoxName = new System.Windows.Forms.TextBox();
            this.textBoxLastName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // textBoxFamaly
            // 
            this.textBoxFamaly.Location = new System.Drawing.Point(22, 4);
            this.textBoxFamaly.Name = "textBoxFamaly";
            this.textBoxFamaly.Size = new System.Drawing.Size(133, 20);
            this.textBoxFamaly.TabIndex = 3;
            this.textBoxFamaly.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // textBoxName
            // 
            this.textBoxName.Location = new System.Drawing.Point(161, 4);
            this.textBoxName.Name = "textBoxName";
            this.textBoxName.Size = new System.Drawing.Size(133, 20);
            this.textBoxName.TabIndex = 4;
            this.textBoxName.TextChanged += new System.EventHandler(this.textBoxName_TextChanged);
            // 
            // textBoxLastName
            // 
            this.textBoxLastName.Location = new System.Drawing.Point(300, 4);
            this.textBoxLastName.Name = "textBoxLastName";
            this.textBoxLastName.Size = new System.Drawing.Size(133, 20);
            this.textBoxLastName.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 7);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 13);
            this.label4.TabIndex = 6;
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // Students
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBoxLastName);
            this.Controls.Add(this.textBoxName);
            this.Controls.Add(this.textBoxFamaly);
            this.Name = "Students";
            this.Size = new System.Drawing.Size(441, 30);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        public System.Windows.Forms.TextBox textBoxFamaly;
        public System.Windows.Forms.TextBox textBoxName;
        public System.Windows.Forms.TextBox textBoxLastName;
        public System.Windows.Forms.Label label4;
    }
}
