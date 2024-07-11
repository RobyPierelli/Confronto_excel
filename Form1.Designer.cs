namespace Excel_Baxter
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnCreateExcel = new Button();
            dataGridView = new DataGridView();
            btnReadExcel = new Button();
            openFileDialog1 = new OpenFileDialog();
            confronto = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            SuspendLayout();
            // 
            // btnCreateExcel
            // 
            btnCreateExcel.Location = new Point(12, 12);
            btnCreateExcel.Name = "btnCreateExcel";
            btnCreateExcel.Size = new Size(157, 33);
            btnCreateExcel.TabIndex = 0;
            btnCreateExcel.Text = "crea nuovo Excel";
            btnCreateExcel.UseVisualStyleBackColor = true;
            btnCreateExcel.Click += btnCreateExcel_Click;
            // 
            // dataGridView
            // 
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Location = new Point(175, 12);
            dataGridView.Name = "dataGridView";
            dataGridView.Size = new Size(604, 426);
            dataGridView.TabIndex = 1;
            dataGridView.CellContentClick += dataGridView1_CellContentClick;
            // 
            // btnReadExcel
            // 
            btnReadExcel.Location = new Point(12, 51);
            btnReadExcel.Name = "btnReadExcel";
            btnReadExcel.Size = new Size(157, 33);
            btnReadExcel.TabIndex = 2;
            btnReadExcel.Text = "Leggi Excel";
            btnReadExcel.UseVisualStyleBackColor = true;
            btnReadExcel.Click += btnReadExcel_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // confronto
            // 
            confronto.Location = new Point(19, 122);
            confronto.Name = "confronto";
            confronto.Size = new Size(139, 41);
            confronto.TabIndex = 3;
            confronto.Text = "Confronta boost jde";
            confronto.UseVisualStyleBackColor = true;
            confronto.Click += confronto_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(confronto);
            Controls.Add(btnReadExcel);
            Controls.Add(dataGridView);
            Controls.Add(btnCreateExcel);
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private Button btnCreateExcel;
        private DataGridView dataGridView;
        private Button btnReadExcel;
        private OpenFileDialog openFileDialog1;
        private Button confronto;
    }
}
