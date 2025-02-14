namespace TeklaMaterialList
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        private void InitializeComponent()
        {
            this.btnGenerate = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(12, 12);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(260, 37);
            this.btnGenerate.Text = "Generate Material List";
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 50); // Adjust position as needed
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(776, 23);
            this.progressBar1.TabIndex = 1;
            this.progressBar1.Visible = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 61);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.progressBar1);
            this.Name = "MainForm";
            this.Text = "Tekla Material List";
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}
