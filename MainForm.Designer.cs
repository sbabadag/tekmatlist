namespace TeklaMaterialList
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabProfiles;
        private System.Windows.Forms.TabPage tabPlates;
        private System.Windows.Forms.TabPage tabBolts;
        private System.Windows.Forms.TabPage tabRods;
        private System.Windows.Forms.DataGridView gridProfiles;
        private System.Windows.Forms.DataGridView gridPlates;
        private System.Windows.Forms.DataGridView gridBolts;
        private System.Windows.Forms.DataGridView gridRods;
        private System.Windows.Forms.Button btnCalculate;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Panel totalsPanel;
        private System.Windows.Forms.Label lblTotalProfiles;
        private System.Windows.Forms.Label lblTotalPlates;
        private System.Windows.Forms.Label lblTotalBolts;
        private System.Windows.Forms.Label lblTotalRods;

        private void InitializeComponent()
        {
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabProfiles = new System.Windows.Forms.TabPage();
            this.tabPlates = new System.Windows.Forms.TabPage();
            this.tabBolts = new System.Windows.Forms.TabPage();
            this.tabRods = new System.Windows.Forms.TabPage();
            this.gridProfiles = new System.Windows.Forms.DataGridView();
            this.gridPlates = new System.Windows.Forms.DataGridView();
            this.gridBolts = new System.Windows.Forms.DataGridView();
            this.gridRods = new System.Windows.Forms.DataGridView();
            this.btnCalculate = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.totalsPanel = new System.Windows.Forms.Panel();
            this.lblTotalProfiles = new System.Windows.Forms.Label();
            this.lblTotalPlates = new System.Windows.Forms.Label();
            this.lblTotalBolts = new System.Windows.Forms.Label();
            this.lblTotalRods = new System.Windows.Forms.Label();

            // Configure form
            this.Text = "Material List";
            this.Size = new System.Drawing.Size(1024, 768);

            // Configure tab control
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Controls.AddRange(new System.Windows.Forms.Control[] { tabProfiles, tabPlates, tabBolts, tabRods });

            // Configure tabs
            this.tabProfiles.Text = "Profiles";
            this.tabPlates.Text = "Plates";
            this.tabBolts.Text = "Bolts";
            this.tabRods.Text = "Çubuklar";

            // Configure grids
            foreach (System.Windows.Forms.DataGridView grid in new[] { gridProfiles, gridPlates, gridBolts, gridRods })
            {
                grid.Dock = System.Windows.Forms.DockStyle.Fill;
                grid.AllowUserToAddRows = false;
                grid.ReadOnly = true;
                grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                grid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            }

            // Add grids to tabs
            this.tabProfiles.Controls.Add(gridProfiles);
            this.tabPlates.Controls.Add(gridPlates);
            this.tabBolts.Controls.Add(gridBolts);
            this.tabRods.Controls.Add(gridRods);

            // Configure buttons panel
            var buttonPanel = new System.Windows.Forms.Panel
            {
                Height = 50,
                Dock = System.Windows.Forms.DockStyle.Bottom
            };

            // Configure buttons
            this.btnCalculate.Text = "Calculate";
            this.btnCalculate.Width = 100;
            this.btnCalculate.Location = new System.Drawing.Point(10, 10);
            this.btnCalculate.Click += new System.EventHandler(this.btnCalculate_Click);

            this.btnExport.Text = "Export to Excel";
            this.btnExport.Width = 100;
            this.btnExport.Location = new System.Drawing.Point(120, 10);
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);

            // Configure progress bar
            this.progressBar1.Width = 200;
            this.progressBar1.Location = new System.Drawing.Point(230, 10);
            this.progressBar1.Visible = false;

            // Add controls
            buttonPanel.Controls.AddRange(new System.Windows.Forms.Control[] { btnCalculate, btnExport, progressBar1 });
            this.Controls.AddRange(new System.Windows.Forms.Control[] { tabControl, buttonPanel });

            // Configure totals panel
            this.totalsPanel.Height = 30;
            this.totalsPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.totalsPanel.BackColor = System.Drawing.Color.LightGray;

            // Configure total labels
            int labelWidth = 250;
            this.lblTotalProfiles.Width = labelWidth;
            this.lblTotalPlates.Width = labelWidth;
            this.lblTotalBolts.Width = labelWidth;
            this.lblTotalRods.Width = labelWidth;

            this.lblTotalProfiles.Location = new System.Drawing.Point(10, 5);
            this.lblTotalPlates.Location = new System.Drawing.Point(labelWidth + 20, 5);
            this.lblTotalBolts.Location = new System.Drawing.Point(2 * labelWidth + 30, 5);
            this.lblTotalRods.Location = new System.Drawing.Point(3 * labelWidth + 40, 5);

            // Add labels to panel
            this.totalsPanel.Controls.AddRange(new System.Windows.Forms.Control[] { 
                lblTotalProfiles, lblTotalPlates, lblTotalBolts, lblTotalRods 
            });

            // Add panel to form
            this.Controls.Add(totalsPanel);
        }
    }
}
