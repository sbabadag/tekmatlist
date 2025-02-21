namespace TeklaMaterialList
{
    using System.Linq; // Add this at the top

    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel totalsPanel;
        private System.Windows.Forms.Label lblTotalProfiles;
        private System.Windows.Forms.Label lblTotalPlates;
        private System.Windows.Forms.Label lblTotalBolts;
        private System.Windows.Forms.Label lblTotalRods;

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();

            // Create standard length combo box first
            this.cmbStandardLength = new System.Windows.Forms.ComboBox();
            this.cmbStandardLength.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbStandardLength.Size = new System.Drawing.Size(80, 23);
            this.cmbStandardLength.Items.AddRange(new object[] { "6 mt", "12 mt" });
            this.cmbStandardLength.SelectedIndex = 1; // Default to 12m
            this.cmbStandardLength.SelectedIndexChanged += new System.EventHandler(CmbStandardLength_SelectedIndexChanged);

            // Add the label for standard length
            var lblStandardLength = new System.Windows.Forms.Label();
            lblStandardLength.Text = "Standard Length:";
            lblStandardLength.AutoSize = true;

            // Create and initialize grids
            this.gridProfiles = new System.Windows.Forms.DataGridView();
            this.gridPlates = new System.Windows.Forms.DataGridView();
            this.gridBolts = new System.Windows.Forms.DataGridView();
            this.gridRods = new System.Windows.Forms.DataGridView();
            this.gridNuts = new System.Windows.Forms.DataGridView();

            // Create tab control and pages
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabProfiles = new System.Windows.Forms.TabPage();
            this.tabPlates = new System.Windows.Forms.TabPage();
            this.tabBolts = new System.Windows.Forms.TabPage();
            this.tabRods = new System.Windows.Forms.TabPage();
            this.tabNuts = new System.Windows.Forms.TabPage();

            // Create buttons and panel
            this.btnCalculate = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.totalsPanel = new System.Windows.Forms.Panel();

            // Initialize labels
            this.lblTotalProfiles = new System.Windows.Forms.Label();
            this.lblTotalPlates = new System.Windows.Forms.Label();
            this.lblTotalBolts = new System.Windows.Forms.Label();
            this.lblTotalRods = new System.Windows.Forms.Label();

            // Set form properties
            this.Text = "Material List";
            this.Size = new System.Drawing.Size(1024, 768);

            // Configure tab control
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);

            // Configure tab pages
            this.tabProfiles.Text = "Profiles";
            this.tabPlates.Text = "Plates";
            this.tabBolts.Text = "Bolts";
            this.tabRods.Text = "Rods";
            this.tabNuts.Text = "Nuts";

            // Configure grids
            foreach (var grid in new[] { gridProfiles, gridPlates, gridBolts, gridRods, gridNuts })
            {
                grid.Dock = System.Windows.Forms.DockStyle.Fill;
                grid.AllowUserToAddRows = false;
                grid.ReadOnly = true;
                grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                grid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
                grid.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            }

            // Add grids to tabs
            this.tabProfiles.Controls.Add(this.gridProfiles);
            this.tabPlates.Controls.Add(this.gridPlates);
            this.tabBolts.Controls.Add(this.gridBolts);
            this.tabRods.Controls.Add(this.gridRods);
            this.tabNuts.Controls.Add(this.gridNuts);

            // Add tabs to tab control
            this.tabControl1.TabPages.AddRange(new[] {
                this.tabProfiles,
                this.tabPlates,
                this.tabBolts,
                this.tabRods,
                this.tabNuts
            });

            // Configure buttons panel
            var buttonPanel = new System.Windows.Forms.Panel
            {
                Dock = System.Windows.Forms.DockStyle.Bottom,
                Height = 50
            };

            // Configure buttons
            this.btnCalculate.Text = "Calculate";
            this.btnCalculate.Size = new System.Drawing.Size(100, 30);
            this.btnCalculate.Location = new System.Drawing.Point(10, 10);
            this.btnCalculate.Click += new System.EventHandler(this.btnCalculate_Click);

            this.btnExport.Text = "Export";
            this.btnExport.Size = new System.Drawing.Size(100, 30);
            this.btnExport.Location = new System.Drawing.Point(120, 10);
            this.btnExport.Enabled = false;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);

            // Configure progress bar
            this.progressBar1.Size = new System.Drawing.Size(200, 20);
            this.progressBar1.Location = new System.Drawing.Point(230, 15);
            this.progressBar1.Visible = false;

            // Configure totals panel
            this.totalsPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.totalsPanel.Height = 30;
            this.totalsPanel.BackColor = System.Drawing.Color.LightGray;

            // Add controls to panels
            buttonPanel.Controls.AddRange(new System.Windows.Forms.Control[] { 
                this.btnCalculate, 
                this.btnExport, 
                this.progressBar1 
            });

            // Set positions based on button panel
            this.cmbStandardLength.Location = new System.Drawing.Point(this.btnCalculate.Left - 250, this.btnCalculate.Top + 3);
            lblStandardLength.Location = new System.Drawing.Point(this.cmbStandardLength.Left - 95, this.cmbStandardLength.Top + 3);

            // Add controls to form in correct order
            this.Controls.AddRange(new System.Windows.Forms.Control[] {
                this.tabControl1,
                buttonPanel,
                this.totalsPanel,
                this.cmbStandardLength,
                lblStandardLength
            });

            this.ResumeLayout(false);
        }

        private void CmbStandardLength_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            currentStandardLength = cmbStandardLength.SelectedIndex == 0 ? 6.0 : 12.0;
            
            if (gridProfiles.DataSource != null)
            {
                // Recalculate profiles with new standard length
                var profiles = gridProfiles.DataSource as System.Collections.Generic.IEnumerable<dynamic>;
                if (profiles != null)
                {
                    var recalculatedProfiles = profiles.Select(p => RecalculateProfile(p)).ToList();
                    gridProfiles.DataSource = recalculatedProfiles;
                    UpdateTotals(recalculatedProfiles, gridPlates.DataSource, gridRods.DataSource, gridBolts.DataSource as System.Collections.Generic.List<BoltItem>);
                }
            }
        }
    }
}
