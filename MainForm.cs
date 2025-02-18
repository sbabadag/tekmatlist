using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing; // Add for Font, Color
using System.ComponentModel; // Add for ListChangedType
using System.IO;
using System.Linq;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;
using ClosedXML.Excel;
using Tekla.Structures.Catalogs;
using Tekla.Structures.Geometry3d;

namespace TeklaMaterialList
{
    public partial class MainForm : Form
    {
        private Model _model;
        private bool _disposed = false;
        private const string TEMPLATE_FILE = "TEKİRDAG MALZEME LİSTESİ.xlsx";
        private const double STANDARD_LENGTH = 12.0; // Standard profile length in meters
        private const double ROD_STANDARD_LENGTH = 1.0; // Standard rod length in meters
        private Dictionary<string, double> _weightCache; // Add weight cache
        private HashSet<double> _allBoltLengths = new HashSet<double>(); // Add this field
        private CheckBox chkSelectedOnly; // Add this field
        private DataGridView gridNuts; // Add this field
        private TabControl tabControl1;

        // Add steel grades list
        private readonly HashSet<string> STEEL_GRADES = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "S235JR", "S235J0", "S235J2",
            "S275JR", "S275J0", "S275J2",
            "S355JR", "S355J0", "S355J2",
            "S420", "S420N", "S420NL",
            "S460", "S460N", "S460NL"
        };

        private class PlateSize
        {
            public int Width { get; set; }
            public int Length { get; set; }
            public double Area => Width * Length;

            public override string ToString() => $"{Width}x{Length}";
        }

        private readonly List<PlateSize> STANDARD_PLATE_SIZES = new List<PlateSize>
        {
            new PlateSize { Width = 1000, Length = 3000 },
            new PlateSize { Width = 1000, Length = 2000 },
            new PlateSize { Width = 1200, Length = 2400 },
            new PlateSize { Width = 1250, Length = 2500 },
            new PlateSize { Width = 1500, Length = 6000 }, // default
            new PlateSize { Width = 1500, Length = 3000 },
            new PlateSize { Width = 2000, Length = 6000 },
            new PlateSize { Width = 2000, Length = 12000 },
            new PlateSize { Width = 2500, Length = 10000 },
            new PlateSize { Width = 3000, Length = 12000 },
            new PlateSize { Width = 3000, Length = 6000 }
        };

        public MainForm()
        {
            InitializeComponent();
            InitializeCustomComponents(); // Add this line
            SetupGridColumns();
            AddSelectionCheckbox();
            ConnectToTekla();
        }

        private void AddSelectionCheckbox()
        {
            chkSelectedOnly = new CheckBox
            {
                Text = "Selected Objects Only",
                Location = new System.Drawing.Point(btnCalculate.Left - 150, btnCalculate.Top + 3),
                AutoSize = true
            };
            this.Controls.Add(chkSelectedOnly);
        }

        private void InitializeNutsGrid()
        {
            // Remove TabControl/Tab creation code since it's now in InitializeComponent
            this.gridNuts.Paint += Grid_Paint;
        }

        private void SetupGridColumns()
        {
            // Configure grids
            foreach (DataGridView grid in new[] { gridProfiles, gridPlates, gridBolts, gridNuts })
            {
                grid.AllowUserToAddRows = false;
                grid.ReadOnly = true;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grid.ScrollBars = ScrollBars.Both;
                
                // Add a row for totals that will be populated later
                grid.RowTemplate.Height = 25;
                
                // Enable grid features
                grid.RowHeadersVisible = true;
                grid.RowHeadersWidth = 45;
                
                // Add event handler for paint instead of data binding
                grid.Paint += Grid_Paint;
            }
            
            // Configure Profiles grid
            gridProfiles.AutoGenerateColumns = false;
            gridProfiles.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Profile", 
                    DataPropertyName = "Profile", 
                    HeaderText = "Profil",
                    Width = 120
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "ActualWeight",
                    DataPropertyName = "ActualWeight", 
                    HeaderText = "Gerçek Ağırlık (kg)",
                    DefaultCellStyle = { Format = "N0" },
                    Width = 100
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Length",
                    DataPropertyName = "Length", 
                    HeaderText = "Toplam Boy (m)",
                    DefaultCellStyle = { Format = "N2" },
                    Width = 100
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "StandardLength",
                    DataPropertyName = "StandardLength", 
                    HeaderText = "Standart Boy (m)",
                    DefaultCellStyle = { Format = "N2" },
                    Width = 100
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Count",
                    DataPropertyName = "Count", 
                    HeaderText = "Standart Boy Adedi",
                    DefaultCellStyle = { Format = "N0" },
                    Width = 100
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Weight",
                    DataPropertyName = "Weight", 
                    HeaderText = "Birim Ağırlık (kg/m)",
                    DefaultCellStyle = { Format = "N2" },
                    Width = 100
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "TotalWeight",
                    DataPropertyName = "TotalWeight", 
                    HeaderText = "Toplam Ağırlık (kg)",
                    DefaultCellStyle = { Format = "N0" },
                    Width = 100
                }
            });

            // Configure Plates grid
            gridPlates.AutoGenerateColumns = false;
            gridPlates.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Profile",
                    DataPropertyName = "Profile", 
                    HeaderText = "Profil",
                    Width = 120
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "TotalWeight",
                    DataPropertyName = "TotalWeight", 
                    HeaderText = "Toplam Ağırlık (kg)",
                    DefaultCellStyle = { Format = "N0" },
                    Width = 100
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "StandardSize",
                    DataPropertyName = "StandardSize", 
                    HeaderText = "Standart Levha (mm)",
                    Width = 120
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Count",
                    DataPropertyName = "Count", 
                    HeaderText = "Standart Levha Adedi",
                    DefaultCellStyle = { Format = "N0" },
                    Width = 100
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "CalculatedTotalWeight",
                    DataPropertyName = "CalculatedTotalWeight", 
                    HeaderText = "Hesaplanan Ağırlık (kg)",
                    DefaultCellStyle = { Format = "N0" },
                    Width = 100
                }
            });

            // Configure Bolts grid
            gridBolts.AutoGenerateColumns = false;
            gridBolts.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn 
                { 
                    Name = "StandardName",
                    DataPropertyName = "StandardName", 
                    HeaderText = "Standard",
                    Width = 120
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Size",
                    DataPropertyName = "Size", 
                    HeaderText = "Çap",
                    Width = 80
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Length",
                    DataPropertyName = "Length", 
                    HeaderText = "Boy",
                    Width = 80
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Quantity",
                    DataPropertyName = "Quantity", 
                    HeaderText = "Adet",
                    DefaultCellStyle = { Format = "N0" },
                    Width = 80
                }
            });

            // Configure Nuts grid
            gridNuts.AutoGenerateColumns = false;
            gridNuts.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Size", 
                    DataPropertyName = "Size",
                    HeaderText = "Somun Ölçüsü",
                    Width = 120
                },
                new DataGridViewTextBoxColumn 
                { 
                    Name = "Quantity",
                    DataPropertyName = "Quantity", 
                    HeaderText = "Adet",
                    DefaultCellStyle = { Format = "N0" },
                    Width = 100
                }
            });
        }

        private void Grid_Paint(object sender, PaintEventArgs e)
        {
            var grid = (DataGridView)sender;
            if (grid.Rows.Count == 0) return;

            // Calculate totals
            var totals = new Dictionary<int, double>();
            for (int col = 0; col < grid.Columns.Count; col++)
            {
                double total = 0;
                bool isNumeric = false;

                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.Cells[col].Value != null &&
                        (row.Cells[col].Value is int || 
                         row.Cells[col].Value is double || 
                         row.Cells[col].Value is decimal))
                    {
                        total += Convert.ToDouble(row.Cells[col].Value);
                        isNumeric = true;
                    }
                }

                if (isNumeric)
                {
                    totals[col] = total;
                }
            }

            // Draw totals row
            if (totals.Any())
            {
                var rect = grid.GetRowDisplayRectangle(grid.Rows.Count - 1, false);
                rect.Y += rect.Height + 1;
                rect.Height = grid.RowTemplate.Height;

                // Draw background
                using (var brush = new SolidBrush(Color.LightGray))
                {
                    e.Graphics.FillRectangle(brush, rect);
                }

                // Draw "TOPLAM" text
                using (var font = new Font(grid.Font, FontStyle.Bold))
                {
                    var cellRect = grid.GetCellDisplayRectangle(0, grid.Rows.Count - 1, false);
                    cellRect.Y = rect.Y;
                    e.Graphics.DrawString("TOPLAM", font, Brushes.Black, cellRect, 
                        new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });

                    // Draw totals
                    foreach (var total in totals)
                    {
                        var col = total.Key;
                        if (col == 0) continue; // Skip first column (TOPLAM text)

                        cellRect = grid.GetCellDisplayRectangle(col, grid.Rows.Count - 1, false);
                        cellRect.Y = rect.Y;

                        var format = grid.Columns[col].DefaultCellStyle.Format;
                        var value = format.Contains("N0") ? $"{total.Value:N0}" :
                                   format.Contains("N2") ? $"{total.Value:N2}" :
                                   total.Value.ToString();

                        e.Graphics.DrawString(value, font, Brushes.Black, cellRect,
                            new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                    }
                }

                // Draw border
                using (var pen = new Pen(Color.Black))
                {
                    e.Graphics.DrawRectangle(pen, rect);
                }
            }
        }

        private void ConnectToTekla()
        {
            try
            {
                _model = new Model();
                if (_model != null && _model.GetConnectionStatus())
                {
                    var info = _model.GetInfo();
                    this.Text = $"Material List - {info.ModelName}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error connecting to Tekla: {ex.Message}");
            }
        }

        private void btnCalculate_Click(object sender, EventArgs e)
        {
            try
            {
                btnCalculate.Enabled = false;
                btnExport.Enabled = false;
                progressBar1.Value = 0;
                progressBar1.Visible = true;

                if (_model == null)
                {
                    UpdateProgress(5, "Connecting to Tekla...");
                    ConnectToTekla();
                }

                if (_model != null && _model.GetConnectionStatus())
                {
                    // Ask user for selection preference
                    var result = MessageBox.Show(
                        "Do you want to analyze selected objects only?\n\nYes = Selected objects only\nNo = All objects",
                        "Selection Mode",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1);

                    chkSelectedOnly.Checked = (result == DialogResult.Yes);

                    UpdateProgress(10, "Getting material list...");
                    var materialList = GetMaterialList();
                    
                    if (materialList.Any())
                    {
                        // Process all materials in a single pass with efficient grouping
                        var groupedMaterials = materialList
                            .AsParallel() // Use parallel processing for large lists
                            .GroupBy(m => 
                            {
                                if (m.Profile.StartsWith("PL", StringComparison.OrdinalIgnoreCase))
                                    return "Plates";
                                if (m.Profile.StartsWith("D", StringComparison.OrdinalIgnoreCase))
                                    return "Rods";
                                return "Profiles";
                            })
                            .ToDictionary(g => g.Key, g => g.ToList());

                        // Process each category
                        var plates = (groupedMaterials.ContainsKey("Plates") ? groupedMaterials["Plates"] : new List<MaterialItem>())
                            .Select(m => new
                            {
                                Profile = m.Profile,
                                TotalWeight = Math.Round(m.TotalLength * _weightCache[m.Profile] / 1000.0, 0),
                                StandardSize = GetOptimalPlateSize(
                                    m.TotalLength * _weightCache[m.Profile] / 1000.0, 
                                    double.Parse(m.Profile.Substring(2))).Size.ToString(),
                                Count = GetOptimalPlateSize(
                                    m.TotalLength * _weightCache[m.Profile] / 1000.0, 
                                    double.Parse(m.Profile.Substring(2))).Count,
                                CalculatedTotalWeight = Math.Round(GetOptimalPlateSize(
                                    m.TotalLength * _weightCache[m.Profile] / 1000.0, 
                                    double.Parse(m.Profile.Substring(2))).Count * 
                                    (double.Parse(m.Profile.Substring(2)) / 1000.0) * 7850.0, 0)
                            })
                            .ToList();

                        var profiles = (groupedMaterials.ContainsKey("Profiles") ? groupedMaterials["Profiles"] : new List<MaterialItem>())
                            .Select(m => CreateProfileData(m))
                            .ToList();

                        var rods = (groupedMaterials.ContainsKey("Rods") ? groupedMaterials["Rods"] : new List<MaterialItem>())
                            .Select(m => CreateProfileData(m))
                            .ToList();

                        // Update data sources all at once
                        gridPlates.DataSource = null;
                        gridProfiles.DataSource = null;
                        gridRods.DataSource = null;
                        
                        gridPlates.DataSource = plates;
                        gridProfiles.DataSource = profiles;
                        gridRods.DataSource = rods;

                        UpdateProgress(90, "Getting bolt list...");
                        var bolts = GetBoltList();
                        gridBolts.DataSource = bolts;

                        btnExport.Enabled = true;
                        UpdateTotals(profiles, plates, rods, bolts);
                    }
                    else
                    {
                        MessageBox.Show("No materials found in the model.");
                    }
                }
                else
                {
                    MessageBox.Show("Please ensure a Tekla model is open.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                progressBar1.Visible = false;
                btnCalculate.Enabled = true;
            }
        }

        private void UpdateProgress(int value, string status)
        {
            if (progressBar1.InvokeRequired)
            {
                progressBar1.Invoke(new Action(() => UpdateProgress(value, status)));
                return;
            }

            progressBar1.Value = value;
            this.Text = $"Material List - {status}";
            Application.DoEvents();
        }

        private string GetPlateKey(string profile)
        {
            try
            {
                // Handle PL profiles
                if (profile.StartsWith("PL", StringComparison.OrdinalIgnoreCase))
                {
                    var dimensions = profile.Substring(2).Split(new[] { '*', 'X', 'x' });
                    if (dimensions.Length > 0 && double.TryParse(dimensions[0], out double thickness))
                    {
                        return $"PL{thickness}";
                    }
                }

                // Handle L profiles with exactly 3 parameters
                if (profile.StartsWith("L", StringComparison.OrdinalIgnoreCase))
                {
                    var dimensions = profile.Substring(1).Split(new[] { '*', 'X', 'x' });
                    if (dimensions.Length == 3)
                    {
                        if (double.TryParse(dimensions[2], out double thickness))
                        {
                            return $"PL{thickness}"; // Group with plates of same thickness
                        }
                    }
                }

                return profile;
            }
            catch
            {
                return profile;
            }
        }

        private class GroupInfo
        {
            public double TotalWeight = 0;
            public double TotalLength = 0;
        }

        private class ProfileData
        {
            public double TotalWeight { get; set; }
            public double TotalLength { get; set; }
            public HashSet<string> OriginalProfiles { get; set; } = new HashSet<string>();
            public int Quantity { get; set; }
        }

        private List<MaterialItem> GetMaterialList()
        {
            var profileGroups = new Dictionary<string, ProfileData>();
            _weightCache = new Dictionary<string, double>();
            
            try
            {
                if (_model == null || !_model.GetConnectionStatus())
                {
                    _model = new Model();
                    if (!_model.GetConnectionStatus())
                    {
                        MessageBox.Show("Cannot connect to Tekla model. Please open a model first.");
                        return new List<MaterialItem>();
                    }
                }

                ModelObjectEnumerator allObjects;
                if (chkSelectedOnly.Checked)
                {
                    var selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                    allObjects = selector.GetSelectedObjects();
                    
                    // Check if any objects are selected
                    if (!allObjects.GetEnumerator().MoveNext())
                    {
                        MessageBox.Show("No objects selected. Please select objects in the model.");
                        return new List<MaterialItem>();
                    }
                    // Reset enumerator
                    allObjects = selector.GetSelectedObjects();
                }
                else
                {
                    var selector = _model.GetModelObjectSelector();
                    allObjects = selector.GetAllObjects();
                }
                
                int totalObjects = 0;
                int processedParts = 0;
                int failedWeight = 0;
                int failedLength = 0;
                int failedZeroLength = 0;
                int failedProfile = 0;
                int failedMaterial = 0;
                string debugInfo = "";

                while (allObjects.MoveNext())
                {
                    totalObjects++;
                    var obj = allObjects.Current;
                    
                    if (obj is Part part)
                    {
                        processedParts++;
                        string profile = "";
                        string material = "";
                        
                        try
                        {
                            profile = part.Profile.ProfileString;
                            if (string.IsNullOrEmpty(profile))
                            {
                                failedProfile++;
                                continue;
                            }

                            material = part.Material.MaterialString;
                            // Add steel grade check
                            if (string.IsNullOrEmpty(material) || 
                                !STEEL_GRADES.Any(grade => material.StartsWith(grade, StringComparison.OrdinalIgnoreCase)))
                            {
                                failedMaterial++;
                                continue;
                            }

                            // Try different weight property names
                            double weight = 0.0, length = 0.0;
                            bool weightFound = false;
                            
                            foreach (var propName in new[] { "WEIGHT", "WEIGHT_NET", "NET_WEIGHT", "GROSS_WEIGHT", "WEIGHT_GROSS" })
                            {
                                if (part.GetReportProperty(propName, ref weight))
                                {
                                    weightFound = true;
                                    break;
                                }
                            }

                            if (!weightFound)
                            {
                                // Try getting weight from volume
                                double volume = 0.0;
                                if (part.GetReportProperty("VOLUME", ref volume))
                                {
                                    // Convert volume to weight (density * volume)
                                    double density = 7850.0; // kg/m³ for steel
                                    weight = volume * density / 1000000000.0; // Convert mm³ to m³
                                }
                                
                                if (weight <= 0)
                                {
                                    failedWeight++;
                                    debugInfo += $"Failed weight for {profile}: No weight property found\n";
                                    continue;
                                }
                            }

                            if (!part.GetReportProperty("LENGTH", ref length))
                            {
                                // Try getting length from solid
                                var solid = part.GetSolid();
                                if (solid != null)
                                {
                                    // Calculate length using coordinates
                                    var minPoint = solid.MinimumPoint;
                                    var maxPoint = solid.MaximumPoint;
                                    length = Math.Sqrt(
                                        Math.Pow(maxPoint.X - minPoint.X, 2) +
                                        Math.Pow(maxPoint.Y - minPoint.Y, 2) +
                                        Math.Pow(maxPoint.Z - minPoint.Z, 2));
                                }
                                
                                if (length <= 0)
                                {
                                    failedLength++;
                                    continue;
                                }
                            }

                            if (length <= 0)
                            {
                                failedZeroLength++;
                                continue;
                            }

                            string groupedProfile = GetPlateKey(profile);
                            string key = $"{material}-{groupedProfile}";

                            if (!profileGroups.ContainsKey(key))
                            {
                                profileGroups[key] = new ProfileData();
                            }

                            var group = profileGroups[key];
                            group.TotalWeight += weight;
                            group.TotalLength += length;
                            group.Quantity++;
                            group.OriginalProfiles.Add(profile);

                            if (processedParts <= 5)
                            {
                                debugInfo += $"Success - Profile: {profile}, Material: {material}, Weight: {weight:F2}, Length: {length:F2}\n";
                            }
                        }
                        catch (Exception ex)
                        {
                            debugInfo += $"Error with part {profile}-{material}: {ex.Message}\n";
                        }
                    }
                }

                var detailedDebug = 
                    $"Total objects: {totalObjects}\n" +
                    $"Processed parts: {processedParts}\n" +
                    $"Groups created: {profileGroups.Count}\n" +
                    $"Failed profile: {failedProfile}\n" +
                    $"Failed material: {failedMaterial}\n" +
                    $"Failed weight: {failedWeight}\n" +
                    $"Failed length: {failedLength}\n" +
                    $"Failed zero length: {failedZeroLength}\n\n" +
                    $"First few parts processed:\n{debugInfo}";

                MessageBox.Show(detailedDebug);

                if (profileGroups.Count == 0)
                {
                    return new List<MaterialItem>();
                }

                // Convert to MaterialItems and cache weights
                var materials = new List<MaterialItem>();
                foreach (var group in profileGroups)
                {
                    var keyParts = group.Key.Split('-');
                    var material = keyParts[0];
                    var profile = keyParts[1];

                    // Cache the weight per meter
                    _weightCache[profile] = (group.Value.TotalWeight / group.Value.TotalLength) * 1000.0;

                    materials.Add(new MaterialItem
                    {
                        Material = material,
                        Profile = profile,
                        Quantity = group.Value.Quantity,
                        TotalLength = group.Value.TotalLength,
                        OriginalProfile = string.Join(", ", group.Value.OriginalProfiles)
                    });
                }

                return materials;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting materials: {ex.Message}\n\nStack: {ex.StackTrace}");
                return new List<MaterialItem>();
            }
        }

        private void ExportToExcel(List<MaterialItem> materials)
        {
            UpdateProgress(75, "Preparing Excel export...");

            // Split materials into plates and other profiles
            var plates = materials.Where(m => m.Profile.StartsWith("PL"))
                .OrderBy(m => m.Profile) // Sort plates by profile name
                .ToList();
            var otherProfiles = materials.Where(m => !m.Profile.StartsWith("PL")).ToList();

            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveDialog.FilterIndex = 1;
                saveDialog.FileName = $"MALZEME LİSTESİ_{DateTime.Now:yyyyMMdd}";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveDialog.FileName;
                    using (var workbook = new XLWorkbook())
                    {
                        UpdateProgress(80, "Exporting profiles...");
                        ExportMaterialsToWorksheet(workbook, "Profiller", otherProfiles);
                        
                        UpdateProgress(90, "Exporting plates...");
                        ExportMaterialsToWorksheet(workbook, "Levhalar", plates);

                        UpdateProgress(90, "Getting bolt list...");
                        var bolts = GetBoltList();
                        
                        UpdateProgress(95, "Exporting bolts...");
                        ExportBoltsToWorksheet(workbook, "Bulonlar", bolts);

                        UpdateProgress(90, "Exporting nuts...");
                        var nuts = gridNuts.DataSource as List<NutItem>;
                        ExportNutsToWorksheet(workbook, "Somunlar", nuts);

                        UpdateProgress(95, "Saving Excel file...");
                        workbook.SaveAs(filePath);
                        MessageBox.Show("Malzeme listesi başarıyla oluşturuldu!");

                        // Open the file with Excel
                        try
                        {
                            System.Diagnostics.Process.Start(filePath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Excel dosyası kaydedildi ancak açılamadı: {ex.Message}");
                        }
                    }
                }
            }
        }

        private (PlateSize Size, int Count) GetOptimalPlateSize(double totalWeight, double thickness)
        {
            var results = new List<(PlateSize Size, int Count, double WasteWeight)>();
            
            foreach (var plateSize in STANDARD_PLATE_SIZES)
            {
                // Calculate weight of one standard plate
                double plateWeight = (thickness / 1000.0) * (plateSize.Width / 1000.0) * 
                                   (plateSize.Length / 1000.0) * 7850.0; // density in kg/m³
                
                int plateCount = (int)Math.Ceiling(totalWeight / plateWeight);
                double totalPlateWeight = plateCount * plateWeight;
                double wasteWeight = totalPlateWeight - totalWeight;

                results.Add((plateSize, plateCount, wasteWeight));
            }

            // Find the option with minimum waste
            var optimalResult = results
                .OrderBy(r => r.WasteWeight)  // First prioritize minimum waste
                .ThenBy(r => r.Count)         // Then minimize plate count
                .First();

            return (optimalResult.Size, optimalResult.Count);
        }

        private void ExportMaterialsToWorksheet(XLWorkbook workbook, string sheetName, List<MaterialItem> materials)
        {
            var worksheet = workbook.Worksheets.Add(sheetName);
            
            // Configure page setup and title (same as before)
            worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            worksheet.PageSetup.FitToPages(1, 1);
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            worksheet.PageSetup.ShowGridlines = true;
            
            // Add title
            worksheet.Cell("A1").Value = "MALZEME LİSTESİ - " + sheetName.ToUpper();
            worksheet.Range("A1:G1").Merge();
            worksheet.Cell("A1").Style
                .Font.SetBold(true)
                .Font.SetFontSize(14)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            // Modify headers based on sheet type
            string[] headers;
            if (sheetName == "Levhalar")
            {
                headers = new[] { 
                    "Profil", 
                    "Toplam Ağırlık (kg)",
                    "Standart Levha (mm)",
                    "Standart Levha Ağırlığı (kg)",
                    "Standart Levha Adedi",
                    "Toplam Ağırlık (kg)"
                };
            }
            else
            {
                // Use existing headers for profiles
                headers = new[] { 
                    "Profil", 
                    "Toplam Boy (m)", 
                    "Standart Boy (m)", 
                    "Standart Boy Adedi",
                    "Hesaplanan Boy (m)",
                    "Birim Ağırlık (kg/m)",
                    "Toplam Ağırlık (kg)"
                };
            }

            // Add headers
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(3, i + 1);
                cell.Value = headers[i];
                cell.Style
                    .Font.SetBold(true)
                    .Fill.SetBackgroundColor(XLColor.LightGray)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            }

            if (sheetName == "Levhalar")
            {
                // Add plate sizes to a hidden sheet for reference
                var sizesSheet = workbook.Worksheets.Add("PlateSizes");
                int sizeRow = 1;
                foreach (var size in STANDARD_PLATE_SIZES)
                {
                    sizesSheet.Cell(sizeRow++, 1).Value = size.ToString();
                }
                sizesSheet.Hide();

                // Add headers
                // ...existing headers code...

                double totalOriginalWeight = 0;
                double totalCalculatedWeight = 0;

                for (int i = 0; i < materials.Count; i++)
                {
                    var row = i + 4;
                    var material = materials[i];
                    double thickness = double.Parse(material.Profile.Substring(2));
                    double totalWeight = material.TotalLength * _weightCache[material.Profile] / 1000.0;

                    // Get optimal size calculations
                    var (optimalSize, plateCount) = GetOptimalPlateSize(totalWeight, thickness);
                    double standardPlateWeight = (thickness / 1000.0) * (optimalSize.Width / 1000.0) * 
                                            (optimalSize.Length / 1000.0) * 7850.0;
                    double calculatedTotalWeight = plateCount * standardPlateWeight;

                    // Update running totals
                    totalOriginalWeight += totalWeight;
                    totalCalculatedWeight += calculatedTotalWeight;

                    // Export plate data with direct values instead of formulas
                    worksheet.Cell(row, 1).Value = material.Profile;
                    worksheet.Cell(row, 2).Value = Math.Round(totalWeight, 0);
                    worksheet.Cell(row, 3).Value = optimalSize.ToString();
                    worksheet.Cell(row, 4).Value = Math.Round(standardPlateWeight, 2);
                    worksheet.Cell(row, 5).Value = plateCount;
                    worksheet.Cell(row, 6).Value = Math.Round(calculatedTotalWeight, 0);

                    // Format numbers
                    worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0";
                    worksheet.Cell(row, 4).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0";
                    worksheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0";
                }

                // Calculate and add totals without formulas
                var lastRow = materials.Count + 4;
                worksheet.Cell(lastRow + 1, 1).Value = "TOPLAM";
                worksheet.Cell(lastRow + 1, 2).Value = Math.Round(totalOriginalWeight, 0);
                worksheet.Cell(lastRow + 1, 6).Value = Math.Round(totalCalculatedWeight, 0);

                if (totalOriginalWeight > 0)
                {
                    worksheet.Cell(lastRow + 2, 1).Value = "TOPLAM FİRE (%)";
                    worksheet.Cell(lastRow + 2, 2).Value = Math.Round(
                        ((totalCalculatedWeight - totalOriginalWeight) / totalOriginalWeight) * 100, 2);
                }

                // Protect specific cells
                var protectedRanges = new[] { "A:A", "B:B", "D:D", "F:F" };
                foreach (var range in protectedRanges)
                {
                    worksheet.Range(range).Style.Protection.SetLocked(true);
                }
                worksheet.Range($"C4:C{lastRow}").Style.Protection.SetLocked(false);

                worksheet.Protect()
                    .AllowElement(XLSheetProtectionElements.SelectLockedCells)
                    .AllowElement(XLSheetProtectionElements.SelectUnlockedCells)
                    .AllowElement(XLSheetProtectionElements.Sort)
                    .AllowElement(XLSheetProtectionElements.AutoFilter);
            }
            else
            {
                // Handle profiles sheet
                // ...existing headers code...

                double totalLength = 0;
                double totalWeight = 0;
                int totalCount = 0;

                for (int i = 0; i < materials.Count; i++)
                {
                    var row = i + 4;
                    var material = materials[i];

                    var totalLengthInMeters = Math.Ceiling(material.TotalLength / 1000);
                    var standardPiecesCount = (int)Math.Ceiling(totalLengthInMeters / STANDARD_LENGTH);
                    var calculatedTotalLength = standardPiecesCount * STANDARD_LENGTH;
                    var unitWeight = _weightCache[material.Profile];
                    var profileTotalWeight = unitWeight * calculatedTotalLength;

                    // Update running totals
                    totalLength += totalLengthInMeters;
                    totalCount += standardPiecesCount;
                    totalWeight += profileTotalWeight;

                    // Export values
                    worksheet.Cell(row, 1).Value = material.Profile;
                    worksheet.Cell(row, 2).Value = totalLengthInMeters;
                    worksheet.Cell(row, 3).Value = STANDARD_LENGTH;
                    worksheet.Cell(row, 4).Value = standardPiecesCount;
                    worksheet.Cell(row, 5).Value = calculatedTotalLength;
                    worksheet.Cell(row, 6).Value = Math.Round(unitWeight, 2);
                    worksheet.Cell(row, 7).Value = Math.Round(profileTotalWeight, 0);

                    // Format numbers
                    worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 3).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0";
                }

                // Add totals directly
                var lastRow = materials.Count + 4;
                worksheet.Cell(lastRow + 1, 1).Value = "TOPLAM";
                worksheet.Cell(lastRow + 1, 2).Value = Math.Round(totalLength, 2);
                worksheet.Cell(lastRow + 1, 4).Value = totalCount;
                worksheet.Cell(lastRow + 1, 5).Value = Math.Round(totalLength, 2);
                worksheet.Cell(lastRow + 1, 7).Value = Math.Round(totalWeight, 0);
            }

            // Add borders to all used cells
            var usedRange = worksheet.Range(worksheet.FirstCellUsed(), worksheet.LastCellUsed());
            usedRange.Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
            usedRange.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            
            // Set zebra striping for better readability
            for (int row = 4; row <= worksheet.LastRowUsed().RowNumber(); row++)
            {
                if (row % 2 == 0)
                {
                    worksheet.Row(row).Style.Fill.SetBackgroundColor(XLColor.FromArgb(242, 242, 242));
                }
            }
        }

        private double GetProfileWeight(string profile)
        {
            try
            {
                if (_model != null && _model.GetConnectionStatus())
                {
                    // Try to get an example part with this profile from the model
                    var selector = _model.GetModelObjectSelector();
                    var allObjects = selector.GetAllObjects();
                    
                    while (allObjects.MoveNext())
                    {
                        if (allObjects.Current is Part part && 
                            part.Profile.ProfileString.Equals(profile, StringComparison.OrdinalIgnoreCase))
                        {
                            double weight = 0.0, length = 0.0;

                            // First try to get weight per meter directly
                            if (part.GetReportProperty("WEIGHT_PER_UNIT_LENGTH", ref weight) && weight > 0)
                            {
                                return weight;
                            }

                            // If that fails, calculate from total weight and length
                            if (part.GetReportProperty("NET_WEIGHT", ref weight) &&
                                part.GetReportProperty("LENGTH", ref length) &&
                                length > 0)
                            {
                                return (weight / length) * 1000.0; // Convert to kg/m
                            }
                        }
                    }
                }

                MessageBox.Show($"Cannot determine weight for profile: {profile}", "Warning");
                return 0.0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting profile weight: {ex.Message}");
                return 0.0;
            }
        }

        private class BoltItem
        {
            public string StandardName { get; set; }
            public string Size { get; set; }
            public string Length { get; set; }
            public int Quantity { get; set; }
            public string Assembly { get; set; }
        }

        private class NutItem
        {
            public int Size { get; set; }
            public int Quantity { get; set; }
        }

        private string GetBoltKey(string standard, string size, double length, string assembly)
        {
            // Find the next available longer bolt length that's within 5mm
            var nextLongerLength = _allBoltLengths
                .Where(l => l > length && l <= length + 5)
                .OrderBy(l => l)
                .FirstOrDefault();

            // If we found a longer bolt within 5mm, use that length instead
            var finalLength = nextLongerLength > 0 ? nextLongerLength : length;
            
            // Remove assembly from key to group all same size/length bolts together
            return $"{standard}-{size}-{finalLength}";
        }

        private List<BoltItem> GetBoltList()
        {
            var boltGroups = new Dictionary<string, BoltItem>();
            var nutGroups = new Dictionary<int, int>(); // Size -> Quantity
            var debugInfo = new List<string>();
            _allBoltLengths.Clear();
            
            try
            {
                ModelObjectEnumerator allObjects;
                if (chkSelectedOnly.Checked)
                {
                    var uiSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                    allObjects = uiSelector.GetSelectedObjects();
                    
                    // Check if any objects are selected
                    if (!allObjects.GetEnumerator().MoveNext())
                    {
                        return new List<BoltItem>();
                    }
                    // Reset enumerator
                    allObjects = uiSelector.GetSelectedObjects();
                }
                else
                {
                    var modelSelector = _model.GetModelObjectSelector();
                    allObjects = modelSelector.GetAllObjects();
                }

                // First pass: collect all bolt lengths
                while (allObjects.MoveNext())
                {
                    if (allObjects.Current is BoltGroup boltGroup)
                    {
                        double length = 0.0;
                        if (boltGroup.GetReportProperty("LENGTH", ref length) || 
                            boltGroup.GetReportProperty("BOLT_LENGTH", ref length))
                        {
                            _allBoltLengths.Add(Math.Round(length)); // Round to nearest mm
                        }
                    }
                }

                // Sort lengths for easier reference - fixed OrderBy method name
                var sortedLengths = _allBoltLengths.OrderBy(l => l).ToList();
                _allBoltLengths = new HashSet<double>(sortedLengths);

                // Second pass: group bolts - get fresh enumerator
                if (chkSelectedOnly.Checked)
                {
                    var uiSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                    allObjects = uiSelector.GetSelectedObjects();
                }
                else
                {
                    var modelSelector = _model.GetModelObjectSelector();
                    allObjects = modelSelector.GetAllObjects();
                }

                while (allObjects.MoveNext())
                {
                    if (allObjects.Current is Part part && 
                        part.Profile.ProfileString.StartsWith("PD", StringComparison.OrdinalIgnoreCase))
                    {
                        var nutSize = ParseNutSize(part.Profile.ProfileString);
                        if (nutSize > 0)
                        {
                            if (!nutGroups.ContainsKey(nutSize))
                                nutGroups[nutSize] = 0;
                            nutGroups[nutSize]++;
                        }
                    }
                    else if (allObjects.Current is BoltGroup boltGroup)
                    {
                        try
                        {
                            var bolt = boltGroup.BoltStandard;
                            var size = boltGroup.BoltSize;
                            double length = 0.0;
                            int count = 0;

                            if (!boltGroup.GetReportProperty("LENGTH", ref length))
                            {
                                boltGroup.GetReportProperty("BOLT_LENGTH", ref length);
                            }

                            // Get bolt count
                            foreach (var countProp in new[] { "BOLT_COUNT", "NUMBER_OF_BOLTS", "BOLT_NUMBER" })
                            {
                                if (boltGroup.GetReportProperty(countProp, ref count) && count > 0)
                                    break;
                            }

                            if (count == 0)
                            {
                                var positions = boltGroup.BoltPositions;
                                count = positions?.Count ?? 0;
                            }
                            
                            var assembly = boltGroup.PartToBeBolted?.GetAssembly()?.AssemblyNumber.Prefix ?? "N/A";
                            var originalLength = Math.Round(length); // Round to nearest mm
                            var key = GetBoltKey(bolt, size.ToString(), originalLength, "");  // Empty assembly

                            if (!boltGroups.ContainsKey(key))
                            {
                                var nextLength = _allBoltLengths
                                    .Where(l => l > originalLength && l <= originalLength + 5)
                                    .OrderBy(l => l)
                                    .FirstOrDefault();

                                var targetLength = nextLength > 0 ? nextLength : originalLength;
                                
                                if (nextLength > 0)
                                {
                                    debugInfo.Add($"Grouped {bolt} M{size}x{originalLength} -> M{size}x{targetLength}");
                                }
                                
                                boltGroups[key] = new BoltItem
                                {
                                    StandardName = bolt,
                                    Size = size.ToString(),
                                    Length = targetLength.ToString(),
                                    Assembly = "ALL", // Mark as combined
                                    Quantity = 0
                                };
                            }

                            boltGroups[key].Quantity += count;
                        }
                        catch (Exception ex)
                        {
                            debugInfo.Add($"Error processing bolt: {ex.Message}");
                            continue;
                        }
                    }
                }

                // Show grouping decisions
                if (debugInfo.Any())
                {
                    MessageBox.Show(
                        "Bolt Grouping Info:\n" + string.Join("\n", debugInfo.Take(10)) + 
                        (debugInfo.Count > 10 ? "\n..." : ""),
                        "Bolt Grouping Debug"
                    );
                }

                // Convert nuts to list and assign to grid
                gridNuts.DataSource = nutGroups
                    .Select(n => new NutItem { Size = n.Key, Quantity = n.Value })
                    .OrderBy(n => n.Size)
                    .ToList();

                return boltGroups.Values
                    .OrderBy(b => b.StandardName)
                    .ThenBy(b => double.Parse(b.Size))
                    .ThenBy(b => double.Parse(b.Length))
                    .ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting bolts: {ex.Message}");
                return new List<BoltItem>();
            }
        }

        private void ExportBoltsToWorksheet(XLWorkbook workbook, string sheetName, List<BoltItem> bolts)
        {
            var worksheet = workbook.Worksheets.Add(sheetName);
            
            // Configure page setup with grid
            worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            worksheet.PageSetup.FitToPages(1, 1);
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            worksheet.PageSetup.ShowGridlines = true;
            
            // Set default column settings
            worksheet.RowHeight = 20;
            worksheet.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

            // Add title and configure sheet
            worksheet.Cell("A1").Value = "BULON LİSTESİ";
            worksheet.Range("A1:D1").Merge();
            worksheet.Cell("A1").Style
                .Font.SetBold(true)
                .Font.SetFontSize(14)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            // Enable filtering and freeze panes
            var dataRange = worksheet.Range("A3:D3");
            dataRange.SetAutoFilter();
            worksheet.SheetView.Freeze(3, 0);

            // Add headers - simplified version
            var headers = new[] { "Standard", "Çap", "Boy", "Adet" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(3, i + 1);
                cell.Value = headers[i];
                cell.Style
                    .Font.SetBold(true)
                    .Fill.SetBackgroundColor(XLColor.LightGray)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            }

            // Add data rows
            int currentRow = 4;
            int totalBolts = 0;
            foreach (var bolt in bolts)
            {
                worksheet.Cell(currentRow, 1).Value = bolt.StandardName;
                worksheet.Cell(currentRow, 2).Value = bolt.Size;
                worksheet.Cell(currentRow, 3).Value = bolt.Length;
                worksheet.Cell(currentRow, 4).Value = bolt.Quantity;

                totalBolts += bolt.Quantity;
                currentRow++;
            }

            // Add single total row
            worksheet.Cell(currentRow, 1).Value = "TOPLAM";
            worksheet.Cell(currentRow, 4).Value = totalBolts;
            
            worksheet.Range($"A{currentRow}:D{currentRow}").Style
                .Font.SetBold(true)
                .Fill.SetBackgroundColor(XLColor.LightGray)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin);

            // Adjust column widths
            worksheet.Column(1).Width = 25;  // Standard
            worksheet.Column(2).Width = 12;  // Size
            worksheet.Column(3).Width = 12;  // Length
            worksheet.Column(4).Width = 15;  // Quantity

            // Add borders to all used cells
            var usedRange = worksheet.Range(worksheet.FirstCellUsed(), worksheet.LastCellUsed());
            usedRange.Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
            usedRange.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            
            // Set zebra striping for better readability
            var lastUsedRow = worksheet.LastRowUsed().RowNumber();
            for (int rowNum = 4; rowNum <= lastUsedRow; rowNum++)
            {
                if (rowNum % 2 == 0)
                {
                    worksheet.Row(rowNum).Style.Fill.SetBackgroundColor(XLColor.FromArgb(242, 242, 242));
                }
            }
        }

        private void ExportNutsToWorksheet(XLWorkbook workbook, string sheetName, List<NutItem> nuts)
        {
            var worksheet = workbook.Worksheets.Add(sheetName);
            
            worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            worksheet.PageSetup.FitToPages(1, 1);
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            worksheet.PageSetup.ShowGridlines = true;
            
            worksheet.Cell("A1").Value = "SOMUN LİSTESİ";
            worksheet.Range("A1:B1").Merge();
            worksheet.Cell("A1").Style
                .Font.SetBold(true)
                .Font.SetFontSize(14)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            var headers = new[] { "Ölçü", "Adet" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(3, i + 1);
                cell.Value = headers[i];
                cell.Style
                    .Font.SetBold(true)
                    .Fill.SetBackgroundColor(XLColor.LightGray)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            }

            int currentRow = 4;
            int totalNuts = 0;
            foreach (var nut in nuts)
            {
                worksheet.Cell(currentRow, 1).Value = $"M{nut.Size}";
                worksheet.Cell(currentRow, 2).Value = nut.Quantity;
                totalNuts += nut.Quantity;
                currentRow++;
            }

            worksheet.Cell(currentRow, 1).Value = "TOPLAM";
            worksheet.Cell(currentRow, 2).Value = totalNuts;
            
            worksheet.Range($"A{currentRow}:B{currentRow}").Style
                .Font.SetBold(true)
                .Fill.SetBackgroundColor(XLColor.LightGray)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin);

            var usedRange = worksheet.Range(worksheet.FirstCellUsed(), worksheet.LastCellUsed());
            usedRange.Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
            usedRange.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (gridProfiles.DataSource == null || 
                gridPlates.DataSource == null || 
                gridBolts.DataSource == null)
            {
                MessageBox.Show("Please calculate materials first.");
                return;
            }

            try
            {
                btnExport.Enabled = false;
                progressBar1.Value = 0;
                progressBar1.Visible = true;

                // Create a new list of MaterialItem objects
                var materials = new List<MaterialItem>();
                
                // Convert profiles to MaterialItems
                var profiles = gridProfiles.DataSource as IEnumerable<dynamic>;
                if (profiles != null)
                {
                    foreach (var p in profiles)
                    {
                        materials.Add(new MaterialItem
                        {
                            Profile = p.Profile,
                            TotalLength = p.Length * 1000, // Convert back to mm
                            Material = "ALL" // We don't need material for export
                        });
                    }
                }

                // Convert plates to MaterialItems
                var plates = gridPlates.DataSource as IEnumerable<dynamic>;
                if (plates != null)
                {
                    foreach (var p in plates)
                    {
                        materials.Add(new MaterialItem
                        {
                            Profile = p.Profile,
                            TotalLength = (p.TotalWeight * 1000.0) / _weightCache[p.Profile], // Calculate length from weight
                            Material = "ALL"
                        });
                    }
                }

                ExportToExcel(materials);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export error: {ex.Message}\n{ex.StackTrace}");
            }
            finally
            {
                progressBar1.Visible = false;
                btnExport.Enabled = true;
            }
        }

        private void UpdateTotals(dynamic profiles, dynamic plates, dynamic rods, List<BoltItem> bolts)
        {
            try
            {
                // Calculate profile totals
                double totalProfileWeight = 0;
                int totalProfileCount = 0;
                foreach (var profile in profiles)
                {
                    totalProfileWeight += (double)profile.Length * (double)profile.Weight;
                    totalProfileCount += (int)profile.Count;
                }

                // Calculate plate totals
                double totalPlateWeight = 0;
                int totalPlateCount = 0;
                foreach (var plate in plates)
                {
                    totalPlateWeight += (double)plate.TotalWeight;
                    totalPlateCount += (int)plate.Count;
                }

                // Calculate rod totals
                double totalRodWeight = 0;
                int totalRodCount = 0;
                foreach (var rod in rods)
                {
                    totalRodWeight += (double)rod.TotalWeight;
                    totalRodCount += (int)rod.Count;
                }

                // Calculate bolt totals
                int totalBoltCount = bolts.Sum(b => b.Quantity);

                // Add nut totals
                var nuts = gridNuts.DataSource as List<NutItem>;
                int totalNutCount = nuts?.Sum(n => n.Quantity) ?? 0;

                // Update labels with proper formatting
                lblTotalProfiles.Text = string.Format("Profil: {0} adet, {1:N0} kg", 
                    totalProfileCount, totalProfileWeight);
                lblTotalPlates.Text = string.Format("Levha: {0} adet, {1:N0} kg", 
                    totalPlateCount, totalPlateWeight);
                lblTotalRods.Text = string.Format("Çubuk: {0} adet, {1:N0} kg", 
                    totalRodCount, totalRodWeight);
                lblTotalBolts.Text = string.Format("Bulon: {0} adet, Somun: {1} adet", 
                    totalBoltCount, totalNutCount);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating totals: {ex.Message}");
            }
        }

        private dynamic CreateProfileData(MaterialItem m)
        {
            var totalLengthInMeters = Math.Ceiling(m.TotalLength / 1000);
            // Use different standard length for rods (D profiles)
            var standardLength = m.Profile.StartsWith("D", StringComparison.OrdinalIgnoreCase) 
                ? ROD_STANDARD_LENGTH 
                : STANDARD_LENGTH;
            
            var standardPiecesCount = (int)Math.Ceiling(totalLengthInMeters / standardLength);
            var unitWeight = _weightCache[m.Profile];
            
            return new
            {
                Profile = m.Profile,
                ActualWeight = Math.Round(unitWeight * (m.TotalLength / 1000.0), 0),
                Length = totalLengthInMeters,
                StandardLength = standardLength,
                Count = standardPiecesCount,
                Weight = Math.Round(unitWeight, 2),
                TotalWeight = Math.Round(
                    standardPiecesCount * standardLength * unitWeight,
                    0)
            };
        }

        private int ParseNutSize(string pdProfile)
        {
            try
            {
                // Format is PD(A)*(B) where nut size = A - 2*B
                var parts = pdProfile.Substring(2).Split('*');
                if (parts.Length == 2)
                {
                    var a = int.Parse(parts[0].Trim('(', ')'));
                    var b = int.Parse(parts[1]);
                    return a - (2 * b);
                }
            }
            catch { }
            return 0;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (_model != null)
            {
                _model.CommitChanges();
                _model = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_model != null)
                    {
                        _model.CommitChanges();
                        _model = null;
                    }
                    if (components != null)
                    {
                        components.Dispose();
                    }
                }
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        private void InitializeCustomComponents()
        {
            // Create tabs
            var profilesTab = new TabPage("Profiller");
            var platesTab = new TabPage("Levhalar");
            var rodsTab = new TabPage("Çubuklar");
            var boltsTab = new TabPage("Bulonlar");
            var nutsTab = new TabPage("Somunlar");

            // Setup TabControl
            this.tabControl1 = new TabControl
            {
                Dock = DockStyle.Fill,
                Location = new System.Drawing.Point(0, 0), // Use fully qualified name
                Name = "tabControl1",
                SelectedIndex = 0,
                Size = new System.Drawing.Size(800, 450), // Use fully qualified name
                TabIndex = 0
            };

            // Add tabs
            this.tabControl1.TabPages.AddRange(new[] {
                profilesTab,
                platesTab,
                rodsTab,
                boltsTab,
                nutsTab
            });

            // Initialize grids
            this.gridProfiles = new DataGridView();
            this.gridPlates = new DataGridView();
            this.gridRods = new DataGridView();
            this.gridBolts = new DataGridView();
            this.gridNuts = new DataGridView();

            // Configure grid properties and add to tabs
            foreach (var pair in new[] { 
                (gridProfiles, profilesTab),
                (gridPlates, platesTab),
                (gridRods, rodsTab),
                (gridBolts, boltsTab),
                (gridNuts, nutsTab)
            })
            {
                var grid = pair.Item1;
                var tab = pair.Item2;
                
                grid.Dock = DockStyle.Fill;
                grid.AllowUserToAddRows = false;
                grid.ReadOnly = true;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grid.ScrollBars = ScrollBars.Both;
                grid.RowTemplate.Height = 25;
                grid.RowHeadersVisible = true;
                grid.RowHeadersWidth = 45;
                
                tab.Controls.Add(grid);
            }

            // Add TabControl to form
            this.Controls.Add(this.tabControl1);
        }
    }
}
