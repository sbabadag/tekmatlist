using System;
using System.Collections.Generic;
using System.Windows.Forms;
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
        private Dictionary<string, double> _weightCache; // Add weight cache
        private HashSet<double> _allBoltLengths = new HashSet<double>(); // Add this field

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
            ConnectToTekla();
        }

        private void ConnectToTekla()
        {
            try
            {
                _model = new Model();
                if (_model != null)
                {
                    if (_model.GetConnectionStatus())
                    {
                        // Connection successful, update form title
                        var info = _model.GetInfo();
                        this.Text = $"Material List - {info.ModelName}";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error connecting to Tekla: {ex.Message}");
            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                btnGenerate.Enabled = false;
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Maximum = 100;

                if (_model == null)
                {
                    UpdateProgress(5, "Connecting to Tekla...");
                    ConnectToTekla();
                }

                if (_model != null && _model.GetConnectionStatus())
                {
                    UpdateProgress(10, "Getting material list...");
                    var materialList = GetMaterialList();
                    
                    if (materialList.Any())
                    {
                        UpdateProgress(70, "Exporting to Excel...");
                        ExportToExcel(materialList);
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
                UpdateProgress(100, "Complete");
                progressBar1.Visible = false;
                btnGenerate.Enabled = true;
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

                var selector = _model.GetModelObjectSelector();
                var allObjects = selector.GetAllObjects();
                
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
                            if (string.IsNullOrEmpty(material))
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
                    var material = group.Key.Split('-')[0];
                    var profile = group.Key.Split('-')[1];

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

            // Add data
            int totalRows = materials.Count;
            for (int i = 0; i < materials.Count; i++)
            {
                if (i % 10 == 0)
                {
                    int progress = sheetName == "Profiller" ? 80 : 90;
                    UpdateProgress(progress + ((i * 5) / totalRows), $"Processing {sheetName}... {i}/{totalRows}");
                }

                var row = i + 4;
                var material = materials[i];

                if (sheetName == "Levhalar")
                {
                    // Extract plate thickness from profile (e.g., "PL10" -> 10)
                    double thickness = double.Parse(material.Profile.Substring(2));
                    double totalWeight = material.TotalLength * _weightCache[material.Profile] / 1000.0;
                    
                    // Get optimal plate size
                    var (optimalSize, plateCount) = GetOptimalPlateSize(totalWeight, thickness);
                    
                    // Calculate standard plate weight for the optimal size
                    double standardPlateWeight = (thickness / 1000.0) * (optimalSize.Width / 1000.0) * 
                                               (optimalSize.Length / 1000.0) * 7850.0;

                    // Export plate data
                    worksheet.Cell(row, 1).Value = material.Profile;
                    worksheet.Cell(row, 2).Value = Math.Round(totalWeight, 0);
                    worksheet.Cell(row, 3).Value = optimalSize.ToString();
                    worksheet.Cell(row, 4).Value = Math.Round(standardPlateWeight, 2);
                    worksheet.Cell(row, 5).Value = plateCount;
                    worksheet.Cell(row, 6).Value = Math.Round(plateCount * standardPlateWeight, 0);

                    // Replace cell comment creation with proper parameters
                    var cell = worksheet.Cell(row, 3);
                    var otherSizes = STANDARD_PLATE_SIZES
                        .Where(s => s != optimalSize)
                        .Take(3)
                        .Select(s => 
                        {
                            double weight = (thickness / 1000.0) * (s.Width / 1000.0) * (s.Length / 1000.0) * 7850.0;
                            int count = (int)Math.Ceiling(totalWeight / weight);
                            return $"{s}: {count} adet";
                        });

                    // Add comment using newer ClosedXML syntax
                    var note = string.Join("\n", new[] { "Alternatif ölçüler:" }.Concat(otherSizes));
                    worksheet.Cell(row, 3).WorksheetColumn().Width = 25; // Make column wider for comment
                    worksheet.Cell(row, 3).GetComment().AddText(note);

                    // Format numbers
                    worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0";
                    worksheet.Cell(row, 4).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0";
                    worksheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0";
                }
                else
                {
                    // Use existing profile export logic
                    var totalLengthInMeters = Math.Ceiling(material.TotalLength / 1000); // Round up to nearest meter
                    var standardPiecesCount = Math.Ceiling(totalLengthInMeters / STANDARD_LENGTH);
                    var calculatedTotalLength = Math.Ceiling(standardPiecesCount * STANDARD_LENGTH);
                    var unitWeight = _weightCache[material.Profile];
                    var totalWeight = unitWeight * (material.TotalLength / 1000.0);

                    worksheet.Cell(row, 1).Value = material.Profile;
                    worksheet.Cell(row, 2).Value = totalLengthInMeters;
                    worksheet.Cell(row, 3).Value = STANDARD_LENGTH;
                    worksheet.Cell(row, 4).Value = standardPiecesCount;
                    worksheet.Cell(row, 5).Value = calculatedTotalLength;
                    worksheet.Cell(row, 6).Value = Math.Round(unitWeight, 2);
                    worksheet.Cell(row, 7).Value = Math.Round(totalWeight, 0); // Use actual total weight from Tekla

                    // Format numbers with proper decimal separator
                    worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 3).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0.00";
                    worksheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0";
                }

                // Format row
                var dataRange = worksheet.Range(row, 1, row, headers.Length);
                dataRange.Style
                    .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            }

            // Format columns with appropriate widths
            if (sheetName == "Levhalar")
            {
                worksheet.Column(1).Width = 20; // Profile
                worksheet.Column(2).Width = 15; // Total Weight
                worksheet.Column(3).Width = 20; // Standard Plate Size
                worksheet.Column(4).Width = 20; // Standard Plate Weight
                worksheet.Column(5).Width = 15; // Plate Count
                worksheet.Column(6).Width = 15; // Total Weight
            }
            else
            {
                // Use existing column widths for profiles
                worksheet.Column(1).Width = 20; // Profile
                worksheet.Column(2).Width = 15; // Total Length
                worksheet.Column(3).Width = 15; // Standard Length
                worksheet.Column(4).Width = 15; // Standard Piece Count
                worksheet.Column(5).Width = 15; // Calculated Length
                worksheet.Column(6).Width = 15; // Unit Weight
                worksheet.Column(7).Width = 15; // Total Weight
            }

            // Add totals row
            var lastRow = materials.Count + 4;
            worksheet.Cell(lastRow + 1, 1).Value = "TOPLAM";
            
            // Add sum formulas for numeric columns
            if (sheetName == "Levhalar")
            {
                for (int col = 2; col <= 6; col++)
                {
                    if (col != 3) // Skip the plate size column
                    {
                        string colLetter = worksheet.Column(col).ColumnLetter();
                        worksheet.Cell(lastRow + 1, col).FormulaA1 = $"=SUM({colLetter}4:{colLetter}{lastRow})";
                        worksheet.Cell(lastRow + 1, col).Style.NumberFormat.Format = 
                            (col == 4) ? "#,##0.00" : "#,##0";
                    }
                }
            }
            else
            {
                // Use existing totals logic for profiles
                for (int col = 2; col <= 7; col++)
                {
                    string colLetter = worksheet.Column(col).ColumnLetter();
                    worksheet.Cell(lastRow + 1, col).FormulaA1 = $"=SUM({colLetter}4:{colLetter}{lastRow})";
                }
                
                // Format totals row with the same number format
                worksheet.Cell(lastRow + 1, 2).Style.NumberFormat.Format = "#,##0.00";
                worksheet.Cell(lastRow + 1, 3).Style.NumberFormat.Format = "#,##0.00";
                worksheet.Cell(lastRow + 1, 5).Style.NumberFormat.Format = "#,##0.00";
                worksheet.Cell(lastRow + 1, 6).Style.NumberFormat.Format = "#,##0.00";
                worksheet.Cell(lastRow + 1, 7).Style.NumberFormat.Format = "#,##0";
            }

            // Format totals row
            worksheet.Range(lastRow + 1, 1, lastRow + 1, headers.Length).Style
                .Font.SetBold(true)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
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

        private string GetBoltKey(string standard, string size, double length, string assembly)
        {
            // Find the next available longer bolt length that's within 5mm
            var nextLongerLength = _allBoltLengths
                .Where(l => l > length && l <= length + 5)
                .OrderBy(l => l)
                .FirstOrDefault();

            // If we found a longer bolt within 5mm, use that length instead
            var finalLength = nextLongerLength > 0 ? nextLongerLength : length;
            
            return $"{standard}-{size}-{finalLength}-{assembly}";
        }

        private List<BoltItem> GetBoltList()
        {
            var boltGroups = new Dictionary<string, BoltItem>();
            var debugInfo = new List<string>();
            _allBoltLengths.Clear();
            
            try
            {
                // First pass: collect all bolt lengths
                var selector = _model.GetModelObjectSelector();
                var allObjects = selector.GetAllObjects();
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

                // Second pass: group bolts
                allObjects = selector.GetAllObjects();
                while (allObjects.MoveNext())
                {
                    if (allObjects.Current is BoltGroup boltGroup)
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
                            var key = GetBoltKey(bolt, size.ToString(), originalLength, assembly);

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
                                    Assembly = assembly,
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
            
            // Configure page setup
            worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            worksheet.PageSetup.FitToPages(1, 1);
            worksheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;

            // Add title
            worksheet.Cell("A1").Value = "BULON LİSTESİ";
            worksheet.Range("A1:E1").Merge();
            worksheet.Cell("A1").Style
                .Font.SetBold(true)
                .Font.SetFontSize(14)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            // Add headers
            var headers = new[] { 
                "Standard", 
                "Çap", 
                "Boy",
                "Adet",
                "Mark"
            };

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

            // Add data
            int row = 4;
            foreach (var bolt in bolts)
            {
                worksheet.Cell(row, 1).Value = bolt.StandardName;
                worksheet.Cell(row, 2).Value = bolt.Size;
                worksheet.Cell(row, 3).Value = bolt.Length;
                worksheet.Cell(row, 4).Value = bolt.Quantity;
                worksheet.Cell(row, 5).Value = bolt.Assembly;

                // Format row
                worksheet.Range(row, 1, row, 5).Style
                    .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                row++;
            }

            // Add totals
            worksheet.Cell(row, 1).Value = "TOPLAM";
            worksheet.Cell(row, 4).FormulaA1 = $"=SUM(D4:D{row-1})";
            
            // Format totals row
            worksheet.Range(row, 1, row, 5).Style
                .Font.SetBold(true)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            // Set column widths
            worksheet.Column(1).Width = 25; // Standard
            worksheet.Column(2).Width = 15; // Size
            worksheet.Column(3).Width = 15; // Length
            worksheet.Column(4).Width = 15; // Quantity
            worksheet.Column(5).Width = 20; // Assembly Mark
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
    }
}
