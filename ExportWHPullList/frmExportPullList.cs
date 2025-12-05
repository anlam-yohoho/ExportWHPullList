using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.MsForms;
using DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office2019.Excel.RichData2;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.SqlServer.Server;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using static ClosedXML.Excel.XLPredefinedFormat;

namespace ExportWHPullList
{
    /// <summary>
    /// Small helper to read a ClosedXML IXLCell safely from a WinForms app.
    /// Returns a CellReadResult describing empty/formula/type/typed value/display text and any error encountered.
    /// </summary>

    public partial class frmExportPullList : Form
    {
        //Globals variables:
        private string _modelVsMaterialsLayoutDB;//VB6: Global ModelvsMaterialsLayoutDB As String
        private string _pulledList_SectorGroup; //VB6: Global PulledListSectorGroup As String
        private string _pulledList_Type; //VB6: Global SectorMaterialsPulledListType As String
        //private int autoLockPlannerRights; //VB6: Global AutoLockPlannerRights As Integer
        private bool _isPulledListInLocalFolder; //VB6: Global GIsLocalPulledListExist As Boolean
        private bool _blnMultiPOExport; //VB6: Global MultiPOExport As Boolean
        //private int intGetPrevStartRow; //VB6: Global GetPrevStartRow As Integer
        private int _repeatedProcess; //VB6: Global RepeatedProcess As Integer
        private bool _blnSMTMaterialsMissingSetup; //VB6: Global SMTMaterialsMissingSetup As Boolean
        //private string pulledListID; //VB6: Global PulledListID As String
        private bool blnGoodKANBANinf; //VB6: Global GoodKANBANinf As Boolean
        private string _strUserName; //VB6: Global strUSERNAME As String
        private double _sumPOValue; //VB6: Global sumPOVal As Double
        private bool _enableToOverrideWritePulledListToLocal; //VB6: Global GEnableToOverrideWritePulledListToLocal As Boolean
        frmProgressBar progressBarForm = new frmProgressBar();
        MSSQL _sql = new MSSQL();

        public frmExportPullList()
        {
            InitializeComponent();
            this.MinimumSize = new System.Drawing.Size(900, 600);

            GetUserNameAndSetup();

            GetDatabaseServer();
            GetDirChangeOverLog();
            GetLabelRow();

            // Populate combo boxes:
            Populate_cbbPulledListLine();

            // Setup DataGridViews:
            Setup_dgvPulledListPO();
            Setup_dgvKANBANPulledList();
            Setup_dgvMaterialsConversionMatrix();
            Setup_dgvMaterialsOnTrayNonProgramMatrix();
            Setup_dgvUniPOnQtyMaterials();
            Setup_dgvMultiUniPhysicalModelPulled();//flxMultiUniPhysicalModelPulled
            Setup_dgvPullListvsPO();
            Setup_dgvPullListvsPO2();
            Setup_dgvPLOverallModelPulled();
            Setup_dgvPLPhysicalModelPulled();
            Setup_dgvUniPhysicalModelPulled();
            //Setup_dgvPartvsQty();
            Setup_dgvQtyvsCountDuplicated();
            Setup_dgvPhysicalSAPModelAfterCOPulled();
            Setup_dgvPhysicalModelRunningPulled();
            Setup_dgvPotentialIssues();
            Setup_dgvMaterialsRarDivideKANBANBox();
            Setup_dgvCommonPartPO();
            Setup_dgvPartFirstPO();
            Setup_dgvPartPCBAOfPO();
            Setup_dgvPartRestOfPO();

            // Initial data load/setup:
            GetAllSector(cbbPulledListLine.Text.Trim());
            GetDefectType();

            UpdateDatabaseAll();
          
            //autoLockPlannerRights = 0;

            // Attach event handler for selection change:
            cbbPulledListLine.SelectionChangeCommitted += cbbLine_SelectionChangeCommitted;
            chkbLabelsPrint.Checked = Properties.Settings.Default.blnLabelsPrint;
        }

        private void UpdateDatabaseAll()
        {
            frmProgressBar updateProgress = new frmProgressBar();
            updateProgress.Show();
            updateProgress.UpdateProgress(1, "Updating Database...");
            updateProgress.Show();

            try
            {
                updateProgress.UpdateProgress(5, "Update_MaterialsMSLnFLControlTxt...");
                Update_MaterialsMSLnFLControlTxt();
                updateProgress.UpdateProgress(10, "Update_MPHAllSectorTxt...");
                Update_MPHAllSectorTxt();
                updateProgress.UpdateProgress(15, "Update_MPHKANBANFixLocTxt...");
                Update_MPHKANBANFixLocTxt();
                updateProgress.UpdateProgress(20, "Update_MPHMROListTxt...");
                Update_MPHMROListTxt();
                updateProgress.UpdateProgress(25, "Update_PartRemovedFromBOMTxt...");
                Update_PartRemovedFromBOMTxt();
                updateProgress.UpdateProgress(30, "Update_PCBModelPanelvsQtyperPanelTxt...");
                Update_PCBModelPanelvsQtyperPanelTxt();
                updateProgress.UpdateProgress(40, "Update_PhantomSubMaterialsvsModelTxt...");
                Update_PhantomSubMaterialsvsModelTxt();
                updateProgress.UpdateProgress(50, "Update_PreAssyModelLayoutTxt...");
                Update_PreAssyModelLayoutTxt();
                updateProgress.UpdateProgress(60, "Update_ProductionMaterialsMissingRateControlTxt...");
                Update_ProductionMaterialsMissingRateControlTxt();
                updateProgress.UpdateProgress(70, "Update_SectorGeneralInforTxt...");
                Update_SectorGeneralInforTxt();
                updateProgress.UpdateProgress(80, "Update_WareHouseMaterialsOnTrayNonProgramControlTxt...");
                Update_WareHouseMaterialsOnTrayNonProgramControlTxt();
                updateProgress.UpdateProgress(90, "Update_WareHouseMaterialsProgrammingControlTxt...");
                Update_WareHouseMaterialsProgrammingControlTxt();
                updateProgress.UpdateProgress(95, "Update_WareHouseMaterialsSourceTxt...");
                Update_WareHouseMaterialsSourceTxt();
                updateProgress.UpdateProgress(100, "All done!!!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating database: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updateProgress.Dispose();
                updateProgress.Close();
            }       
        }

        private void Setup_dgvPartRestOfPO()
        {

        }


        /// <summary>
        /// This is done in frmMPHII (Code) Form_Load in VB6.
        /// </summary>
        private void GetUserNameAndSetup()
        {
            // Implementation for getting username and setup
            _strUserName = Environment.UserName;
            _strUserName = _strUserName.ToUpper().Trim();
        }

        private void Populate_cbbPulledListLine()
        {
            // Implementation for populating cbbPulledListLine ComboBox
            cbbPulledListLine.Items.Clear();
            List<string> line = new List<string>
            {
                "SMTA_Line",
                "SMTB_Line",
                "SMTC_Line",
                "SMTD_Line",
                "SMTE_Line",
                "SMTF_Line",
                "SMTG_Line"
            };
            cbbPulledListLine.Items.Clear();
            cbbPulledListLine.Items.AddRange(line.ToArray());
        }

        private void RequireOutlookOpen()
        {
            // Implementation for requiring Outlook to be open
        }

        /// <summary>
        /// This will handle the event when the selection in the cbbLine ComboBox is changed.
        /// Load the PO Pulled List based on the selected line.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbbLine_SelectionChangeCommitted(object sender, EventArgs e) //VB6: Private Sub lstPulledListLine_Click()
        {
            // Disable the buttons until processing is done
            LockSelection();

            // Set the ModelvsMaterialsLayoutDB variable
            _modelVsMaterialsLayoutDB = "ModelvsPhysicalMaterials_" + cbbPulledListLine.Text.Trim();

            //if (cbbPulledListLine.Text.Trim().Substring(0,3).ToUpper() == "SMT")
            //{
            //    //VB6 this will do some set up, doing nothing for now
            //}
            //FormatString = "^No.|^PONumber|^ModelNumber|^TOP/BOT|^PO Qty|^PulledListID|^Planner'sNote|^POChangeInf"
            string newValue = cbbPulledListLine.SelectedItem.ToString(); // Get the updated value
            _pulledList_SectorGroup = GetSectorGroup(newValue);
            _pulledList_Type = GetPulledListType(newValue);
            LoadActiveDateTime(newValue);

            // Enable the buttons after processing is done:
            UnlockSelection();
        }

        /// <summary>
        /// Input sector, get group of that sector from SectorGeneralInfor.txt file.
        /// </summary>
        /// <param name="sector"></param>
        /// <returns></returns>
        private string GetSectorGroup(string sector)
        {
            //Public Function getSectorType(getSector As String) As String
            //Public Function GetSectorGroup(getSector As String) As String
            string sectorGroup = "NA";
            string targetFile = @"C:\MPH - KANBAN Control Local Data\SectorGeneralInfor.txt";

            if (string.IsNullOrEmpty(sector))
            {
                return sectorGroup;
            }         

            if (!File.Exists(targetFile))
            {
                return sectorGroup;
            }

            try
            {
                using (var sr = new StreamReader(targetFile))
                {
                    string line;
                    int rowStart = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        // split on tab (vbTab)
                        var fields = line.Split('\t');

                        // Skip header line (rowStart == 0) and stop if first field is "End"
                        if (rowStart > 0 && fields.Length > 0 && !string.Equals(fields[0].Trim(), "End", StringComparison.OrdinalIgnoreCase))
                        {
                            var getLocalSector = fields[0].Trim(); //SMTA_Line, SMTB_Line, ...

                            if (string.Equals(sector, getLocalSector, StringComparison.Ordinal))
                            {
                                // VB6 used index 5 (6th column). Return trimmed value if present.
                                if (fields.Length >= 5)
                                    return fields[5].Trim(); //SMT_Group
                                else
                                    return sectorGroup;
                            }
                        }
                        rowStart++;
                    }
                }
            }
            catch
            {
                // Preserve original behavior of returning "NA" on errors instead of throwing.
                return sectorGroup;
            }

            return sectorGroup;
        }

        /// <summary>
        /// Set the _pulledList_Type variable based on sector from SectorGeneralInfor.txt file.
        /// </summary>
        /// <param name="sector"></param>
        /// <returns></returns>
        private string GetPulledListType(string sector)
        {
            string pulledListType = "NA";
            string targetFile = @"C:\MPH - KANBAN Control Local Data\SectorGeneralInfor.txt";

            if (string.IsNullOrEmpty(sector))
            {
                return pulledListType;
            }

            if (!File.Exists(targetFile))
            {
                return pulledListType;
            }

            try
            {
                using (var sr = new StreamReader(targetFile))
                {
                    string line;
                    int rowStart = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        // split on tab (vbTab)
                        var fields = line.Split('\t');

                        // Skip header line (rowStart == 0) and stop if first field is "End"
                        if (rowStart > 0 && fields.Length > 0 && !string.Equals(fields[0].Trim(), "End", StringComparison.OrdinalIgnoreCase))
                        {
                            var getLocalSector = fields[0].Trim(); //SMTA_Line, SMTB_Line, ...

                            if (string.Equals(sector, getLocalSector, StringComparison.Ordinal))
                            {
                                // VB6 used index 7 (8th column). Return trimmed value if present.
                                if (fields.Length >= 7)
                                    return fields[7].Trim(); //SMT = DividePOByPOnLoc; (DividePOByPO, GroupPOByList,...)
                                else
                                    return pulledListType;
                            }
                        }
                        rowStart++;
                    }
                }
            }
            catch
            {
                // Preserve original behavior of returning "NA" on errors instead of throwing.
                return pulledListType;
            }

            return pulledListType;
        }

        private int LoadActiveDateTime(string sector) //VB6: Public Sub LoadActivatedDatevsLine(getSector As String)
        {
            int selectedIndex = -1;
            int ii = 0;
            int getListIndex = 0;

            cbbActiveDate.Items.Clear();

            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;

            const string query = @"
                SELECT DISTINCT ActiveDateTime
                FROM OpenPOPlanner
                WHERE TypeCO <> 'None'
                  AND Priority <> '0'
                  AND TypeCO <> 'Pending PO'
                  AND TypeCO <> 'Done PO'
                  AND Sector = @sector
                ORDER BY ActiveDateTime";
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@sector", sector ?? (object)DBNull.Value);
                        conn.Open();
                        using (var rdr = cmd.ExecuteReader())
                        {
                            if (rdr.HasRows)
                            {
                                System.DateTime tomorrow = System.DateTime.Now.Date.AddDays(1);
                                while (rdr.Read())
                                {
                                    ii++;
                                    // read as DateTime if possible
                                    System.DateTime activeDate;
                                    object val = rdr["ActiveDateTime"];
                                    if (val == DBNull.Value)
                                    {
                                        // skip nulls
                                        continue;
                                    }
                                    else if (val is System.DateTime dt)
                                    {
                                        activeDate = dt;
                                    }
                                    else
                                    {
                                        // fallback: try parsing
                                        if (!System.DateTime.TryParse(val.ToString(), out activeDate))
                                            continue;
                                    }
                                    // Format like VB6 "dd/mmm/yyyy"->use "dd/MMM/yyyy" with invariant culture for english month abbrev
                                    string formatted = activeDate.ToString("dd/MMM/yyyy", CultureInfo.InvariantCulture);
                                    cbbActiveDate.Items.Add(formatted);

                                    // Compare date-only to tomorrow
                                    if (activeDate.Date == tomorrow)
                                    {
                                        getListIndex = ii - 1; // VB6 used zero-based ListIndex but 1-based ii
                                    }
                                }
                            }
                        }
                    }
                }
                if (cbbActiveDate.Items.Count > 0)
                {
                    // ensure index is valid
                    selectedIndex = Math.Max(0, Math.Min(getListIndex, cbbActiveDate.Items.Count - 1));
                    cbbActiveDate.SelectedIndex = selectedIndex;
                }
            }
            catch (Exception ex)
            { 
                selectedIndex = -1;
                MessageBox.Show("Error loading activated dates: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return selectedIndex;
        }

        private string GetAllSector(string sector) //VB6: Public Function GetSectorGroup(getSector As String) As String
        {
            const string defaultResult = "NA";
            const string targetFile = @"C:\MPH - KANBAN Control Local Data\SectorGeneralInfor.txt";

            if (string.IsNullOrEmpty(sector))
                return defaultResult;

            if (!File.Exists(targetFile))
            {
                Update_SectorGeneralInforTxt();
                //return defaultResult;
            }

            try
            {
                using (var sr = new StreamReader(targetFile))
                {
                    string line;
                    int rowStart = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        var fields = line.Split('\t');

                        // Skip header line (rowStart == 0). Also skip lines where first field is "End".
                        if (rowStart > 0 && fields.Length > 0 && !string.Equals(fields[0].Trim(), "End", StringComparison.OrdinalIgnoreCase))
                        {
                            var getLocalSector = fields[0].Trim();

                            // VB6 compared sectors using default string equality; preserve ordinal comparison.
                            if (string.Equals(sector, getLocalSector, StringComparison.Ordinal))
                            {
                                // VB6 returned the 6th column (index 5).
                                return fields.Length > 5 ? fields[5].Trim() : defaultResult;
                            }
                        }

                        rowStart++;
                    }
                }
            }
            catch
            {
                // Keep behavior similar to original: return "NA" on any failure instead of throwing.
                return defaultResult;
            }

            return defaultResult;
        }

        private void GetDefectType()
        {

        }

        private void PartIOPIC()
        {

        }

        private void NumofTO()
        {

        }

        private void Setup_dgvPulledListPO() //frmMPHII.flxPulledListPO
        {
            // Set up columns
            dgvPulledListPO.Columns.Clear();
            dgvPulledListPO.AllowUserToAddRows = false;
            dgvPulledListPO.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^PONumber|^ModelNumber|^PO Qty|^PulledListID|^Planner'sNotice|^POChangeInf"
            //dgvPulledListPO: No(0); PONumber(1); ModelNumber(2); Side(3); POQty(4); PulledListID(5); PlannersNotice(6); POChangeInf(7)
            dgvPulledListPO.Columns.Add("No", "No.");
            dgvPulledListPO.Columns.Add("PONumber", "PONumber");
            dgvPulledListPO.Columns.Add("ModelNumber", "ModelNumber");
            dgvPulledListPO.Columns.Add("Side", "TOP/BOT");
            dgvPulledListPO.Columns.Add("POQty", "PO Qty");
            dgvPulledListPO.Columns.Add("PulledListID", "PulledListID");
            dgvPulledListPO.Columns.Add("PlannersNotice", "Planner'sNotice");
            dgvPulledListPO.Columns.Add("POChangeInf", "POChangeInf");

            // Set column widths (in pixels)
            dgvPulledListPO.Columns[0].Width = 50;   // No.
            dgvPulledListPO.Columns[1].Width = 250;  // PONumber
            dgvPulledListPO.Columns[2].Width = 125;  // ModelNumber
            dgvPulledListPO.Columns[3].Width = 80;   // TOP/BOT
            dgvPulledListPO.Columns[4].Width = 120;   // PO Qty
            dgvPulledListPO.Columns[5].Width = 140;  // PulledListID
            dgvPulledListPO.Columns[6].Width = 180;  // Planner'sNotice
            dgvPulledListPO.Columns[7].Width = 150;  // POChangeInf

            // Set column alignment (center left)
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 7; idx++)
            {
                dgvPulledListPO.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row
            dgvPulledListPO.EnableHeadersVisualStyles = false; // Allow custom header styles
            dgvPulledListPO.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvPulledListPO.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvPulledListPO.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvPulledListPO.Font, FontStyle.Bold);
        }

        private void Setup_dgvPhysicalModelRunningPulled()
        {
            // Set up columns for dgvPhysicalModelRunningPulled
            dgvPhysicalModelRunningPulled.Columns.Clear();
            dgvPhysicalModelRunningPulled.AllowUserToAddRows = false;
            dgvPhysicalModelRunningPulled.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^Part Name|^Part Description|^Qty|^UOM"
            dgvPhysicalModelRunningPulled.Columns.Add("No", "No.");
            dgvPhysicalModelRunningPulled.Columns.Add("PartName", "Part.Name");
            dgvPhysicalModelRunningPulled.Columns.Add("PartDesc", "Part.Desc");
            dgvPhysicalModelRunningPulled.Columns.Add("Qty", "Qty");
            dgvPhysicalModelRunningPulled.Columns.Add("UOM", "UOM");

            // Set column widths (pixels, adjust as needed)
            dgvPhysicalModelRunningPulled.Columns[0].Width = 50;    // No.
            dgvPhysicalModelRunningPulled.Columns[1].Width = 100;    // Part.Name
            dgvPhysicalModelRunningPulled.Columns[2].Width = 50;    // Part.Desc
            dgvPhysicalModelRunningPulled.Columns[3].Width = 60;    // Qty
            dgvPhysicalModelRunningPulled.Columns[4].Width = 60;   // UOM

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 4; idx++)
            {
                dgvPhysicalModelRunningPulled.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvPhysicalModelRunningPulled.EnableHeadersVisualStyles = false;
            dgvPhysicalModelRunningPulled.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvPhysicalModelRunningPulled.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvPhysicalModelRunningPulled.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvPhysicalModelRunningPulled.Font, FontStyle.Bold);
        }

        private void Setup_dgvPhysicalSAPModelAfterCOPulled()
        {
            // Set up columns for dgvPhysicalSAPModelAfterCOPulled
            dgvPhysicalSAPModelAfterCOPulled.Columns.Clear();
            dgvPhysicalSAPModelAfterCOPulled.AllowUserToAddRows = false;
            dgvPhysicalSAPModelAfterCOPulled.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^Part Name|^Part Description|^Qty|^UOM"
            dgvPhysicalSAPModelAfterCOPulled.Columns.Add("No", "No.");
            dgvPhysicalSAPModelAfterCOPulled.Columns.Add("PartName", "Part.Name");
            dgvPhysicalSAPModelAfterCOPulled.Columns.Add("PartDesc", "Part.Desc");
            dgvPhysicalSAPModelAfterCOPulled.Columns.Add("Qty", "Qty");
            dgvPhysicalSAPModelAfterCOPulled.Columns.Add("UOM", "UOM");

            // Set column widths (pixels, adjust as needed)
            dgvPhysicalSAPModelAfterCOPulled.Columns[0].Width = 50;    // No.
            dgvPhysicalSAPModelAfterCOPulled.Columns[1].Width = 100;    // Part.Name
            dgvPhysicalSAPModelAfterCOPulled.Columns[2].Width = 50;    // Part.Desc
            dgvPhysicalSAPModelAfterCOPulled.Columns[3].Width = 60;    // Qty
            dgvPhysicalSAPModelAfterCOPulled.Columns[4].Width = 60;   // UOM

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 4; idx++)
            {
                dgvPhysicalSAPModelAfterCOPulled.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvPhysicalSAPModelAfterCOPulled.EnableHeadersVisualStyles = false;
            dgvPhysicalSAPModelAfterCOPulled.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvPhysicalSAPModelAfterCOPulled.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvPhysicalSAPModelAfterCOPulled.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvPhysicalSAPModelAfterCOPulled.Font, FontStyle.Bold);
        }

        private void Setup_dgvPullListvsPO()
        {
            //'flxPullListvsPO.FormatString = "^No.|^PONumber|^Model|^Part Name|^Part Description|^Qty|^UOM|^NextPOUsed|^Materials Class|^TopBot"
            // Set up columns for dgvPullListvsPO
            dgvPullListvsPO.Columns.Clear();
            dgvPullListvsPO.AllowUserToAddRows = false;
            dgvPullListvsPO.RowHeadersVisible = false;

            // Add columns with header text
            //dgvPullListvsPO: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); NextPOUsed(7); MatClass(8); TopBot(9)
            dgvPullListvsPO.Columns.Add("No", "No.");
            dgvPullListvsPO.Columns.Add("PONumber", "PONumber");
            dgvPullListvsPO.Columns.Add("Model", "Model");
            dgvPullListvsPO.Columns.Add("PartName", "Part.Name");
            dgvPullListvsPO.Columns.Add("PartDesc", "Part.Desc");
            dgvPullListvsPO.Columns.Add("Qty", "Qty");
            dgvPullListvsPO.Columns.Add("UOM", "UOM");
            dgvPullListvsPO.Columns.Add("NextPOUsed", "NextPOUsed");
            dgvPullListvsPO.Columns.Add("MaterialClass", "Matrl.Class");
            dgvPullListvsPO.Columns.Add("TopBot", "TopBot");

            // Set column widths (pixels, adjust as needed)
            dgvPullListvsPO.Columns[0].Width = 50;    // No.
            dgvPullListvsPO.Columns[1].Width = 120;    // PONumber
            dgvPullListvsPO.Columns[2].Width = 120;   // Model
            dgvPullListvsPO.Columns[3].Width = 100;    // Part.Name
            dgvPullListvsPO.Columns[4].Width = 50;    // Part.Desc
            dgvPullListvsPO.Columns[5].Width = 60;    // Qty
            dgvPullListvsPO.Columns[6].Width = 60;   // UOM
            dgvPullListvsPO.Columns[7].Width = 140;    // NextPOUsed
            dgvPullListvsPO.Columns[8].Width = 150;    // MatClass
            dgvPullListvsPO.Columns[9].Width = 60;   // TopBot

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 9; idx++)
            {
                dgvPullListvsPO.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvPullListvsPO.EnableHeadersVisualStyles = false;
            dgvPullListvsPO.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvPullListvsPO.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvPullListvsPO.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvPullListvsPO.Font, FontStyle.Bold);
        }

        private void Setup_dgvPullListvsPO2()
        {
            // Set up columns for dgvPullListvsPO2
            dgvPullListvsPO2.Columns.Clear();
            dgvPullListvsPO2.AllowUserToAddRows = false;
            dgvPullListvsPO2.RowHeadersVisible = false;

            // Add columns with header text
            //dgvPullListvsPO2: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); NextPOUsed(7); MatClass(8); TopBot(9)
            dgvPullListvsPO2.Columns.Add("No", "No.");
            dgvPullListvsPO2.Columns.Add("PONumber", "PONumber");
            dgvPullListvsPO2.Columns.Add("Model", "Model");
            dgvPullListvsPO2.Columns.Add("PartName", "Part.Name");
            dgvPullListvsPO2.Columns.Add("PartDesc", "Part.Desc");
            dgvPullListvsPO2.Columns.Add("Qty", "Qty");
            dgvPullListvsPO2.Columns.Add("UOM", "UOM");
            dgvPullListvsPO2.Columns.Add("NextPOUsed", "NextPOUsed");
            dgvPullListvsPO2.Columns.Add("MatClass", "Matrl.Class");
            dgvPullListvsPO2.Columns.Add("TopBot", "TopBot");

            // Set column widths (pixels, adjust as needed)
            dgvPullListvsPO2.Columns[0].Width = 50;    // No.
            dgvPullListvsPO2.Columns[1].Width = 120;    // PONumber
            dgvPullListvsPO2.Columns[2].Width = 120;   // Model
            dgvPullListvsPO2.Columns[3].Width = 100;    // Part.Name
            dgvPullListvsPO2.Columns[4].Width = 50;    // Part.Desc
            dgvPullListvsPO2.Columns[5].Width = 60;    // Qty
            dgvPullListvsPO2.Columns[6].Width = 60;   // UOM
            dgvPullListvsPO2.Columns[7].Width = 140;    // NextPOUsed
            dgvPullListvsPO2.Columns[8].Width = 150;    // MatClass
            dgvPullListvsPO2.Columns[9].Width = 60;   // TopBot

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 9; idx++)
            {
                dgvPullListvsPO2.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvPullListvsPO2.EnableHeadersVisualStyles = false;
            dgvPullListvsPO2.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvPullListvsPO2.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvPullListvsPO2.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvPullListvsPO2.Font, FontStyle.Bold);
        }

        private void Setup_dgvPLOverallModelPulled() //Vb6: flxPLOverallModelPulled
        {
            // Set up columns for dgvPLOverallModelPulled
            dgvPLOverallModelPulled.Columns.Clear();
            dgvPLOverallModelPulled.AllowUserToAddRows = false;
            dgvPLOverallModelPulled.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^Part Name|^Part Description|^Part Qty|^Part UOM|^Mother Part"
            dgvPLOverallModelPulled.Columns.Add("No", "No.");
            dgvPLOverallModelPulled.Columns.Add("PartName", "Part.Name");
            dgvPLOverallModelPulled.Columns.Add("PartDesc", "Part.Desc");
            dgvPLOverallModelPulled.Columns.Add("Qty", "Qty");
            dgvPLOverallModelPulled.Columns.Add("UOM", "UOM");
            dgvPLOverallModelPulled.Columns.Add("MotherPart", "MotherPart");

            // Set column widths (pixels, adjust as needed)
            dgvPLOverallModelPulled.Columns[0].Width = 50;    // No.
            dgvPLOverallModelPulled.Columns[1].Width = 150;    // PartName
            dgvPLOverallModelPulled.Columns[2].Width = 150;   // PartDesc
            dgvPLOverallModelPulled.Columns[3].Width = 100;    // Qty
            dgvPLOverallModelPulled.Columns[4].Width = 100;    // UOM
            dgvPLOverallModelPulled.Columns[5].Width = 150;    // MotherPart

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 5; idx++)
            {
                dgvPLOverallModelPulled.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvPLOverallModelPulled.EnableHeadersVisualStyles = false;
            dgvPLOverallModelPulled.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvPLOverallModelPulled.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvPLOverallModelPulled.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvPLOverallModelPulled.Font, FontStyle.Bold);
        }

        private void Setup_dgvUniPhysicalModelPulled() //VB6: flxUniPhysicalModelPulled
        {
            // Set up columns for dgvUniPhysicalModelPulled
            dgvUniPhysicalModelPulled.Columns.Clear();
            dgvUniPhysicalModelPulled.AllowUserToAddRows = false;
            dgvUniPhysicalModelPulled.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^PONumber|^Model|^Part Name|^Part Description|^Part Qty|^Part UOM|^PCBABin|^TopBot"
            //dgvUniPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); PCBABin(7); TopBot(8)
            dgvUniPhysicalModelPulled.Columns.Add("No", "No.");
            dgvUniPhysicalModelPulled.Columns.Add("PONumber", "PONumber");
            dgvUniPhysicalModelPulled.Columns.Add("Model", "Model");
            dgvUniPhysicalModelPulled.Columns.Add("PartName", "Part.Name");
            dgvUniPhysicalModelPulled.Columns.Add("PartDesc", "Part.Desc");
            dgvUniPhysicalModelPulled.Columns.Add("Qty", "Qty");
            dgvUniPhysicalModelPulled.Columns.Add("UOM", "UOM");
            dgvUniPhysicalModelPulled.Columns.Add("PCBABin", "PCBABin");
            dgvUniPhysicalModelPulled.Columns.Add("TopBot", "TopBot");

            // Set column widths (pixels, adjust as needed)
            dgvUniPhysicalModelPulled.Columns[0].Width = 50;    // No.
            dgvUniPhysicalModelPulled.Columns[1].Width = 120;    // PONumber
            dgvUniPhysicalModelPulled.Columns[2].Width = 120;   // Model
            dgvUniPhysicalModelPulled.Columns[3].Width = 100;    // Part.Name
            dgvUniPhysicalModelPulled.Columns[4].Width = 50;    // Part.Desc
            dgvUniPhysicalModelPulled.Columns[5].Width = 60;    // Qty
            dgvUniPhysicalModelPulled.Columns[6].Width = 60;   // UOM
            dgvUniPhysicalModelPulled.Columns[7].Width = 140;    // PCBABin
            dgvUniPhysicalModelPulled.Columns[8].Width = 60;   // TopBot

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 8; idx++)
            {
                dgvUniPhysicalModelPulled.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvUniPhysicalModelPulled.EnableHeadersVisualStyles = false;
            dgvUniPhysicalModelPulled.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvUniPhysicalModelPulled.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvUniPhysicalModelPulled.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvUniPhysicalModelPulled.Font, FontStyle.Bold);
        }

        private void Setup_dgvPartvsQty()
        {
            
        }

        private void Setup_dgvQtyvsCountDuplicated() //VB6: flxQtyvsCountDuplicated
        {
            // Set up columns for dgvQtyvsCountDuplicated
            dgvQtyvsCountDuplicated.Columns.Clear();
            dgvQtyvsCountDuplicated.AllowUserToAddRows = false;
            dgvQtyvsCountDuplicated.RowHeadersVisible = false;

            // Add columns with header text
            // .FormatString = "^No.|^Q.ty|^Count"
            dgvQtyvsCountDuplicated.Columns.Add("No", "No.");
            dgvQtyvsCountDuplicated.Columns.Add("Qty", "Qty");
            dgvQtyvsCountDuplicated.Columns.Add("Count", "Count");

            // Set column widths (pixels, adjust as needed)
            dgvQtyvsCountDuplicated.Columns[0].Width = 100;    // No.
            dgvQtyvsCountDuplicated.Columns[1].Width = 100;    // Qty
            dgvQtyvsCountDuplicated.Columns[2].Width = 100;   // Count

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 2; idx++)
            {
                dgvQtyvsCountDuplicated.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvQtyvsCountDuplicated.EnableHeadersVisualStyles = false;
            dgvQtyvsCountDuplicated.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvQtyvsCountDuplicated.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvQtyvsCountDuplicated.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvQtyvsCountDuplicated.Font, FontStyle.Bold);
        }

        private void Setup_dgvKANBANPulledList() //VB6: flxKANBANPulledList
        {
            // Set up columns for dgvKANBANPulledList
            dgvKANBANPulledList.Columns.Clear();
            dgvKANBANPulledList.AllowUserToAddRows = false;
            dgvKANBANPulledList.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^PONumber|^Model|^Mat.Name|^Mat.Desc|^Mat.QTY|^Mat.DEST|^UOM|^FMR|^BoxType|^Mat.Source"
            dgvKANBANPulledList.Columns.Add("No", "No.");
            dgvKANBANPulledList.Columns.Add("PONumber", "PONumber");
            dgvKANBANPulledList.Columns.Add("Model", "Model");
            dgvKANBANPulledList.Columns.Add("MatName", "Mat.Name");
            dgvKANBANPulledList.Columns.Add("MatDesc", "Mat.Desc");
            dgvKANBANPulledList.Columns.Add("MatQTY", "Mat.QTY");
            dgvKANBANPulledList.Columns.Add("MatDEST", "Mat.DEST");
            dgvKANBANPulledList.Columns.Add("UOM", "UOM");
            dgvKANBANPulledList.Columns.Add("FMR", "FMR");
            dgvKANBANPulledList.Columns.Add("BoxType", "BoxType");
            dgvKANBANPulledList.Columns.Add("MatSource", "Mat.Source");

            // Set column widths (pixels, adjust as needed)
            dgvKANBANPulledList.Columns[0].Width = 46;    // No.
            dgvKANBANPulledList.Columns[1].Width = 90;    // PONumber
            dgvKANBANPulledList.Columns[2].Width = 150;   // Model
            dgvKANBANPulledList.Columns[3].Width = 90;    // Mat.Name
            dgvKANBANPulledList.Columns[4].Width = 50;    // Mat.Desc
            dgvKANBANPulledList.Columns[5].Width = 60;    // Mat.QTY
            dgvKANBANPulledList.Columns[6].Width = 120;   // Mat.DEST
            dgvKANBANPulledList.Columns[7].Width = 30;    // UOM
            dgvKANBANPulledList.Columns[8].Width = 30;    // FMR
            dgvKANBANPulledList.Columns[9].Width = 90;    // BoxType
            dgvKANBANPulledList.Columns[10].Width = 20;   // Mat.Source

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 10; idx++)
            {
                dgvKANBANPulledList.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvKANBANPulledList.EnableHeadersVisualStyles = false;
            dgvKANBANPulledList.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvKANBANPulledList.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvKANBANPulledList.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvKANBANPulledList.Font, FontStyle.Bold);
        }

        private void Setup_dgvMaterialsConversionMatrix() //VB6: With flxMaterialsConversionMatrix
        {
            // Set up columns for dgvMaterialsConversionMatrix
            dgvMaterialsConversionMatrix.Columns.Clear();
            dgvMaterialsConversionMatrix.AllowUserToAddRows = false;
            dgvMaterialsConversionMatrix.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^MaterialsBef|^MaterialsAft|^PONumber|^POModel|^POQty|^WSCode|^WSLocCode|^QtyPerLoc|^QtyPerPOLoc|^InStockLoc|^TopBot"
            //dgvMaterialsConversionMatrix: No(0), MaterialsBef(1), MaterialsAft(2), PONumber(3), POModel(4), POQty(5), WSCode(6), WSLocCode(7), QtyPerLoc(8), QtyPerPOLoc(9), InStockLoc(10), TopBot(11)
            dgvMaterialsConversionMatrix.Columns.Add("No", "No."); //0
            dgvMaterialsConversionMatrix.Columns.Add("MaterialsBef", "MaterialsBef"); //1
            dgvMaterialsConversionMatrix.Columns.Add("MaterialsAft", "MaterialsAft"); //2
            dgvMaterialsConversionMatrix.Columns.Add("PONumber", "PONumber"); //3
            dgvMaterialsConversionMatrix.Columns.Add("POModel", "POModel"); //4
            dgvMaterialsConversionMatrix.Columns.Add("POQty", "POQty"); //5
            dgvMaterialsConversionMatrix.Columns.Add("WSCode", "WSCode"); //6
            dgvMaterialsConversionMatrix.Columns.Add("WSLocCode", "WSLocCode"); //7
            dgvMaterialsConversionMatrix.Columns.Add("QtyPerLoc", "QtyPerLoc"); //8
            dgvMaterialsConversionMatrix.Columns.Add("QtyPerPOLoc", "QtyPerPOLoc"); //9
            dgvMaterialsConversionMatrix.Columns.Add("InStockLoc", "InStockLoc"); //10
            dgvMaterialsConversionMatrix.Columns.Add("TopBot", "TopBot"); //11

            // Set column widths (pixels, adjust as needed for your UI)
            dgvMaterialsConversionMatrix.Columns[0].Width = 46;    // No.
            dgvMaterialsConversionMatrix.Columns[1].Width = 120;   // MaterialsBef
            dgvMaterialsConversionMatrix.Columns[2].Width = 120;   // MaterialsAft
            dgvMaterialsConversionMatrix.Columns[3].Width = 120;   // PONumber
            dgvMaterialsConversionMatrix.Columns[4].Width = 120;   // POModel
            dgvMaterialsConversionMatrix.Columns[5].Width = 69;    // POQty
            dgvMaterialsConversionMatrix.Columns[6].Width = 120;   // WSCode
            dgvMaterialsConversionMatrix.Columns[7].Width = 120;   // WSLocCode
            dgvMaterialsConversionMatrix.Columns[8].Width = 120;   // QtyPerLoc
            dgvMaterialsConversionMatrix.Columns[9].Width = 150;   // QtyPerPOLoc
            dgvMaterialsConversionMatrix.Columns[10].Width = 120;  // InStockLoc
            dgvMaterialsConversionMatrix.Columns[11].Width = 69;   // TopBot

            // Style the header row (blue background, bold, white text)
            dgvMaterialsConversionMatrix.EnableHeadersVisualStyles = false;
            dgvMaterialsConversionMatrix.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvMaterialsConversionMatrix.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvMaterialsConversionMatrix.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvMaterialsConversionMatrix.Font, FontStyle.Bold);
        }

        private void Setup_dgvMaterialsOnTrayNonProgramMatrix() //Vb6: With flxMaterialsOnTrayNonProgramMatrix
        {
            // Set up columns for dgvMaterialsOnTrayNonProgramMatrix
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Clear();
            dgvMaterialsOnTrayNonProgramMatrix.AllowUserToAddRows = false;
            dgvMaterialsOnTrayNonProgramMatrix.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^Materials|^PONumber|^POModel|^POQty|^WSCode|^WSLocCode|^QtyPerLoc|^QtyPerPOLoc|^InStockLoc|^TopBot"
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("No", "No."); //0
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("Materials", "Materials"); //1
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("PONumber", "PONumber"); //2
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("POModel", "POModel"); //3
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("POQty", "POQty"); //4
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("WSCode", "WSCode"); //5
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("WSLocCode", "WSLocCode"); //6
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("QtyPerLoc", "QtyPerLoc"); //7
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("QtyPerPOLoc", "QtyPerPOLoc"); //8
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("InStockLoc", "InStockLoc"); //9
            dgvMaterialsOnTrayNonProgramMatrix.Columns.Add("TopBot", "TopBot"); //10

            // Set column widths (pixels, adjust as needed for your layout)
            dgvMaterialsOnTrayNonProgramMatrix.Columns[0].Width = 46;    // No.
            dgvMaterialsOnTrayNonProgramMatrix.Columns[1].Width = 120;   // Materials
            dgvMaterialsOnTrayNonProgramMatrix.Columns[2].Width = 120;   // PONumber
            dgvMaterialsOnTrayNonProgramMatrix.Columns[3].Width = 120;   // POModel
            dgvMaterialsOnTrayNonProgramMatrix.Columns[4].Width = 69;    // POQty
            dgvMaterialsOnTrayNonProgramMatrix.Columns[5].Width = 120;   // WSCode
            dgvMaterialsOnTrayNonProgramMatrix.Columns[6].Width = 120;   // WSLocCode
            dgvMaterialsOnTrayNonProgramMatrix.Columns[7].Width = 120;   // QtyPerLoc
            dgvMaterialsOnTrayNonProgramMatrix.Columns[8].Width = 150;   // QtyPerPOLoc
            dgvMaterialsOnTrayNonProgramMatrix.Columns[9].Width = 120;   // InStockLoc
            dgvMaterialsOnTrayNonProgramMatrix.Columns[10].Width = 69;   // TopBot

            // Style the header row (blue background, bold, white text)
            dgvMaterialsOnTrayNonProgramMatrix.EnableHeadersVisualStyles = false;
            dgvMaterialsOnTrayNonProgramMatrix.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvMaterialsOnTrayNonProgramMatrix.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvMaterialsOnTrayNonProgramMatrix.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvMaterialsOnTrayNonProgramMatrix.Font, FontStyle.Bold);
        }
            
        private void Setup_dgvPLPhysicalModelPulled() //VB6: flxPLPhysicalModelPulled
        {
            // Set up columns for dgvPLPhysicalModelPulled
            dgvPLPhysicalModelPulled.Columns.Clear();
            dgvPLPhysicalModelPulled.AllowUserToAddRows = false;
            dgvPLPhysicalModelPulled.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^PONumber|^Model|^Part Name|^Part Description|^Part Qty|^Part UOM|^TopBot"
            //dgvPLPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); TopBot(7)
            dgvPLPhysicalModelPulled.Columns.Add("No", "No.");
            dgvPLPhysicalModelPulled.Columns.Add("PONumber", "PONumber");
            dgvPLPhysicalModelPulled.Columns.Add("Model", "Model");
            dgvPLPhysicalModelPulled.Columns.Add("PartName", "Part.Name");
            dgvPLPhysicalModelPulled.Columns.Add("PartDesc", "Part.Desc");
            dgvPLPhysicalModelPulled.Columns.Add("Qty", "Qty");
            dgvPLPhysicalModelPulled.Columns.Add("UOM", "UOM");
            dgvPLPhysicalModelPulled.Columns.Add("TopBot", "TopBot");

            // Set column widths (pixels, adjust as needed)
            dgvPLPhysicalModelPulled.Columns[0].Width = 50;    // No.
            dgvPLPhysicalModelPulled.Columns[1].Width = 120;    // PONumber
            dgvPLPhysicalModelPulled.Columns[2].Width = 120;   // Model
            dgvPLPhysicalModelPulled.Columns[3].Width = 100;    // Part.Name
            dgvPLPhysicalModelPulled.Columns[4].Width = 50;    // Part.Desc
            dgvPLPhysicalModelPulled.Columns[5].Width = 60;    // Qty
            dgvPLPhysicalModelPulled.Columns[6].Width = 60;   // UOM
            dgvPLPhysicalModelPulled.Columns[7].Width = 60;   // TopBot

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 7; idx++)
            {
                dgvPLPhysicalModelPulled.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvPLPhysicalModelPulled.EnableHeadersVisualStyles = false;
            dgvPLPhysicalModelPulled.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvPLPhysicalModelPulled.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvPLPhysicalModelPulled.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvPLPhysicalModelPulled.Font, FontStyle.Bold);
        }

        private void Setup_dgvUniPOnQtyMaterials()
        {
            // Set up columns for dgvUniPOnQtyMaterials
            dgvUniPOnQtyMaterials.Columns.Clear();
            dgvUniPOnQtyMaterials.AllowUserToAddRows = false;
            dgvUniPOnQtyMaterials.RowHeadersVisible = false;

            // Add columns with header text
            //flxUniPOnQtyMaterials.FormatString = "^No.|^Materials|^PONumber|^POModel|^POQty|^WSCode|^WSLocCode|^QtyPerLoc|^QtyPerPOLoc|^InStockLoc|^TopBot"
            //dgvUniPOnQtyMaterials: No(0); Materials(1); PONumber(2); POModel(3); POQty(4); WSCode(5); WSLocCode(6); QtyPerLoc(7); QtyPerPOLoc(8); InStockLoc(9); TopBot(10); POMat(11)
            dgvUniPOnQtyMaterials.Columns.Add("No", "No.");
            dgvUniPOnQtyMaterials.Columns.Add("Materials", "Materials");
            dgvUniPOnQtyMaterials.Columns.Add("PONumber", "PONumber");
            dgvUniPOnQtyMaterials.Columns.Add("POModel", "POModel");
            dgvUniPOnQtyMaterials.Columns.Add("POQty", "POQty"); //4
            dgvUniPOnQtyMaterials.Columns.Add("WSCode", "WSCode");
            dgvUniPOnQtyMaterials.Columns.Add("WSLocCode", "WSLocCode");
            dgvUniPOnQtyMaterials.Columns.Add("QtyPerLoc", "QtyPerLoc"); //7
            dgvUniPOnQtyMaterials.Columns.Add("QtyPerPOLoc", "QtyPerPOLoc");
            dgvUniPOnQtyMaterials.Columns.Add("InStockLoc", "InStockLoc");
            dgvUniPOnQtyMaterials.Columns.Add("TopBot", "TopBot");
            dgvUniPOnQtyMaterials.Columns.Add("POMat", "POMat");

            // Set column widths (pixels, adjust as needed for your layout)
            dgvUniPOnQtyMaterials.Columns[0].Width = 40;    // No.
            dgvUniPOnQtyMaterials.Columns[1].Width = 150;   // Materials
            dgvUniPOnQtyMaterials.Columns[2].Width = 150;   // PONumber
            dgvUniPOnQtyMaterials.Columns[3].Width = 150;   // POModel
            dgvUniPOnQtyMaterials.Columns[4].Width = 60;
            dgvUniPOnQtyMaterials.Columns[5].Width = 80;
            dgvUniPOnQtyMaterials.Columns[6].Width = 80;
            dgvUniPOnQtyMaterials.Columns[7].Width = 80;
            dgvUniPOnQtyMaterials.Columns[8].Width = 80;
            dgvUniPOnQtyMaterials.Columns[9].Width = 80;
            dgvUniPOnQtyMaterials.Columns[10].Width = 80;
            dgvUniPOnQtyMaterials.Columns[11].Width = 120;  // POMat

            // Set alignment (MiddleLeft) for columns 1, 2, 3 (matching VB6 settings)
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 11; idx++)
            {
                dgvUniPOnQtyMaterials.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvUniPOnQtyMaterials.EnableHeadersVisualStyles = false;
            dgvUniPOnQtyMaterials.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvUniPOnQtyMaterials.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvUniPOnQtyMaterials.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvUniPOnQtyMaterials.Font, FontStyle.Bold);
        }

        private void Setup_dgvPhysicalModelAfterCOPulled()
        {
            /*With flxPhysicalModelAfterCOPulled
            .FixedRows = 1
            .FormatString = "^No.|^Part Name|^Part Description|^Qty|^UOM"*/
            // Set up columns for dgvMaterialsRarDivideKANBANBox
            dgvPhysicalModelAfterCOPulled.Columns.Clear();
            dgvPhysicalModelAfterCOPulled.AllowUserToAddRows = false;
            dgvPhysicalModelAfterCOPulled.RowHeadersVisible = false;

            // Add columns with header text
            dgvPhysicalModelAfterCOPulled.Columns.Add("No", "No.");
            dgvPhysicalModelAfterCOPulled.Columns.Add("PONumber", "PONumber");
            dgvPhysicalModelAfterCOPulled.Columns.Add("Model", "Model");
            dgvPhysicalModelAfterCOPulled.Columns.Add("PartName", "Part.Name");
            dgvPhysicalModelAfterCOPulled.Columns.Add("PartDesc", "Part.Desc");
            dgvPhysicalModelAfterCOPulled.Columns.Add("Qty", "Qty");
            dgvPhysicalModelAfterCOPulled.Columns.Add("UOM", "UOM");


            // Set column widths (pixels, adjust as needed)
            dgvPhysicalModelAfterCOPulled.Columns[0].Width = 50;    // No.
            dgvPhysicalModelAfterCOPulled.Columns[1].Width = 120;    // PONumber
            dgvPhysicalModelAfterCOPulled.Columns[2].Width = 1170;   // Model
            dgvPhysicalModelAfterCOPulled.Columns[3].Width = 100;    // Part.Name
            dgvPhysicalModelAfterCOPulled.Columns[4].Width = 50;    // Part.Desc
            dgvPhysicalModelAfterCOPulled.Columns[5].Width = 60;    // Qty
            dgvPhysicalModelAfterCOPulled.Columns[6].Width = 60;   // UOM

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 6; idx++)
            {
                dgvPhysicalModelAfterCOPulled.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvPhysicalModelAfterCOPulled.EnableHeadersVisualStyles = false;
            dgvPhysicalModelAfterCOPulled.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvPhysicalModelAfterCOPulled.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvPhysicalModelAfterCOPulled.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvPhysicalModelAfterCOPulled.Font, FontStyle.Bold);
        }

        private void Setup_dgvMaterialsRarDivideKANBANBox()
        {
            /*With flxMaterialsRarDivideKANBANBox
        .FixedRows = 1
        .FormatString = "^No.|^PONumber|^Model|^Part Name|^Part Description|^Qty|^UOM|^NextPOUsed|^Materials Class|^MaxKBBox"*/
            // Set up columns for dgvMaterialsRarDivideKANBANBox
            dgvMaterialsRarDivideKANBANBox.Columns.Clear();
            dgvMaterialsRarDivideKANBANBox.AllowUserToAddRows = false;
            dgvMaterialsRarDivideKANBANBox.RowHeadersVisible = false;

            // Add columns with header text
            dgvMaterialsRarDivideKANBANBox.Columns.Add("No", "No.");
            dgvMaterialsRarDivideKANBANBox.Columns.Add("PONumber", "PONumber");
            dgvMaterialsRarDivideKANBANBox.Columns.Add("Model", "Model");
            dgvMaterialsRarDivideKANBANBox.Columns.Add("PartName", "Part.Name");
            dgvMaterialsRarDivideKANBANBox.Columns.Add("PartDesc", "Part.Desc");
            dgvMaterialsRarDivideKANBANBox.Columns.Add("Qty", "Qty");
            dgvMaterialsRarDivideKANBANBox.Columns.Add("UOM", "UOM");
            dgvMaterialsRarDivideKANBANBox.Columns.Add("NextPOUsed", "NextPOUsed");
            dgvMaterialsRarDivideKANBANBox.Columns.Add("MaterialsClass", "Materials Class");
            dgvMaterialsRarDivideKANBANBox.Columns.Add("MaxKBBox", "MaxKBBox");

            // Set column widths (pixels, adjust as needed)
            dgvMaterialsRarDivideKANBANBox.Columns[0].Width = 50;    // No.
            dgvMaterialsRarDivideKANBANBox.Columns[1].Width = 120;    // PONumber
            dgvMaterialsRarDivideKANBANBox.Columns[2].Width = 1170;   // Model
            dgvMaterialsRarDivideKANBANBox.Columns[3].Width = 100;    // Part.Name
            dgvMaterialsRarDivideKANBANBox.Columns[4].Width = 50;    // Part.Desc
            dgvMaterialsRarDivideKANBANBox.Columns[5].Width = 60;    // Qty
            dgvMaterialsRarDivideKANBANBox.Columns[6].Width = 60;   // UOM
            dgvMaterialsRarDivideKANBANBox.Columns[7].Width = 60;   // NextPOUsed
            dgvMaterialsRarDivideKANBANBox.Columns[8].Width = 60;   // MaterialsClass
            dgvMaterialsRarDivideKANBANBox.Columns[9].Width = 60;   // MaxKBBox

            // Set alignment (MiddleLeft) for columns 1 to 10
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 9; idx++)
            {
                dgvMaterialsRarDivideKANBANBox.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvMaterialsRarDivideKANBANBox.EnableHeadersVisualStyles = false;
            dgvMaterialsRarDivideKANBANBox.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvMaterialsRarDivideKANBANBox.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvMaterialsRarDivideKANBANBox.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvMaterialsRarDivideKANBANBox.Font, FontStyle.Bold);
        }

        private void Setup_dgvPotentialIssues()
        {
            // Set up columns for dgvPotentialIssues
            dgvPotentialIssues.Columns.Clear();
            dgvPotentialIssues.AllowUserToAddRows = false;
            dgvPotentialIssues.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^POAfterCO|^Potential Issue"
            dgvPotentialIssues.Columns.Add("No", "No.");
            dgvPotentialIssues.Columns.Add("PotentialPONumber", "POAfterCO");
            dgvPotentialIssues.Columns.Add("PotentialPendingIssue", "Issue");

            // Set column widths (pixels, adjust as needed for your layout)
            dgvPotentialIssues.Columns[0].Width = 40;    // No.
            dgvPotentialIssues.Columns[1].Width = 129;   // PotentialPONumber
            dgvPotentialIssues.Columns[2].Width = 129;   // PotentialPendingIssue

            // Set alignment (MiddleLeft) for columns 1, 2, 3 (matching VB6 settings)
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 2; idx++)
            {
                dgvPotentialIssues.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvPotentialIssues.EnableHeadersVisualStyles = false;
            dgvPotentialIssues.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvPotentialIssues.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvPotentialIssues.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvPotentialIssues.Font, FontStyle.Bold);
        }

        private void Setup_dgvMultiUniPhysicalModelPulled()
        {
            // Set up columns for dgvMultiUniPhysicalModelPulled
            dgvMultiUniPhysicalModelPulled.Columns.Clear();
            dgvMultiUniPhysicalModelPulled.AllowUserToAddRows = false;
            dgvMultiUniPhysicalModelPulled.RowHeadersVisible = false;

            // Add columns with header text
            //.FormatString = "^No.|^PONumber|^Model|^Part Name|^Part Description|^Qty|^UOM|^PCBABin|^CommonPart|^TopBot"
            //dgvMultiUniPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDescription(4); Qty(5); UOM(6); PCBABin(7); CommonPart(8); TopBot(9)
            dgvMultiUniPhysicalModelPulled.Columns.Add("No", "No.");
            dgvMultiUniPhysicalModelPulled.Columns.Add("PONumber", "PONumber");
            dgvMultiUniPhysicalModelPulled.Columns.Add("Model", "Model");
            dgvMultiUniPhysicalModelPulled.Columns.Add("PartName", "Part Name");
            dgvMultiUniPhysicalModelPulled.Columns.Add("PartDescription", "Part Description");
            dgvMultiUniPhysicalModelPulled.Columns.Add("Qty", "Qty");
            dgvMultiUniPhysicalModelPulled.Columns.Add("UOM", "UOM");
            dgvMultiUniPhysicalModelPulled.Columns.Add("PCBABin", "PCBABin");
            dgvMultiUniPhysicalModelPulled.Columns.Add("CommonPart", "CommonPart");
            dgvMultiUniPhysicalModelPulled.Columns.Add("TopBot", "TopBot");

            // Set column widths (pixels, adjust as needed)
            dgvMultiUniPhysicalModelPulled.Columns[0].Width = 40;    // No.
            dgvMultiUniPhysicalModelPulled.Columns[1].Width = 129;   // PONumber
            dgvMultiUniPhysicalModelPulled.Columns[2].Width = 129;   // Model
            dgvMultiUniPhysicalModelPulled.Columns[3].Width = 129;   // Part Name
            dgvMultiUniPhysicalModelPulled.Columns[4].Width = 50;    // Part Description
            dgvMultiUniPhysicalModelPulled.Columns[5].Width = 60;    // Qty
            dgvMultiUniPhysicalModelPulled.Columns[6].Width = 60;    // UOM
            dgvMultiUniPhysicalModelPulled.Columns[7].Width = 160;   // PCBABin
            dgvMultiUniPhysicalModelPulled.Columns[8].Width = 180;   // CommonPart
                                                                     // Can set dgvMultiUniPhysicalModelPulled.Columns[9].Width for TopBot if needed

            // Set alignment (MiddleLeft) for specified columns
            DataGridViewCellStyle leftCenterStyle = new DataGridViewCellStyle();
            leftCenterStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            for (int idx = 1; idx <= 8; idx++)
            {
                dgvMultiUniPhysicalModelPulled.Columns[idx].DefaultCellStyle = leftCenterStyle;
            }

            // Style the header row (blue background, bold, white text)
            dgvMultiUniPhysicalModelPulled.EnableHeadersVisualStyles = false;
            dgvMultiUniPhysicalModelPulled.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Blue;
            dgvMultiUniPhysicalModelPulled.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvMultiUniPhysicalModelPulled.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgvMultiUniPhysicalModelPulled.Font, FontStyle.Bold);
        }

        private void GetDatabaseServer()
        {
            
        }

        private void GetDirChangeOverLog()
        {
            string targetFile = @"C:\MROform\ChangeOverLogDir.txt";
            string changeOverLogDir;

            if (File.Exists(targetFile))
            {
                using (StreamReader sr = new StreamReader(targetFile))
                {
                    changeOverLogDir = sr.ReadLine()?.Trim();
                }

                this.txtbInf.Text = "The Log of ChangeOver Is Loaded Into:";
                this.txtbCOLogDir.Text = changeOverLogDir;

                // Adjust width based on text length in pixels
                int textWidth = TextRenderer.MeasureText(this.txtbCOLogDir.Text, this.txtbCOLogDir.Font).Width;
                this.txtbCOLogDir.Width = textWidth;
            }
            else
            {
                MessageBox.Show(
                    "Chua Set-up Duong Dan Luu Tru ChangeOver Log Cho Chuong Trinh!!!",
                    "ChangeOver Log Set-up Missing!!!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                Application.Exit();
            }
        }

        private void GetLabelRow()
        {

        }

        /// <summary>
        /// VB6: cmdExport_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            // Variables declaration:
            string models_missing_layout_setup = "";

            // Disable buttons:
            LockSelection();

            // Show progress bar:
            progressBarForm.Show();

            // Verify the POs loaded in:
            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                string poNumber = dgvPulledListPO.Rows[i].Cells["PONumber"].Value.ToString();
                EnsurePOExistIn_dboPOpairunderCO(poNumber);
            }

            // Clear old database:
            progressBarForm.UpdateProgress(6, "Clearing old database...");//-- PROGRESS 6%
            ClearOldDatabase(); // "old" means >15 days

            // Copy BOM server to local folder:
            progressBarForm.UpdateProgress(9, "Copying BOM from shared drive to local folder...");
            CopyBOMTxtFilesFromSharedDriveToLocalFolder();



            // Check if all setup done for PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix:
            progressBarForm.UpdateProgress(12, "Check models layout setup...");
            if (_pulledList_Type == "DividePOByPOnLoc") //_pulledList_Type = GetPulledListType (cbbPulledListLine.Text), set everytime line changed
            {
                models_missing_layout_setup = CheckSetup_ModelvsMaterialsLayoutMatrix();
                if (models_missing_layout_setup == "No PO added")
                {
                    MessageBox.Show("No PO A");
                }
                if (models_missing_layout_setup != "NONE")
                {
                    progressBarForm.Close();
                    var result = MessageBox.Show("The Following Models are missing setup in PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix:" 
                        + Environment.NewLine + models_missing_layout_setup 
                        + "\r\nClick Yes to stop and check!" 
                        + "\r\nClick No to continue!", 
                        "Models Missing Setup Detected:", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                    if (result == DialogResult.Yes)
                    {
                        btnExport.Enabled = true;
                        return;
                    }
                }
            }

            // Block Planner User Right:
            progressBarForm.UpdateProgress(15, "Blocking planner user rights...");
            BlockPlannerUserRight();



            // Get sumPOValID - this is calculated to use as a unique ID for a pulled list (made from several POs added):
            progressBarForm.UpdateProgress(18, "Checking if BOM lists in local folder...");
            _sumPOValue = GetsumPOValID();



            // Is Pulled List in Local already???
            _isPulledListInLocalFolder = IsPulledListAlreadyInLocalFolder(cbbPulledListLine.Text, cbbPulledListShift.Text, Convert.ToString(_sumPOValue));
            if (_isPulledListInLocalFolder)
            {
                // Load Pulled List from Local Folder:
                progressBarForm.UpdateProgress(21, "Load Pulled List from Local Folder...");
                ProcessPulledListFromLocal();
                progressBarForm.UpdateProgress(94, "Load Pulled List from Local Folder...");//  PROGRESS 94%:
                if (_enableToOverrideWritePulledListToLocal)
                {
                    OverrideWritePulledListToLocal(cbbPulledListLine.Text.Trim(), cbbPulledListShift.Text.Trim() , Convert.ToString(_sumPOValue));
                }
            }
            else
            {
                // Process New Pulled List:
                progressBarForm.UpdateProgress(21, "Process new Pulled List...");
                ProcessNewPulledList();
            }

            progressBarForm.UpdateProgress(100, "Done exporting pulled list!"); // PROGRESS 100%:


            // Record Pulled List in [dbo].[PulledListPreAssyModelvsSectorLog]:
            if (_pulledList_Type == "DividePOByPOnLoc")
            {
                RecordPulledListPreAssyModelvsSector(cbbPulledListLine.Text.Trim());
            }

            // Sort dgvMultiUniPhysicalModelPulled by column 3 in ascending order
            dgvMultiUniPhysicalModelPulled.Sort(
                dgvMultiUniPhysicalModelPulled.Columns[2], // Index 2 because columns are zero-based in C#
                System.ComponentModel.ListSortDirection.Ascending
            );

            _blnMultiPOExport = false; // Reset flag
            // Enable buttons:
            UnlockSelection();
        }

        /// <summary>
        /// VB6: Public Sub OverrideWritePulledListToLocal(getSector As String, getShift As String, getSumPOValID As String);
        /// 
        /// </summary>
        /// <param name="sector"></param>
        /// <param name="shift"></param>
        /// <param name="sumPOValID"></param>
        private void OverrideWritePulledListToLocal(string sector, string shift, string sumPOValID)
        {
            string baseFolder = @"C:\MPH - KANBAN Control Local Data\PulledListLog";

            if (sector == null) sector = string.Empty;
            if (shift == null) shift = string.Empty;
            if (sumPOValID == null) sumPOValID = string.Empty;

            // Build file name like VB: "PulledList_<sectorTrim>_<shiftChar>_<sum>_WithoutPK.txt"
            // VB used Mid(sector, 1, Len(sector) - 5) and Mid(shift, 7, 1)
            string sectorPart;
            try
            {
                var sectorStr = sector;
                sectorPart = sectorStr.Length > 5 ? sectorStr.Substring(0, sectorStr.Length - 5) : sectorStr;
            }
            catch
            {
                sectorPart = sector;
            }

            string shiftChar;
            try
            {
                shiftChar = shift.Length >= 7 ? shift.Substring(6, 1) : shift;
            }
            catch
            {
                shiftChar = shift;
            }

            string fileName = $"PulledList_{sectorPart}_{shiftChar}_{sumPOValID}_WithoutPK.txt";
            string targetFile = System.IO.Path.Combine(baseFolder, fileName);

            try
            {
                // Ensure directory exists
                if (!Directory.Exists(baseFolder))
                    Directory.CreateDirectory(baseFolder);

                // Write header (overwrite any existing file)
                string header = "No\tPO Number\tModel\tMaterials\tMaterial Description\tQuantity\tUOM\tMaterials Class\tPriority\tTONumber\tTOSt.Type\tTOSt.Bin\tTOSt.TypeTOSt.BinMaterials";
                File.WriteAllText(targetFile, header + Environment.NewLine);

                // Open for appending the data rows
                using (var sw = new StreamWriter(targetFile, append: true, encoding: System.Text.Encoding.UTF8))
                {
                    int mm = 0;

                    // Helper local to process a given grid (matching VB's repeated code)
                    void ProcessGrid(dynamic grid)
                    {
                        if (grid == null) return;

                        int rows;
                        try
                        {
                            rows = Convert.ToInt32(grid.Rows);
                        }
                        catch
                        {
                            // If the grid doesn't expose Rows as expected, skip it
                            return;
                        }

                        // VB loop: For j = 1 To .Rows - 2
                        for (int j = 1; j <= rows - 2; j++)
                        {
                            var flagObj = grid.TextMatrix(j, 11);
                            string flag = flagObj?.ToString().Trim() ?? string.Empty;
                            if (!string.Equals(flag, "YES", StringComparison.OrdinalIgnoreCase))
                                continue;

                            mm++;
                            int getOrder = mm;
                            string getPONumber = (grid.TextMatrix(j, 1)?.ToString() ?? string.Empty).Trim();
                            string getModel = (grid.TextMatrix(j, 2)?.ToString() ?? string.Empty).Trim();
                            string getMaterials = (grid.TextMatrix(j, 3)?.ToString() ?? string.Empty).Trim();
                            string getMaterialsDesc = (grid.TextMatrix(j, 4)?.ToString() ?? string.Empty).Trim();

                            // VB used Val(...) which is lenient. Try parse double, fallback 0.
                            double getQTy = 0.0;
                            var qtyObj = grid.TextMatrix(j, 5);
                            string qtyText = qtyObj?.ToString().Trim() ?? string.Empty;
                            if (!double.TryParse(qtyText, NumberStyles.Any, CultureInfo.InvariantCulture, out getQTy))
                            {
                                // try removing commas / other culture formats
                                var cleaned = qtyText.Replace(",", "");
                                double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out getQTy);
                            }

                            string getUOM = (grid.TextMatrix(j, 6)?.ToString() ?? string.Empty).Trim();
                            string getMaterialsClass = (grid.TextMatrix(j, 8)?.ToString() ?? string.Empty).Trim();
                            string getPriority = (grid.TextMatrix(j, 9)?.ToString() ?? string.Empty).Trim();

                            string tonCell = (grid.TextMatrix(j, 10)?.ToString() ?? string.Empty).Trim();
                            string getTONumber = string.IsNullOrWhiteSpace(tonCell) ? "None" : tonCell;

                            string getTOStType = (grid.TextMatrix(j, 12)?.ToString() ?? string.Empty).Trim();
                            string getTOStBin = (grid.TextMatrix(j, 13)?.ToString() ?? string.Empty).Trim();
                            string getTOStTypeTOStBinMaterials = $"({getTOStType})({getTOStBin}){getMaterials}";

                            string record = string.Join("\t",
                                getOrder.ToString(CultureInfo.InvariantCulture),
                                getPONumber,
                                getModel,
                                getMaterials,
                                getMaterialsDesc,
                                // write numeric with invariant culture
                                getQTy.ToString(CultureInfo.InvariantCulture),
                                getUOM,
                                getMaterialsClass,
                                getPriority,
                                getTONumber,
                                getTOStType,
                                getTOStBin,
                                getTOStTypeTOStBinMaterials
                            );

                            sw.WriteLine(record);
                        }
                    }

                    // Process the four grids in the same order as VB
                    ProcessGrid(dgvCommonPartPO);
                    ProcessGrid(dgvPartFirstPO);
                    ProcessGrid(dgvPartRestOfPO);
                    ProcessGrid(dgvPartPCBAOfPO);

                    // Write terminator line (mirrors VB terminating "End" fields)
                    // VB wrote many "End" separated by tabs; replicate a long sequence to be safe.
                    string terminator = string.Join("\t", new string[]
                    {
                    "End","End","End","End","End","End","End","End","End","End","End","End","End","End","End"
                    });
                    sw.Write(terminator);
                }
            }
            catch (Exception ex)
            {
                // VB was silent on errors; log for diagnostics
                MessageBox.Show($"OverrideWritePulledListToLocal failed: {ex}", "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Optionally rethrow depending on your app's error policy:
                // throw;
            }
        }

        /// <summary>
        /// VB6: Public Sub RecordPulledListPreAssyModelvsSector(getPulledListSector As String)
        /// </summary>
        /// <param name="getPulledListSector"></param>
        public void RecordPulledListPreAssyModelvsSector(string getPulledListSector)
        {
            // Loop through the rows of the grid (replace with your actual control)
            for (int ii = 0; ii <= dgvPulledListPO.Rows.Count; ii++)
            {
                string getModelSide = GetGridCellAsString(dgvPulledListPO, ii, 3);

                // Check if the record already exists
                if (!IsPulledListPreAssyModelvsSector(GetGridCellAsString(dgvPulledListPO, ii, 2), getModelSide, getPulledListSector))
                {
                    string pulledListPO = GetGridCellAsString(dgvPulledListPO, ii, 2);
                    string pulledListDateTime = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    string strQuery = $@"
                    INSERT INTO PulledListPreAssyModelvsSectorLog
                        (PulledListModel, PulledListSide, PulledListSector, PulledListDateTime) 
                    VALUES 
                        ('{pulledListPO}', '{getModelSide}', '{getPulledListSector}', '{pulledListDateTime}')";

                    // Execute the query
                    MSSQL _sql = new MSSQL();
                    SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

                    using (cnnDLVNDB)
                    {
                        cnnDLVNDB.Open();
                        using (SqlCommand cmd = new SqlCommand(strQuery, cnnDLVNDB))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// VB6: Public Function IsPulledListPreAssyModelvsSector(getPreAssyModel As String, getModelSide As String, getSector As String) As Boolean;
        /// </summary>
        /// <param name="getPreAssyModel"></param>
        /// <param name="getModelSide"></param>
        /// <param name="getSector"></param>
        /// <returns></returns>
        public bool IsPulledListPreAssyModelvsSector(string getPreAssyModel, string getModelSide, string getSector)
        {
            bool result = false;

            string query = @"
            SELECT TOP 1 * 
            FROM PulledListPreAssyModelvsSectorLog
            WHERE PulledListSide = @PulledListSide
            AND PulledListModel = @PulledListModel
            AND PulledListSector = @PulledListSector
            AND DATEDIFF(DAY, PulledListDateTime, GETDATE()) = 0";

            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            using (cnnDLVNDB)
            {
                using (SqlCommand command = new SqlCommand(query, cnnDLVNDB))
                {
                    // Add parameters to prevent SQL injection
                    command.Parameters.AddWithValue("@PulledListSide", getModelSide);
                    command.Parameters.AddWithValue("@PulledListModel", getPreAssyModel);
                    command.Parameters.AddWithValue("@PulledListSector", getSector);

                    cnnDLVNDB.Open();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        // Check if at least one row is returned
                        if (reader.HasRows)
                        {
                            result = true;
                        }
                    }
                }
            }

            return result;
        }

        private void ProcessPulledListFromLocal()
        {
            progressBarForm.UpdateProgress(88, "Exporting Pulled List from Local...");

            // Export Pulled List from Local:
            string strSector = cbbPulledListLine.Text.Trim();
            string strShift = cbbPulledListShift.Text.Trim();
            string strSumPOValID = Convert.ToString(_sumPOValue);
            if (_pulledList_Type == "DividePOByPOnLoc")
            {
                ExportFIAPulledListFromLocalByLocation(strSector, strShift, strSumPOValID);
            }
            else
            {
                ExportPulledListFromLocal(strSector, strShift, strSumPOValID);
            }
        }

        private void ExportFIAPulledListFromLocalByLocation(string sector, string shift, string sumPOValID)
        {

        }

        private void ExportPulledListFromLocal(string sector, string shift, string sumPOValID)
        {

        }

        /// <summary>
        /// This is part of cmdExport_Click in VB6;
        ///  Populate dgvPulledListPO from dgvMultiUniPhysicalModelPulled:
        ///  
        /// For new pulled lists not already in local folder.
        /// </summary>
        private void ProcessNewPulledList()
        {
            int intFullPOGroup;//VB6: FullPOGroup
            int intRow;//VB6: getRow
            string strKittingActive;
            string strModel;
            string strMaterial;
            string strMaterialDesc;
            string strTopBot;
            string models_missing_layout_setup = "";

            // AccessPastList will populate dgv
            progressBarForm.UpdateProgress(21, "Creating new Pulled List...");
            progressBarForm.UpdateProgress(22, "AccessPastListPO first time...");
            _repeatedProcess = 0;

            if (!AccessPastListPO())
            {
                // Enable combo boxes and button:
                cbbPulledListLine.Enabled = true;
                cbbActiveDate.Enabled = true;
                cbbPulledListShift.Enabled = true;
                RC_AddPOsKitting.Enabled = true;
                RC_Add1PO.Enabled = true;
                btnExport.Enabled = true;
                progressBarForm.Close();
                return;
            }

            //progressBarForm.UpdateProgress(44, "Grouping The Common Materials vs Production Order In PulledList Plan....");
            //int j = 0;
            //for (int i = 0; i < dgvMultiUniPhysicalModelPulled.Rows.Count; i++)
            //{
            //    var common_part_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 8);
            //    var model_in_multi = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 2);
            //    var material_in_multi = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 3);
            //    var partdesc_in_multi = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 4);
            //    var qty_in_multi = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 5);
            //    var uom_in_multi = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 6);
            //    var topbot_in_multi = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 9);

            //    if (!String.IsNullOrEmpty(common_part_in_multi_uni)) //CommonPart column not empty
            //    {
            //        //'flxMultiUniPhysicalModelPulled.FormatString =    "^No.|^PONumber|^Model|^Part Name|^Part Description|^Qty|^UOM|^PCBABin|^CommonPart|^TopBot"
            //        //'flxPullListvsPO.FormatString =                   "^No.|^PONumber|^Model|^Part Name|^Part Description|^Qty|^UOM|^NextPOUsed|^Materials Class|^TopBot"

            //        // Populate dgvPullListvsPO from dgvMultiUniPhysicalModelPulled:
            //        j++;

            //        SetGridCell(dgvPullListvsPO, i, 0, j); //No
            //        SetGridCell(dgvPullListvsPO, i, 1, common_part_in_multi_uni); //PONumber - CommonPart
            //        SetGridCell(dgvPullListvsPO, i, 2, model_in_multi); //ModelNumber
            //        SetGridCell(dgvPullListvsPO, i, 3, material_in_multi); //PartName
            //        SetGridCell(dgvPullListvsPO, i, 4, partdesc_in_multi); //PartDescription
            //        SetGridCell(dgvPullListvsPO, i, 5, qty_in_multi); //Qty
            //        SetGridCell(dgvPullListvsPO, i, 6, uom_in_multi); //UOM
            //        SetGridCell(dgvPullListvsPO, i, 7, ""); //^NextPOUsed
            //        SetGridCell(dgvPullListvsPO, i, 8, ""); //^Materials Class
            //        SetGridCell(dgvPullListvsPO, i, 9, topbot_in_multi); //TopBot

            //        intFullPOGroup = Convert.ToInt32(common_part_in_multi_uni.Substring(0, 1)); //ex: Get "1" from "1PO - (SMT102150988A_E)"
            //        if (intFullPOGroup == dgvPulledListPO.Rows.Count) //All added POs use this material (common part for all)
            //        {
            //            // Populate dgvPullListvsPO2 from dgvMultiUniPhysicalModelPulled if all POs use it:
            //            intRow = AddRowIfNeeded(dgvPullListvsPO2);
            //            SetGridCell(dgvPullListvsPO2, i, 0, j); //No
            //            SetGridCell(dgvPullListvsPO2, i, 1, common_part_in_multi_uni); //PONumber - CommonPart
            //            SetGridCell(dgvPullListvsPO2, i, 2, model_in_multi); //ModelNumber
            //            SetGridCell(dgvPullListvsPO2, i, 3, material_in_multi); //PartName
            //            SetGridCell(dgvPullListvsPO2, i, 4, partdesc_in_multi); //PartDescription
            //            SetGridCell(dgvPullListvsPO2, i, 5, qty_in_multi); //Qty
            //            SetGridCell(dgvPullListvsPO2, i, 6, uom_in_multi); //UOM
            //            SetGridCell(dgvPullListvsPO2, i, 7, ""); //^NextPOUsed
            //            SetGridCell(dgvPullListvsPO2, i, 8, ""); //^Materials Class
            //            SetGridCell(dgvPullListvsPO2, i, 9, topbot_in_multi); //TopBot
            //        }
            //    }
            //}
            // Create Pulled List in Local Folder:
            PhysicalBOMLog(dgvPullListvsPO, "020_MultiUniPhysicalModelPulled_To_flxPullListvsPO");
            PhysicalBOMLog(dgvPullListvsPO2, "021_MultiUniPhysicalModelPulled_To_flxPullListvsPO");

            //// Access Past Pulled List Production Orders:
            //progressBarForm.UpdateProgress(25, "AccessPastListPO again...");
            //dgvPullListvsPO.Rows.Clear();
            ////_repeatedProcess = 1;
            //if (!AccessPastListPO())
            //{
            //    // Enable combo boxes and button:
            //    cbbPulledListLine.Enabled = true;
            //    cbbActiveDate.Enabled = true;
            //    cbbPulledListShift.Enabled = true;
            //    RC_AddPOsKitting.Enabled = true;
            //    RC_Add1PO.Enabled = true;
            //    btnExport.Enabled = true;
            //    progressBarForm.Close();
            //    return;
            //}

            progressBarForm.UpdateProgress(30, "Processesing data tables...");
            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                for (int j = 0; j < dgvMultiUniPhysicalModelPulled.Rows.Count; j++)
                {
                    var common_part_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 8); //GetGridCellAsString(dgvMultiUniPhysicalModelPulled, jj, 1)
                    var po_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 1);
                    var po_in_pulled_list = GetGridCellAsString(dgvPulledListPO, i, 1);
                    var qty_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 5);
                    var model_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 2);
                    var material_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 3);
                    var partdesc_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 4);
                    var uom_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 6);
                    var topbot_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 9);

                    // Populate dgvPullListvsPO from dgvMultiUniPhysicalModelPulled:
                    int index_pulllistvspo = AddRowIfNeeded(dgvPullListvsPO);
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 0, index_pulllistvspo + 1); //No
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 1, common_part_in_multi_uni); //PONumber - CommonPart
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 2, model_in_multi_uni); //ModelNumber
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 3, material_in_multi_uni); //PartName
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 4, partdesc_in_multi_uni); //PartDescription
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 5, qty_in_multi_uni); //Qty
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 6, uom_in_multi_uni); //UOM
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 7, ""); //^NextPOUsed
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 8, ""); //^Materials Class
                    SetGridCell(dgvPullListvsPO, index_pulllistvspo, 9, topbot_in_multi_uni); //TopBot
                    
                    try
                    {
                        if (common_part_in_multi_uni.Length > 1)
                        {
                            common_part_in_multi_uni = common_part_in_multi_uni.Trim().Substring(0, 1);
                        }
                        else
                        {
                            common_part_in_multi_uni = "0"; // Invalid value to skip
                        }

                        intFullPOGroup = Convert.ToInt32(common_part_in_multi_uni); //first row may be "*xXx*" //PONumber
                    }
                    catch
                    {
                        intFullPOGroup = 0; // Invalid value to skip
                    }

                    int index_pulllistvspo2 = AddRowIfNeeded(dgvPullListvsPO2);
                    if (intFullPOGroup != dgvPulledListPO.Rows.Count)
                    {
                        if (po_in_multi_uni == po_in_pulled_list)
                        {
                            
                            if (Convert.ToInt32(qty_in_multi_uni) > 0)
                            {
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 0, index_pulllistvspo2 + 1); //No
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 1, po_in_multi_uni); //PONumber
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 2, model_in_multi_uni); //ModelNumber
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 3, material_in_multi_uni); //PartName
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 4, partdesc_in_multi_uni); //PartDescription
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 5, qty_in_multi_uni); //Qty
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 6, uom_in_multi_uni); //UOM
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 7, ""); //^NextPOUsed
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 8, ""); //^Materials Class
                                SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 9, topbot_in_multi_uni); //TopBot
                            }
                        }
                    }
                    else
                    {
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 0, index_pulllistvspo2 + 1); //No
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 1, common_part_in_multi_uni); //PONumber - CommonPart
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 2, model_in_multi_uni); //ModelNumber
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 3, material_in_multi_uni); //PartName
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 4, partdesc_in_multi_uni); //PartDescription
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 5, qty_in_multi_uni); //Qty
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 6, uom_in_multi_uni); //UOM
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 7, ""); //^NextPOUsed
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 8, ""); //^Materials Class
                        SetGridCell(dgvPullListvsPO2, index_pulllistvspo2, 9, topbot_in_multi_uni); //TopBot
                    }
                }
            }

            // Create Pulled List in Local Folder:
            PhysicalBOMLog(dgvPullListvsPO2, "130_EachPulledPO_FullFillTo_flxPullListvsPO2");

            //
            progressBarForm.UpdateProgress(63, "AddLocalPhantomSubPO..."); 
            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                var po_in_pulled_list = GetGridCellAsString(dgvPulledListPO, i, 1);
                var model_in_pulled_list = GetGridCellAsString(dgvPulledListPO, i, 2);
                var qty_in_pulled_list = Convert.ToDouble(GetGridCellAsString(dgvPulledListPO, i, 4));

                AddLocalPhantomSubPO(po_in_pulled_list, model_in_pulled_list, qty_in_pulled_list);
                //if (GetSectorGroup(cbbPulledListLine.Text.Trim()) != "BASE")
                //{
                //    // This will add a row with OUTER PACKAGING MATERIALS ...
                //    AddOverallPackaging(po_in_pulled_list, model_in_pulled_list);
                //}
            }

            // Clear grids and prepare to refill:
            progressBarForm.UpdateProgress(69, "Processing...");
            dgvPullListvsPO.Rows.Clear();
            dgvMultiUniPhysicalModelPulled.Rows.Clear();

            strKittingActive = IsKittingActive(cbbPulledListLine.Text.Trim()).ToUpper();

            progressBarForm.UpdateProgress(75, "Verify the materials setup...");
            for (int ii = 0; ii < dgvPullListvsPO2.Rows.Count; ii++)
            {
                var po_in_pulllistvspo2 = GetGridCellAsString(dgvPullListvsPO2, ii, 1).Trim();

                if (!String.IsNullOrEmpty(po_in_pulllistvspo2))
                {
                    intRow = AddRowIfNeeded(dgvPullListvsPO);
                    var model_in_pulllistvspo2 = GetGridCellAsString(dgvPullListvsPO2, ii, 2).Trim();
                    var material_in_pulllistvspo2 = GetGridCellAsString(dgvPullListvsPO2, ii, 3).Trim();
                    var partdesc_in_pulllistvspo2 = GetGridCellAsString(dgvPullListvsPO2, ii, 4).Trim();
                    var qty_in_pulllistvspo2 = GetGridCellAsString(dgvPullListvsPO2, ii, 5).Trim();
                    var uom_in_pulllistvspo2 = GetGridCellAsString(dgvPullListvsPO2, ii, 6).Trim();
                    var topbot_in_pulllistvspo2 = GetGridCellAsString(dgvPullListvsPO2, ii, 9).Trim();

                    // Populate dgvPullListvsPO from dgvPullListvsPO2 (stores materials used for all POs):
                    SetGridCell(dgvPullListvsPO, intRow, 0, intRow);
                    SetGridCell(dgvPullListvsPO, intRow, 1, po_in_pulllistvspo2);
                    SetGridCell(dgvPullListvsPO, intRow, 2, model_in_pulllistvspo2);
                    SetGridCell(dgvPullListvsPO, intRow, 3, material_in_pulllistvspo2);
                    SetGridCell(dgvPullListvsPO, intRow, 4, partdesc_in_pulllistvspo2);
                    SetGridCell(dgvPullListvsPO, intRow, 5, qty_in_pulllistvspo2);
                    SetGridCell(dgvPullListvsPO, intRow, 6, uom_in_pulllistvspo2);
                    SetGridCell(dgvPullListvsPO, intRow, 7, "");
                    SetGridCell(dgvPullListvsPO, intRow, 8, "");
                    SetGridCell(dgvPullListvsPO, intRow, 9, topbot_in_pulllistvspo2);

                    // Populate dgvMultiUniPhysicalModelPulled from dgvPullListvsPO2 (stores materials used for all POs):
                    intRow = AddRowIfNeeded(dgvMultiUniPhysicalModelPulled);
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 0, intRow);
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 1, po_in_pulllistvspo2);
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 2, model_in_pulllistvspo2);
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 3, material_in_pulllistvspo2);
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 4, partdesc_in_pulllistvspo2);
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 5, qty_in_pulllistvspo2);
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 6, uom_in_pulllistvspo2);
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 7, "");
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 8, "");
                    SetGridCell(dgvMultiUniPhysicalModelPulled, intRow, 9, topbot_in_pulllistvspo2);

                    strModel = GetGridCellAsString(dgvPullListvsPO, ii, 2);
                    strMaterial = GetGridCellAsString(dgvPullListvsPO, ii, 3);
                    strMaterialDesc = GetGridCellAsString(dgvPullListvsPO, ii, 4);

                    if (strMaterial.Length > 1)
                    {
                        //Trim blank space char
                        strMaterial = strMaterial.Replace("\u00A0", string.Empty).Trim();

                        if (_pulledList_Type == "DividePOByPOnLoc")
                        {
                            strTopBot = GetGridCellAsString(dgvPullListvsPO, ii, 9).Trim();
                            if (strMaterial.Substring(0, 3) != "100")
                            {
                                if (IsSetupMissingIn_PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix(strModel, strTopBot, strMaterial, strMaterialDesc, strKittingActive))
                                {
                                    models_missing_layout_setup += strModel + "/t" + strMaterial + "/t" + strMaterialDesc;
                                    MessageBox.Show("Model: " + strModel + " with Material: " + strMaterial + " - " + strMaterialDesc + " is missing setup in PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix!", "Layout Setup Missing Detected!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                    ////Function to send mail alert can be placed here
                                    //recordMissingSetUpWarning = recordMissingSetUpWarning & vbNewLine & strMissingSetup
                                    //Call ReleaseMissingDestinationSetup(frmMPHII.lstPulledListLine.Text, recordMissingSetUpWarning)

                                    //Exit the process if any setup missing found:
                                    //return;
                                }
                            }
                        }
                        else
                        {
                            if (IsSetupMissingIn_ModelvsPhysicalMaterials(strModel, strMaterial, strMaterialDesc, strKittingActive))
                            //Public Function IsMissingModelMaterialsLayoutMatrixSetup(getModel As String, getMaterials As String, getMaterialsDesc As String, getActiveKitting As String) As Boolean

                            {
                                models_missing_layout_setup += strModel + "/t" + strMaterial + "/t" + strMaterialDesc;
                                MessageBox.Show("Model: " + strModel + " with Material: " + strMaterial + " - " + strMaterialDesc + " is missing setup in PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix!", "Layout Setup Missing Detected!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                ////Function to send mail alert can be placed here
                                //recordMissingSetUpWarning = recordMissingSetUpWarning & vbNewLine & strMissingSetup
                                //Call ReleaseMissingDestinationSetup(frmMPHII.lstPulledListLine.Text, recordMissingSetUpWarning)

                                //Exit the process if any setup missing found:
                                //return;
                            }
                        }
                    }
                }
            }

            PhysicalBOMLog(dgvPullListvsPO, "140_flxPullListvsPO2_To_flxPullListvsPO");
            PhysicalBOMLog(dgvMultiUniPhysicalModelPulled, "141_flxPullListvsPO2_To_flxMultiUniPhysicalModelPulled");

            progressBarForm.UpdateProgress(81, "...");
            if ((!_blnSMTMaterialsMissingSetup && _pulledList_Type == "DividePOByPOnLoc") || _pulledList_Type != "DividePOByPOnLoc")
            {
                if (_pulledList_Type == "DividePOByPOnLoc")
                {
                    ExportPRAPulledListFromServerByLocation();
                }
                else //PARaw_Line is "DividePOByPO"
                {
                    ExportPulledListFromServer();
                }
            }

            progressBarForm.UpdateProgress(88, "...");
            if ((_pulledList_Type != "DividePOByPOnLoc") && (_pulledList_Type != "DividePOByPO"))
            {
                Fill_dgvPullListvsPO_withMaterialsClass();
                WriteNewPulledListToLocal(cbbPulledListLine.Text, cbbPulledListShift.Text, Convert.ToString(_sumPOValue));
            }
        }

        /// <summary>
        /// VB6: Public Sub NewWritePulledListToLocal(getSector As String, getShift As String, getSumPOValID As String);
        /// </summary>
        /// <param name="sector"></param>
        /// <param name="shift"></param>
        /// <param name="sumPOValID"></param>
        private void WriteNewPulledListToLocal(string sector, string shift, string sumPOValID)
        {
            // Generate file names
            string getFileNameWithoutPackagingMaterials = $"PulledList_{sector.Substring(0, sector.Length - 5)}_{shift.Substring(6, 1)}_{sumPOValID}_WithoutPK";
            string getFileNameWithPackagingMaterials = $"PulledList_{sector.Substring(0, sector.Length - 5)}_{shift.Substring(6, 1)}_{sumPOValID}_WithPK";

            // Generate file paths
            string targetFileWithoutPackagingMaterials = System.IO.Path.Combine(@"C:\MPH - KANBAN Control Local Data\PulledListLog\", $"{getFileNameWithoutPackagingMaterials}.txt");
            string targetFileWithPackagingMaterials = System.IO.Path.Combine(@"C:\MPH - KANBAN Control Local Data\PulledListLog\", $"{getFileNameWithPackagingMaterials}.txt");

            // Write to PulledList without Packaging Materials
            WritePulledListWithoutPackagingMaterials(targetFileWithoutPackagingMaterials);

            // Write to PulledList with Packaging Materials
            WritePulledListWithPackagingMaterials(targetFileWithPackagingMaterials);
        }

        private void WritePulledListWithoutPackagingMaterials(string targetFile)
        {
            StringBuilder recordSetUpData = new StringBuilder();
            recordSetUpData.AppendLine("No\tPO Number\tModel\tMaterials\tMaterial Description\tQuantity\tUOM\tMaterials Class\tPriority\tTONumber\tTOSt.Type\tTOSt.Bin\tTOSt.TypeTOSt.BinMaterials");

            int mm = 0;

            // Sample data block (replace these with your form data logic)
            // Replace this section with form row iteration
            for (int j = 1; j <= 10; j++) // Example loop
            {
                mm++;
                // Populate fields with dummy/test data
                int getOrder = mm;
                string getPONumber = $"PO_{j}";
                string getModel = $"Model_{j}";
                string getMaterials = $"Material_{j}";
                string getMaterialsDesc = "NA";
                double getQty = 10 * j;
                string getUOM = "NA";
                string getMaterialsType = "NA";
                string getPriority = "NA";
                string getTONumber = "NA";
                string getTOStType = "NA";
                string getTOStBin = "NA";
                string getTOStTypeTOStBinMaterials = $"Bin_Material_{j}";

                recordSetUpData.AppendLine($"{getOrder}\t{getPONumber}\t{getModel}\t{getMaterials}\t{getMaterialsDesc}\t{getQty}\t{getUOM}\t{getMaterialsType}\t{getPriority}\t{getTONumber}\t{getTOStType}\t{getTOStBin}\t{getTOStTypeTOStBinMaterials}");
            }

            recordSetUpData.AppendLine("End\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd");

            File.WriteAllText(targetFile, recordSetUpData.ToString());
        }

        private void WritePulledListWithPackagingMaterials(string targetFile)
        {
            StringBuilder recordSetUpData = new StringBuilder();
            recordSetUpData.AppendLine("No\tPO Number\tModel\tMaterials\tMaterial Description\tQuantity\tUOM\tMaterials Class");

            int mm = 0;

            // Sample data block (replace these with your form data logic)
            // Replace this section with form row iteration
            for (int j = 1; j <= 10; j++) // Example loop
            {
                mm++;
                // Populate fields with dummy/test data
                int getOrder = mm;
                string getPONumber = $"PO_{j}";
                string getModel = $"Model_{j}";
                string getMaterials = $"Material_{j}";
                string getMaterialsDesc = $"Description_{j}";
                double getQty = 10 * j;
                string getUOM = "UOM";
                string getMaterialsType = "Class";

                recordSetUpData.AppendLine($"{getOrder}\t{getPONumber}\t{getModel}\t{getMaterials}\t{getMaterialsDesc}\t{getQty}\t{getUOM}\t{getMaterialsType}");
            }

            recordSetUpData.AppendLine("End\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd");

            File.WriteAllText(targetFile, recordSetUpData.ToString());
        }

        /// <summary>
        /// Public Sub ExportPulledListFromServer()
        /// </summary>
        private void ExportPulledListFromServer()
        {
            string kittingActive = string.Empty;
            bool isPCBModel = false;
            bool isPCBAModel = false;
            bool isPCBComponent = false;
            bool isPCBAComponent = false;
            int fullPOGroup = 0;

            //dgvPullListvsPO: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); NextPOUsed(7); MatClass(8); TopBot(9)

            // Sort dgvUniPhysicalModelPulled based on column 3 (Materials), ascending order:
            dgvUniPhysicalModelPulled.Sort(dgvUniPhysicalModelPulled.Columns[3], ListSortDirection.Ascending);

            // Sort dgvMultiUniPhysicalModelPulled based on column 3 (Materials), ascending order:
            dgvMultiUniPhysicalModelPulled.Sort(dgvMultiUniPhysicalModelPulled.Columns[3], ListSortDirection.Ascending);


            // Sort dgvPullListvsPO based on column 1 (PONumber), descending order:
            dgvPullListvsPO.Sort(dgvPullListvsPO.Columns[1], ListSortDirection.Ascending);

            // Reset the priority of dgvPullListvsPO:
            for (int i = 0; i < dgvPullListvsPO.Rows.Count; i++)
            {
                SetGridCell(dgvPullListvsPO, i, 0, i + 1); // No column
            }

            for (int i = 0; i < dgvMultiUniPhysicalModelPulled.Rows.Count; i++)
            {

            }

            dgvCrossPlanningDone.Sort(dgvCrossPlanningDone.Columns[3], ListSortDirection.Descending);

            kittingActive = IsKittingActive(cbbPulledListLine.Text.Trim()).ToUpper();
            
            dgvUniPOnQtyMaterials.Rows.Clear();
            dgvKANBANPulledList.Rows.Clear();
            dgvMaterialsPONumber.Rows.Clear();
            dgvPartFirstPO.Rows.Clear();
            dgvCommonPartPO.Rows.Clear();
            dgvPartPCBAOfPO.Rows.Clear();
            dgvPartRestOfPO.Rows.Clear();

            int m = 1;
            int mm = 0;
            for (int i = 0; i < dgvMultiUniPhysicalModelPulled.Rows.Count; i++)
            {
                //dgvMultiUniPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDescription(4); Qty(5); UOM(6); PCBABin(7); CommonPart(8); TopBot(9)
                var poNumber_col_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 1).Trim();
                var model_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 2).Trim();
                var material_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 3).Trim();
                var partdesc_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 4).Trim();
                var qty_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 5).Trim();
                var uom_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 6).Trim();

                if (!string.IsNullOrEmpty(material_in_multi_uni))
                {
                    isPCBModel = IsPCBModel(material_in_multi_uni);
                    isPCBAModel = IsPCBAModel(material_in_multi_uni);
                    isPCBComponent = IsPCBComponent(material_in_multi_uni);
                    isPCBAComponent = IsPCBAComponent(material_in_multi_uni);

                    if ((_pulledList_SectorGroup == "POSTASSY_Group" && !isPCBComponent && isPCBAComponent && !isPCBModel && !isPCBAModel) ||
                        (_pulledList_SectorGroup == "SMT_Group" && isPCBComponent && !isPCBAComponent && !isPCBAModel) ||
                        (_pulledList_SectorGroup == "SMT_Group" && isPCBComponent && !isPCBAComponent && isPCBAModel) ||
                        (_pulledList_SectorGroup != "SMT_Group" && _pulledList_SectorGroup == "POSTASSY_Group"))
                    {
                        var pcbaBin = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 7).Trim();
                        if (string.IsNullOrEmpty(pcbaBin) && kittingActive == "YES")
                        {
                            m++;

                            if (poNumber_col_in_multi_uni.Length > 1)
                            {
                                fullPOGroup = Convert.ToInt32(poNumber_col_in_multi_uni.Substring(0, 1));
                            }
                            else
                            {
                                fullPOGroup = 0;
                            }

                            if (fullPOGroup == dgvPulledListPO.Rows.Count)
                            {
                                mm++;
                                // All added POs use this material (common part for all)
                                int rowIndex = AddRowIfNeeded(dgvCommonPartPO);
                                SetGridCell(dgvCommonPartPO, rowIndex, 0, mm); // No
                                SetGridCell(dgvCommonPartPO, rowIndex, 1, poNumber_col_in_multi_uni); // PONumber - CommonPart
                                SetGridCell(dgvCommonPartPO, rowIndex, 2, model_in_multi_uni); // ModelNumber
                                SetGridCell(dgvCommonPartPO, rowIndex, 3, material_in_multi_uni); // PartName
                                SetGridCell(dgvCommonPartPO, rowIndex, 4, partdesc_in_multi_uni); // PartDescription
                                SetGridCell(dgvCommonPartPO, rowIndex, 5, qty_in_multi_uni); // Qty
                                SetGridCell(dgvCommonPartPO, rowIndex, 6, uom_in_multi_uni); // UOM
                            }
                        }
                    }
                }
            }
            ExportPulledListFromServerToExcel();    
        }

        private void Setup_dgvCommonPartPO()
        {

        }

        private void Setup_dgvPartFirstPO()
        {

        }

        private void Setup_dgvPartPCBAOfPO()
        {

        }

        /// <summary>
        /// VB6: Public Sub ExportPulledListFromServerToExcel()
        /// </summary>
        private void ExportPulledListFromServerToExcel()
        {
            // Implement your Excel export logic here
        }

        /// <summary>
        /// VB6 equivalent: Public Sub CheckMaterialsClass()
        /// </summary>
        private void Fill_dgvPullListvsPO_withMaterialsClass() 
        {
            // lastGetQty mirrors the VB getQTy shown in the error MsgBox.
            double lastGetQty = 0.0;

            try
            {
                var grid = dgvPullListvsPO;
                if (grid == null) return;

                // Use rowCount to mimic VB's frmMPHII.flxPullListvsPO.Rows
                // Important: VB code loops i = 1 To (Rows - 2). VB is 1-based, so we convert.
                int vbRows = grid.RowCount; // replace with correct value if needed

                for (int i = 1; i <= vbRows - 2; i++)
                {
                    int r = i - 1; // convert VB 1-based row index to C# 0-based

                    // Check column 5 in VB -> index 4 in DataGridView
                    var cell5 = grid.Rows[r].Cells.Count > 4 ? grid.Rows[r].Cells[4].Value : null;
                    if (cell5 != null && !string.IsNullOrWhiteSpace(cell5.ToString()))
                    {
                        string partCurrentPO = grid.Rows[r].Cells.Count > 0 ? grid.Rows[r].Cells[0].Value?.ToString() : null;
                        string partNextPO = (r + 1 < grid.RowCount && grid.Rows[r + 1].Cells.Count > 0)
                            ? grid.Rows[r + 1].Cells[0].Value?.ToString()
                            : null;
                        string partPOModel = grid.Rows[r].Cells.Count > 1 ? grid.Rows[r].Cells[1].Value?.ToString() : null;
                        string partMaterials = grid.Rows[r].Cells.Count > 2 ? grid.Rows[r].Cells[2].Value?.ToString() : null;

                        // Parse quantity (VB used Double)
                        double getQty = 0.0;
                        Double.TryParse(cell5.ToString(), out getQty);
                        lastGetQty = getQty; // keep for error reporting like original VB MsgBox getQTy
                        double partQty = getQty;

                        string partCommonNextPO = grid.Rows[r].Cells.Count > 6 ? grid.Rows[r].Cells[6].Value?.ToString() : null;

                        // Call to your material-lookup routine (implement AccessMaterialsClass accordingly)
                        string MaterialType = GetMaterialClass(partMaterials?.Trim());

                        // Write result to column 8 in VB -> index 7 in DataGridView
                        if (grid.Rows[r].Cells.Count > 7)
                        {
                            grid.Rows[r].Cells[7].Value = MaterialType;
                        }
                        else
                        {
                            // If there aren't enough columns, optionally add or ignore.
                            // For simplicity we ignore; you can expand the grid if needed.
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Match original VB behavior which displayed getQTy in an error MsgBox.
                MessageBox.Show(lastGetQty.ToString(), "Error");
            }
        }

        private string IsKittingActive(string inputSector)
        {
            const string targetFile = @"C:\MPH - KANBAN Control Local Data\SectorGeneralInfor.txt";
            string result = "NA";

            if (string.IsNullOrWhiteSpace(inputSector))
                return result;

            if (!File.Exists(targetFile))
                return result;

            try
            {
                using (var sr = new StreamReader(targetFile))
                {
                    string strLine;
                    int rowStart = 0;
                    while ((strLine = sr.ReadLine()) != null)
                    {
                        var fields = strLine.Split('\t');

                        if (rowStart > 0 && !(fields.Length > 0 && string.Equals(fields[0].Trim(), "End", StringComparison.OrdinalIgnoreCase)))
                        {
                            string inFileSector = fields.Length > 0 ? fields[0].Trim() : string.Empty;
                            if (string.Equals(inputSector, inFileSector, StringComparison.OrdinalIgnoreCase))
                            {
                                if (fields.Length > 4)
                                    result = fields[4].Trim().ToUpperInvariant(); //YES or NO
                                else
                                    result = "NA";
                                break;
                            }
                        }

                        rowStart++;
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMessage = "Error in IsKittingActive: " + ex.Message;
                MessageBox.Show(errorMessage, "Exception Caught:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// VB6: With frmMPHII.flxPulledListPO - For ii = 1 To(.Rows - 2)
        /// </summary>
        /// <returns></returns>
        private double GetsumPOValID()
        {
            double sumPOVal = 1;

            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                //PO must be in the form of 102160897; not SMT102160897A_B
                string PONumber = Convert.ToString(dgvPulledListPO.Rows[i].Cells["PONumber"].Value);
                PONumber = PONumber.Substring(3, 9);
                sumPOVal += i * Convert.ToDouble(PONumber);
            }

            return sumPOVal;
        }

        /// <summary>
        /// Private Sub mAddPO_Click()
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RC_AddPOKitting_Click(object sender, EventArgs e) //VB6: mAddPO_Click():
        {
            // If Line chosen is SMT >> modify PO Number to match SMT format:
            if (cbbPulledListLine.Text.Trim().Length < 3)
            {
                MessageBox.Show("Please select a valid Line before adding PO Numbers!", "Line Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            PastePOsToGrid();
        }

        /// <summary>
        /// Private Sub mAddPO_Click()
        /// </summary>
        private void PastePOsToGrid()
        {
            string[] po_line_split; //VB6: splitedSubPO
            string[] po_list_table; //VB6: splittedPO
            string addedPO;
            string strTopBot;
            string pulled_po;
            string text;
            string models_missing_layout_setup = "None";

            // Get clipboard text:
            text = Clipboard.GetText() ?? string.Empty;

            // Split clipboard text into lines:
            po_list_table = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None); //VB6: splittedPO()

            if (po_list_table.Length > 1)
            {
                // Loop through each line except the last (which may be empty):
                for (int i = 0; i < po_list_table.Length - 1; i++)
                {
                    // use po_list_table[i]
                    if (po_list_table != null && i >= 0 && i < po_list_table.Length)
                    {
                        text = po_list_table[i]; //paste line to text
                    }
                    else
                    {
                        text = string.Empty;
                    }

                    // Split line by tab character: we copy 2 cells: 101829358	T
                    po_line_split = text.Split(new[] { "\t" }, StringSplitOptions.None);

                    // Get Added PO from :
                    addedPO = ""; //ex: 101829358
                    if (po_line_split.Length > 0)
                    {
                        addedPO = po_line_split[0].Replace("\u00A0", string.Empty).Trim();//Replace blank space char
                    }

                    // If copy with Top/Bot then strTopBot = copied value; else strTopBot = "A":
                    strTopBot = "A"; //All default TopBot = A
                    if (po_line_split.Length > 1) //at least 2 columns copied
                    {
                        if (!string.IsNullOrEmpty(po_line_split[1].Trim().ToUpperInvariant()))
                        {
                            strTopBot = po_line_split[1].Trim().ToUpperInvariant();
                        }
                    }

                    // Set pulled_po value initially:
                    pulled_po = addedPO + strTopBot;

                    var result = AddSinglePO(pulled_po);
                    if (result != "OK")
                    {
                        MessageBox.Show("Can't add PO: " + pulled_po + "\r\n" + result, "Add PO Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        //MessageBox.Show("PO added to list: " + pulled_po + "\r\n" + result, "Add PO:", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("No PO Numbers detected in clipboard!", "No PO Numbers", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Check if any components not setting up...
            if (_pulledList_SectorGroup == "SMT_Group")
            {
                models_missing_layout_setup = CheckSetup_ModelvsMaterialsLayoutMatrix();
                if (models_missing_layout_setup.ToUpper() != "NONE")
                {
                    btnExport.Enabled = false;
                    MessageBox.Show("The following Model(s) with Material(s) is/are missing setup in PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix: " + Environment.NewLine + models_missing_layout_setup, "Missing Setup Detected!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    btnExport.Enabled = true;
                }
            }
            dgvPulledListPO.Refresh();
        }

        private bool IsPOQtyChanged(string PO)
        {
            bool is_change_qty = false;
            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;
            string query = @"SELECT TOP 1 * FROM OpenPOPlanner WHERE TopPONumber = @ponumber AND POChangeInf = 'YES'";
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@poNumber", PO);
                        conn.Open();
                        var result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            is_change_qty = true;
                        }
                        else
                        {
                            is_change_qty = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error checking PO change quantity: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return is_change_qty;
        }

        /// <summary>
        /// VB6: Public Function GetFullPOQty(ByVal getPONumber As String) As Integer
        /// </summary>
        /// <param name="PO"></param>
        /// <returns></returns>
        private int GetPOQty(string PO)
        {
            int po_qty = 0;
            string convertedponumber = PO.Substring(3,9);//Ex: SMT102099984A_G >> get 102099984
            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;
            string query = @"SELECT TOP 1 * FROM OpenPOPlanner " +
                            "WHERE (TopPONumber = @ponumber OR TopPONumber = @convertedponumber) AND POQty > '0'";
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@poNumber", PO);
                        cmd.Parameters.AddWithValue("@convertedponumber", PO);
                        conn.Open();
                        var reader = cmd.ExecuteReader();
                        if (reader.HasRows)
                        {
                            reader.Read();
                            po_qty = Convert.ToInt32(reader["POQty"]);
                        }
                        else
                        {
                            po_qty = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error retrieving PO quantity: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return po_qty;
        }

        /// <summary>
        /// VB6: Public Function GetPOModel(getPONumber As String) As String
        /// </summary>
        /// <param name="PO"></param>
        /// <returns></returns>
        private string GetPOModelFromOpenPOPlanner(string PO)
        {
            string model = "NA";
            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;
            string query = @"SELECT TOP 1 * FROM OpenPOPlanner where TopPONumber = @ponumber";
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@poNumber", PO);
                        conn.Open();
                        var reader = cmd.ExecuteReader();
                        if (reader.HasRows)
                        {
                            reader.Read();
                            model = Convert.ToString(reader["TopModel"]);
                        }
                        else
                        {
                            model = "NA";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error retrieving PO model: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return model;
        }

        /// <summary>
        /// VB6: Public Function GetSectorPO(getPONumber As String) As String;
        /// </summary>
        /// <param name="PO"></param>
        /// <returns></returns>
        private string GetPOSector(string PO)
        {
            string sector = "NA";
            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;
            string query = @"SELECT TOP 1 * FROM OpenPOPlanner WHERE TopPONumber = @ponumber";
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@poNumber", PO);
                        conn.Open();
                        var reader = cmd.ExecuteReader();
                        if (reader.HasRows)
                        {
                            reader.Read();
                            sector = Convert.ToString(reader["Sector"]);
                        }
                        else
                        {
                            if (cbbPulledListLine.Text.Trim() == "PostAssyA_Line")
                            {
                                sector = "PostAssyA_Line";
                            }
                            else
                            {
                                sector = "NONE";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error retrieving PO sector: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sector;
        }

        private bool IsPOUnique(string PO)
        {
            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                string existingPO = dgvPulledListPO.Rows[i].Cells["PONumber"].Value.ToString().Trim();
                if (string.Equals(existingPO, PO, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("PO Number: " + PO + " is already in the list!", "Duplicate PO Number", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false; // PO is not unique
                }
            }
            // PO is unique
            return true;
        }

        /// <summary>
        /// VB6 equivalent: Public Function IsMissingModelandMaterialsLayoutMatrixPreAssySetup(ByVal getModel As String, ByVal getTopBot As String, ByVal getMaterials As String, ByVal getMaterialsDesc As String, ByVal getActiveKitting As String) As Boolean
        /// </summary>
        /// <returns></returns>
        private bool IsSetupMissingIn_PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix(string modelNumber, string topBot, string materialNumber, string materialDesc, string activeKitting)
        {
            bool result = false; // matches VB default behavior where function result is only set to False in code paths
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            // Normalize materials (remove non-breaking space char 160)
            if (!string.IsNullOrEmpty(materialNumber))
                materialNumber = materialNumber.Replace(((char)160).ToString(), "").Trim();//DLVNpn

            string getModelSeries = GetSeries(modelNumber);

            // getSector / getSMTLine come from UI
            string getSector = string.Empty;
            string getSMTLine = string.Empty;
            try
            {
                // frmMPHII.lstPulledListLine.Text in VB
                getSector = cbbPulledListLine.Text;
                getSMTLine = cbbPulledListLine.Text;
            }
            catch
            {
                // ignore UI errors here; proceed with empty values
            }

            // Early exits matching VB logic
            if (_pulledList_SectorGroup == "POSTASSY_Group")
                return false;

            if (_pulledList_SectorGroup == "SMT_Group" && !string.Equals(topBot, "A", StringComparison.OrdinalIgnoreCase))
                return false;

            // Compute SMTLine character equivalent to VB: Mid(getSMTLine, Len(getSMTLine) - 5, 1)
            string smtLineChar = string.Empty;
            if (!string.IsNullOrEmpty(getSMTLine) && getSMTLine.Length >= 6)
            {
                int index = getSMTLine.Length - 6; // VB 1-based -> C# 0-based
                if (index >= 0 && index < getSMTLine.Length)
                    smtLineChar = getSMTLine[index].ToString();
            }

            // Build and run SQL query (TOP 1)
            string strQuery = $"SELECT TOP 1 * FROM PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix " +
                              $"WHERE Model = @model AND DLVNpn = @material AND SMTLine = @smtLine";

            try
            {
                using (var cmd = new SqlCommand(strQuery, cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.AddWithValue("@model", modelNumber ?? string.Empty);
                    cmd.Parameters.AddWithValue("@material", materialNumber ?? string.Empty);
                    cmd.Parameters.AddWithValue("@smtLine", smtLineChar ?? string.Empty);

                    bool openedHere = false;
                    try
                    {
                        if (cnnDLVNDB == null)
                            throw new InvalidOperationException("Database connection (cnnDLVNDB) is not initialized.");

                        if (cnnDLVNDB.State != ConnectionState.Open)
                        {
                            cnnDLVNDB.Open();
                            openedHere = true;
                        }

                        using (var reader = cmd.ExecuteReader(CommandBehavior.SingleRow))
                        {
                            if (reader != null && reader.Read())
                            {
                                // Found setup record
                                result = false;
                            }
                            else
                            {
                                // No record found
                                if (string.Equals(activeKitting, "YES", StringComparison.OrdinalIgnoreCase))
                                {
                                    string strMaterialAfterProgramming = GetMaterialAfterProgramming(materialNumber, modelNumber);

                                    if (!string.Equals(strMaterialAfterProgramming, "NA", StringComparison.OrdinalIgnoreCase))
                                    {
                                        result = false;
                                        _blnSMTMaterialsMissingSetup = false;
                                    }
                                    else
                                    {
                                        // Mark missing setup and notify user as in VB
                                        _blnSMTMaterialsMissingSetup = true;
                                        bool isPCBModel = IsPCBModel(materialNumber);
                                        string isPCBAModel = IsPCBAModel(materialNumber) ? "True" : "False";
                                        bool isPCBComponent = IsPCBComponent(materialNumber);
                                        bool isPCBAComponent = IsPCBAComponent(materialNumber);

                                        string msg =
                                            "Vui Long Yeu Cau Ky Su Phu Trach Set-up Sheet Tren May SMT De Set-up Linh Kien Sau:" + Environment.NewLine +
                                            "      - SMT PCB Panel MODEL: " + modelNumber + Environment.NewLine +
                                            "      - Materials Missing Destination: " + materialNumber + Environment.NewLine +
                                            "      - Materials Information: " + Environment.NewLine +
                                            "               + isPCBModel: " + isPCBModel + Environment.NewLine +
                                            "               + isPCBAModel: " + isPCBAModel + Environment.NewLine +
                                            "               + isPCBComponent: " + isPCBComponent + Environment.NewLine +
                                            "               + isPCBAComponent: " + isPCBAComponent + Environment.NewLine + Environment.NewLine +
                                            "Note: See Module **IsMissingModelandMaterialsLayoutMatrixPreAssySetup**" + Environment.NewLine + Environment.NewLine +
                                            "DATABASE Query: " + strQuery;

                                        MessageBox.Show(msg,
                                            "MISSING SET-UP SHEET IN DATABASE [DLVNDB].[dbo].[PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix]",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);

                                        // In original VB there was commented logic to allow certain users to bypass;
                                        // we preserve current behavior (SMTMaterialsMissingSetup = true). result remains false (VB did not set it true).
                                    }
                                }
                                else
                                {
                                    result = false;
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
                        {
                            try { cnnDLVNDB.Close(); } catch { /* swallow to mimic VB */ }
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("**IsMissingModelandMaterialsLayoutMatrixPreAssySetup** Module Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// VB6 equivalent: Public Function IsMissingModelMaterialsLayoutMatrixSetup(getModel As String, getMaterials As String, getMaterialsDesc As String, getActiveKitting As String) As Boolean
        /// This will search data from 
        /// </summary>
        /// <param name="modelNumber"></param>
        /// <param name="materialName"></param>
        /// <param name="materialDesc"></param>
        /// <param name="activeKitting"></param>
        /// <returns></returns>
        private bool IsSetupMissingIn_ModelvsPhysicalMaterials(string modelNumber, string materialNumber, string materialDesc, string activeKitting)
        {
            bool blnResult = false; // matches VB default behavior where function result is only set to False in code paths

            materialNumber = materialNumber.Replace(((char)160).ToString(), "").Trim();

            string modelSeries = GetSeries(modelNumber);
                
            string sector = cbbPulledListLine.Text;
            string getSMTLine = cbbPulledListLine.Text;


            // Early exits matching VB logic
            if (_pulledList_SectorGroup == "POSTASSY_Group")
            {
                blnResult = false;
                return blnResult;
            }

            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            // Build and run SQL query (TOP 1)
            string strQuery = @"SELECT TOP 1 * FROM " + _modelVsMaterialsLayoutDB + " " +
                              "WHERE Model = @model " +
                              "AND Materials = @material";

            try
            {
                using (var cmd = new SqlCommand(strQuery, cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.AddWithValue("@model", modelNumber ?? string.Empty);
                    cmd.Parameters.AddWithValue("@material", materialNumber ?? string.Empty);

                    bool openedHere = false;
                    try
                    {
                        if (cnnDLVNDB == null)
                            throw new InvalidOperationException("Database connection (cnnDLVNDB) is not initialized.");

                        if (cnnDLVNDB.State != ConnectionState.Open)
                        {
                            cnnDLVNDB.Open();
                            openedHere = true;
                        }

                        using (var reader = cmd.ExecuteReader(CommandBehavior.SingleRow))
                        {
                            if (reader != null && reader.Read())
                            {
                                // Found setup record
                                blnResult = false;
                            }
                            else
                            {
                                // No record found
                                if (string.Equals(activeKitting, "YES", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (_pulledList_SectorGroup == "SMT_Group")
                                    {
                                        blnResult = true;

                                        string msg =
                                            "Vui Long Yeu Cau Ky Su Phu Trach (Thang Heo Con) De Set-up Linh Kien Sau:" + Environment.NewLine +
                                            "      - SMT PCB Panel MODEL: " + modelNumber + Environment.NewLine +
                                            "      - Materials Missing Destination: " + materialNumber + Environment.NewLine +
                                            "      - ModelSeries:  " + modelSeries + Environment.NewLine +
                                            "      - Materials Missing Destination: " + materialNumber + Environment.NewLine + Environment.NewLine +
                                            "Note: See Module **IsSetupMissingIn_ModelvsPhysicalMaterials**" + Environment.NewLine + Environment.NewLine +
                                            "DATABASE Query: " + strQuery;

                                        MessageBox.Show(msg,
                                            "MISSING SET-UP SHEET IN DATABASE " + _modelVsMaterialsLayoutDB,
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    else
                                    {
                                        if (!IsMaterialAlreadyInSeriesLayout(modelNumber, sector, materialNumber))
                                        {
                                            if (IsMaterialIn_ModelvsMaterialsLayoutMatrixWarningNoSetup(modelNumber, materialNumber))
                                            {
                                                string insertCmd = @"INSERT INTO ModelvsMaterialsLayoutMatrixWarningNoSetup (Sector,Model,Materials,AddToKittingDateTime) " +
                                                                       "VALUES (@sector, @model, @material, @dateTime)";
                                                using (var insertCommand = new SqlCommand(insertCmd, cnnDLVNDB))
                                                {
                                                    insertCommand.CommandType = CommandType.Text;
                                                    insertCommand.Parameters.AddWithValue("@sector", sector);
                                                    insertCommand.Parameters.AddWithValue("@model", modelNumber);
                                                    insertCommand.Parameters.AddWithValue("@material", materialNumber);
                                                    insertCommand.Parameters.AddWithValue("@dateTime", System.DateTime.Now);
                                                    try
                                                    {
                                                        insertCommand.ExecuteNonQuery();
                                                        blnResult = true;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show("Error logging missing material setup: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                        blnResult = false;
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            blnResult = false;
                                        }
                                    }
                                }
                                else
                                {
                                    blnResult = false;
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
                        {
                            try { cnnDLVNDB.Close(); } catch { /* swallow to mimic VB */ }
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("**IsMissingModelandMaterialsLayoutMatrixPreAssySetup** Module Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return blnResult;
        }

        /// <summary>
        /// VB6: Public Function IsModelvsMaterialsLayoutNoSetup(getModel As String, getMaterials As String) As Boolean
        /// </summary>
        /// <param name="modelNumber"></param>
        /// <param name="materialNumber"></param>
        /// <returns></returns>
        private bool IsMaterialIn_ModelvsMaterialsLayoutMatrixWarningNoSetup(string modelNumber, string materialNumber)
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            // Emulate VB: getMaterials = Trim(Replace(getMaterials, Chr(60), ""))
            materialNumber = (materialNumber ?? string.Empty).Replace("<", "").Trim();
            modelNumber = modelNumber ?? string.Empty;

            const string sql = @"SELECT COUNT(1) FROM ModelvsMaterialsLayoutMatrixWarningNoSetup " +
                "WHERE Model = @Model " + 
                "AND Materials = @Materials";

            using (var cmd = new SqlCommand(sql, cnnDLVNDB))
            {
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@Model", modelNumber);
                cmd.Parameters.AddWithValue("@Materials", materialNumber);

                // ExecuteScalar is efficient for existence/count checks
                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value) return false;

                if (int.TryParse(result.ToString(), out int count))
                    return count > 0;

                // Fallback: non-integer result -> treat as false
                return false;
            }
        }

        /// <summary>
        /// VB6: Public Function IsMaterialsInSeriesLayoutExistingAndUpdate(getModel As String, getSector As String, getMaterials As String) As Boolean
        /// This will check if this material (ex.025001007) exists in database (ex. [DLVNDB].[dbo].[ModelvsPhysicalMaterials_SMTA_Line]) but for different Model (ex.3-0998-05)
        /// copy that to input model (ex. 9-9999-99)
        /// </summary>
        /// <returns></returns>
        private bool IsMaterialAlreadyInSeriesLayout(string modelNumber, string sector, string materialNumber)
        {
            bool blnResult = false;

            if (_pulledList_SectorGroup == "SMT_Group")
            {
                blnResult = false;
            }
            else
            {
                // Clean materials string: remove '<' (Chr(60)) and NBSP (Chr(160)) and trim
                if (materialNumber == null) materialNumber = string.Empty;
                materialNumber = materialNumber.Replace("<", string.Empty).Replace(((char)160).ToString(), string.Empty).Trim();

                // Build SELECT query (explicit column list to avoid reading all columns)
                string selectSql = @"SELECT TOP(1) FROM " + _modelVsMaterialsLayoutDB + " " +
                    "WHERE Sector = @sector " +
                    "AND Materials = @material";

                MSSQL _sql = new MSSQL();
                SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

                using (var selectCmd = new SqlCommand(selectSql, cnnDLVNDB))
                {
                    selectCmd.CommandType = CommandType.Text;
                    selectCmd.Parameters.AddWithValue("@sector", sector);
                    selectCmd.Parameters.AddWithValue("@material", materialNumber);

                    using (var reader = selectCmd.ExecuteReader())
                    {
                        if (!reader.HasRows)
                        {
                            // No records -> return false
                            blnResult = false;
                            return blnResult;
                        }
                        else
                        {
                            while (reader.Read())
                            {
                                string queriedSector = Convert.ToString(reader["Sector"]).Trim().ToUpper();
                                double queriedQtyperProduct = Convert.ToDouble(reader["QtyperProduct"]);
                                string queriedPSA = Convert.ToString(reader["PSA"]).Trim().ToUpper();
                                string queriedWorkstationCode = Convert.ToString(reader["WorkstationCode"]).Trim().ToUpper();
                                string queriedMaterialsLocOnWSCode = Convert.ToString(reader["MaterialsLocOnWSCode"]).Trim().ToUpper();
                                materialNumber = materialNumber?.Replace("\u00A0", "").Trim() ?? string.Empty; //remove space char

                                if (!IsModelvsMaterialsLayoutAlreadySetup(modelNumber, materialNumber, queriedWorkstationCode, queriedMaterialsLocOnWSCode, queriedSector))
                                {
                                    if (!IsModelvsMaterialsLayoutDup(modelNumber, queriedMaterialsLocOnWSCode))
                                    {
                                        // Build INSERT query to copy record to input model
                                        string insertSql = @"INSERT INTO " + _modelVsMaterialsLayoutDB + " " +
                                            "(Model, Sector, QtyperProduct, PSA, WorkstationCode, MaterialsLocOnWSCode, Materials, TopBotRunning) " +
                                            "VALUES (@model, @sector, @qtyperproduct, @psa, @workstationcode, @materialsloconwscode, @material, 'A')";

                                        using (var insertCmd = new SqlCommand(insertSql, cnnDLVNDB))
                                        {
                                            insertCmd.CommandType = CommandType.Text;
                                            insertCmd.Parameters.AddWithValue("@model", modelNumber);
                                            insertCmd.Parameters.AddWithValue("@sector", queriedSector);
                                            insertCmd.Parameters.AddWithValue("@qtyperproduct", queriedQtyperProduct);
                                            insertCmd.Parameters.AddWithValue("@psa", queriedPSA);
                                            insertCmd.Parameters.AddWithValue("@workstationcode", queriedWorkstationCode);
                                            insertCmd.Parameters.AddWithValue("@materialsloconwscode", queriedMaterialsLocOnWSCode);
                                            insertCmd.Parameters.AddWithValue("@material", materialNumber);

                                            try
                                            {
                                                insertCmd.ExecuteNonQuery();
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("Error copying material setup to model: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                blnResult = false;
                                            }
                                        }
                                    } 
                                }
                            }
                            blnResult = true; // Successfully copied
                        }    
                    }
                }
            }
            return blnResult;
        }

        /// <summary>
        /// VB6: Public Function IsModelvsMaterialsLayoutDup(getModel As String, getMaterialsOnWSCode As String) As Boolean
        /// </summary>
        /// <returns></returns>
        private bool IsModelvsMaterialsLayoutDup(string modelNumber, string materialsOnWSCode)
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            string sql = @"SELECT TOP(1) * FROM " + _modelVsMaterialsLayoutDB + " WHERE" +
                                    " Model = @modelNumber" +
                                    " AND MaterialsLocOnWSCode = @materialsOnWSCode";

            using (var selectCmd = new SqlCommand(sql, cnnDLVNDB))
            {
                selectCmd.CommandType = CommandType.Text;
                selectCmd.Parameters.AddWithValue("@modelNumber", modelNumber);
                selectCmd.Parameters.AddWithValue("@materialsOnWSCode", materialsOnWSCode);
                using (var reader = selectCmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        return true; // Record exists
                    }
                    else
                    {
                        return false; // No record
                    }
                }
            }
        }

        /// <summary>
        /// VbB: Public Function IsModelvsMaterialsLayoutAlreadySetup(getModel As String, getMaterials As String, getWSCode As String, getMaterialsOnWSCode As String, getSector As String) As Boolean
        /// </summary>
        /// <returns></returns>
        private bool IsModelvsMaterialsLayoutAlreadySetup(string modelNumber, string materialNumber, string wsCode, string materialOnWSCode, string sector)
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            string sql = @"SELECT TOP(1) * FROM " + _modelVsMaterialsLayoutDB + " WHERE" +
                                    " Model = @modelNumber" +
                                    " AND Materials = @materialNumber" +
                                    " AND WorkstationCode = @wsCode" +
                                    " AND MaterialsLocOnWSCode = @materialsOnWSCode" +
                                    " AND Sector = @sector";

            using (var selectCmd = new SqlCommand(sql, cnnDLVNDB))
            {
                selectCmd.CommandType = CommandType.Text;
                selectCmd.Parameters.AddWithValue("@modelNumber", modelNumber);
                selectCmd.Parameters.AddWithValue("@materialNumber", materialNumber);
                selectCmd.Parameters.AddWithValue("@wsCode", wsCode);
                selectCmd.Parameters.AddWithValue("@materialsOnWSCode", materialOnWSCode);
                selectCmd.Parameters.AddWithValue("@sector", sector);
                using (var reader = selectCmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        return true; // Record exists
                    }
                    else
                    {
                        return false; // No record
                    }
                }
            }
        }

        /// <summary>
        /// //VB6: Public Function PostAssyModelMissingMaterialsLayout() As String
        /// this will check all the models of added POs if they have been setup in PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix
        /// </summary>
        /// <returns></returns>
        private string CheckSetup_ModelvsMaterialsLayoutMatrix() 
        {
            string models_missing_layout_setup = "";
            string query;
            string po_model;
            string po_side;

            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;

            if (dgvPulledListPO.Rows.Count == 0)
            {
                return "No PO added";
            }
            
            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                //dgvPulledListPO: No(0); PONumber(1); ModelNumber(2); Side(3); POQty(4); PulledListID(5); PlannersNotice(6); POChangeInf(7)
                po_model = GetGridCellAsString(dgvPulledListPO, i, 2);
                po_side = GetGridCellAsString(dgvPulledListPO, i, 3);
                query = @"SELECT TOP (1) * FROM PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix " +
                        "WHERE Model = @model AND SMTLine = @smtline";
                if (po_side != "A")
                {
                    query += @" AND ModelSide = @side";
                }

                try
                {
                    using (var conn = new SqlConnection(connectionString))
                    {
                        using (var cmd = new SqlCommand(query, conn))
                        {
                            cmd.CommandType = CommandType.Text;
                            cmd.Parameters.AddWithValue("@model", po_model);
                            cmd.Parameters.AddWithValue("@smtline", cbbPulledListLine.Text.Trim().Substring(3,1));
                            if (po_side != "A")
                            {
                                cmd.Parameters.AddWithValue("@side", po_side);
                            }
                            conn.Open();
                            var exists = cmd.ExecuteScalar();
                            // If no record found (ExecuteScalar returns null), this model lacks setup
                            if (exists == null)
                            {
                                //return po_model;//this will break the loop and return the first missing model
                                models_missing_layout_setup += po_model + Environment.NewLine; //this will continue the loop and get all the model
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error verifying PO Change Over existence: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return "Error";
                }
            }    
            if (string.IsNullOrEmpty(models_missing_layout_setup))
            {
                return "NONE";//No models missing layout, good!
            }
            else
            {
                return models_missing_layout_setup.TrimEnd(); // Remove trailing newline
            }
        }

        private bool IsValidPO(string PO)
        {
            bool isValid = false;
            // Check length and format
            if (
                (PO.Substring(0, 3) == "100" && PO.Length == 10 && PO.Substring(8, 1) == "T") ||
                (PO.Substring(0, 3) == "101" && PO.Length == 10 && PO.Substring(8, 1) == "T") ||
                (PO.Substring(0, 3) == "102" && PO.Length == 10 && PO.Substring(8, 1) == "T") ||
                (PO.Substring(0, 3) == "103" && PO.Length == 10 && PO.Substring(8, 1) == "T") ||
                (PO.Substring(0, 3) == "100" && PO.Length == 10 && PO.Substring(8, 1) == "B") ||
                (PO.Substring(0, 3) == "101" && PO.Length == 10 && PO.Substring(8, 1) == "B") ||
                (PO.Substring(0, 3) == "102" && PO.Length == 10 && PO.Substring(8, 1) == "B") ||
                (PO.Substring(0, 3) == "103" && PO.Length == 10 && PO.Substring(8, 1) == "B") ||
                (PO.Substring(0, 3) == "900" && PO.Length == 10 && PO.Substring(8, 1) == "T") ||
                (PO.Substring(0, 3) == "900" && PO.Length == 10 && PO.Substring(8, 1) == "B") ||
                (PO.Substring(0, 3) == "100" && PO.Length == 9) ||
                (PO.Substring(0, 3) == "101" && PO.Length == 9) ||
                (PO.Substring(0, 3) == "102" && PO.Length == 9) ||
                (PO.Substring(0, 3) == "103" && PO.Length == 9) ||
                (PO.Substring(0, 3) == "800" && PO.Length == 9) ||
                (PO.Substring(0, 3) == "900" && PO.Length == 9) ||
                (PO.Substring(0, 3) == "500" && PO.Length == 9) ||// Rework POs
                (PO.Substring(0, 3) == "SMT" && PO.Length == 15) //SMT POs
                )
            {
                isValid = true;
            }
            return isValid;
        }

        /// <summary>
        /// Vb6: Public Function VerifyPOPairCOExist(ByVal getPOaftCO As String);
        /// Input PO must exist in POpairunderCO table as POaftCO
        /// </summary>
        /// <param name="po_aft_co"></param>
        private void EnsurePOExistIn_dboPOpairunderCO(string poNumber) 
        {
            bool blnExist = false;
            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;

            string query = "SELECT TOP(1) * FROM POpairunderCO WHERE POaftCO = @poaft";

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@poaft", poNumber);
                        conn.Open();
                        var result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            blnExist = true;
                        }
                    }
                }

                if (!blnExist)
                {
                    query = "INSERT INTO POpairunderCO (PObefCO,POaftCO,COTime,ActiveDateTime) VALUES (@po_bfr_co, @po_aft_co, '0', @active_datetime)";
                    string po_bfr_co = GetPObefCOFromPOpairunderCO(poNumber);
                    using (var conn = new SqlConnection(connectionString))
                    {
                        using (var cmd = new SqlCommand(query, conn))
                        {
                            cmd.CommandType = CommandType.Text;
                            cmd.Parameters.AddWithValue("@po_bfr_co", "NA");
                            cmd.Parameters.AddWithValue("@po_aft_co", poNumber);
                            cmd.Parameters.AddWithValue("@active_datetime", System.DateTime.Now);
                            conn.Open();
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error verifying PO Change Over existence: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        /// <summary>
        /// Vb6: Public Function PCBABin(getMaterials As String) As String
        /// </summary>
        /// <param name="partNumber"></param>
        /// <returns></returns>
        private string GetPCBABin(string partNumber)
        {
            if (string.IsNullOrWhiteSpace(partNumber))
                return string.Empty;

            const string path = @"C:\MPH - KANBAN Control Local Data\MPHKANBANFixLoc.txt";

            if (!File.Exists(path))
            {
                // In the original VB this called UpdateLocalPCBABin and exited the function.
                // Preserve that behavior by invoking the updater and returning empty.
                try
                {
                    Update_MPHKANBANFixLocTxt();
                }
                catch
                {
                    // Swallow exceptions to match VB behavior; consider logging in real code.
                }
                return string.Empty;
            }

            string result = string.Empty;

            try
            {
                using (var sr = new StreamReader(path))
                {
                    string line;
                    // VB code keeps a rowStart counter but doesn't use it for logic except incrementing.
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var fields = line.Split('\t');

                        // VB: getLocalPartName = Trim(splitedFields(2))
                        // VB: getLocalLocation = Trim(splitedFields(1))
                        string getLocalPartName = fields.Length > 2 ? fields[2].Trim() : string.Empty;
                        string getLocalLocation = fields.Length > 1 ? fields[1].Trim() : string.Empty;

                        if (string.Equals(getLocalPartName, partNumber, StringComparison.OrdinalIgnoreCase))
                        {
                            // VB sets PCBABin = getLocalLocation (overwrites, does not Exit For)
                            result = getLocalLocation;
                        }
                    }
                }
            }
            catch
            {
                // Match VB style: do not throw; return empty string on error.
                result = string.Empty;
            }

            return result;
        }

        /// <summary>
        /// Get PObefCO from POpairunderCO table based on POaftCO input
        /// </summary>
        /// <param name="po"></param>
        /// <returns></returns>
        private string GetPObefCOFromPOpairunderCO(string po)
        {
            string po_bfr_co = "NA";
            int intPOPriority = 0;

            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;
            string query = @"
                SELECT TOP(1) Priority
                FROM [DLVNDB].[dbo].[OpenPOPlanner]
                WHERE TopPONumber = @topPONumber";

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@topPONumber", po);
                        conn.Open();
                        var reader = cmd.ExecuteReader();
                        if (reader.HasRows)
                        {
                            reader.Read();
                            intPOPriority = Convert.ToInt32(reader["Priority"]);
                        }
                        else
                        {
                            intPOPriority = 0;
                        }
                    }
                }

                if (intPOPriority > 0)
                {
                    query = @"
                        SELECT TOP(1) TopPONumber
                        FROM [DLVNDB].[dbo].[OpenPOPlanner]
                        WHERE Priority = @priority
                          AND Sector = @sector
                          AND TypeCO = 'Next PO'
                        ORDER BY ActiveDateTime DESC";
                    using (var conn = new SqlConnection(connectionString))
                    {
                        using (var cmd = new SqlCommand(query, conn))
                        {
                            cmd.CommandType = CommandType.Text;
                            cmd.Parameters.AddWithValue("@priority", intPOPriority);
                            cmd.Parameters.AddWithValue("@sector", cbbPulledListLine.Text.Trim());
                            conn.Open();
                            object beforeObj = cmd.ExecuteScalar();
                            if (beforeObj != null && beforeObj != DBNull.Value)
                            {
                                po_bfr_co = beforeObj.ToString().Trim();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error retrieving PObefCO: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "NONE";
            }

            return po_bfr_co;
        }

        /// <summary>
        /// VB6: Public Sub AddLocalPhantomSubPO(ByVal getPONumber As String, ByVal getProductModel As String, ByVal getPOQty As Double)
        /// </summary>
        /// <param name="poNumber"></param>
        /// <param name="productModel"></param>
        /// <param name="poQty"></param>
        private void AddLocalPhantomSubPO(string poNumber, string productModel, double poQty) 
        {
            if (string.IsNullOrWhiteSpace(poNumber) || string.IsNullOrWhiteSpace(productModel))
                return;

            string targetFile = @"C:\MPH - KANBAN Control Local Data\PhantomSubMaterialsvsModel.txt";
            if (!File.Exists(targetFile))
            {
                Update_PhantomSubMaterialsvsModelTxt();
                return;
            }
                

            try
            {
                using (var sr = new StreamReader(targetFile))
                {
                    string line;
                    int rowStart = 0;

                    // Read file line by line (skip header) and process matching product model rows
                    while ((line = sr.ReadLine()) != null)
                    {
                        rowStart++;
                        var fields = line.Split('\t');

                        string getLocalProductModel = fields.Length > 0 ? fields[0].Trim() : string.Empty;

                        if (rowStart > 1 && !string.Equals(getLocalProductModel, "End", StringComparison.OrdinalIgnoreCase))
                        {
                            string getLocalSubModel = fields.Length > 1 ? fields[1].Trim() : string.Empty;
                            string getLocalMaterials = fields.Length > 2 ? fields[2].Trim() : string.Empty;
                            string getLocalMaterialsDesc = fields.Length > 3 ? fields[3].Trim() : string.Empty;
                            double getLocalQtyMaterialsPerProduct = fields.Length > 4 ? ConvertToDoubleSafe(fields[4]) : 0.0;

                            if (string.Equals(productModel, getLocalProductModel, StringComparison.OrdinalIgnoreCase))
                            {
                                bool existingMaterials = false;
                                double getQtyPO = getLocalQtyMaterialsPerProduct * poQty;

                                for (int ii = 0; ii < dgvPullListvsPO2.Rows.Count; ii++)
                                {
                                    double getCurrentQty = ConvertToDoubleSafe(GetGridCellAsString(dgvPullListvsPO2, ii, 5));
                                    string getPrevPOgroup = GetGridCellAsString(dgvPullListvsPO2, ii, 1).Trim();
                                    string getCurrentPOgroup = getPrevPOgroup;

                                    string gridMaterial = GetGridCellAsString(dgvPullListvsPO2, ii, 3).Trim();
                                    if (string.Equals(getLocalMaterials, gridMaterial, StringComparison.OrdinalIgnoreCase))
                                    {
                                        int getExistingPO = getCurrentPOgroup.IndexOf(poNumber, StringComparison.OrdinalIgnoreCase);
                                        if (getExistingPO == -1)
                                        {
                                            // If second and third char form "PO" (VB Mid(getCurrentPOgroup,2,2) = "PO")
                                            if (getCurrentPOgroup.Length >= 3 && getCurrentPOgroup.Substring(1, 2).Equals("PO", StringComparison.OrdinalIgnoreCase))
                                            {
                                                // getPrefixGroupPO = Val(Mid(getCurrentPOgroup, 1, 1)) + 1
                                                int prefix = 1;
                                                if (int.TryParse(getCurrentPOgroup.Substring(0, 1), out int parsedPrefix))
                                                    prefix = parsedPrefix + 1;

                                                // getSuffixGroupPO = Mid(getCurrentPOgroup, 7, Len(getCurrentPOgroup) - 6)
                                                string suffix = getCurrentPOgroup.Length > 6 ? getCurrentPOgroup.Substring(6) : string.Empty;

                                                getCurrentPOgroup = prefix.ToString() + "PO - " + suffix + "(" + poNumber + ")";
                                            }
                                            else
                                            {
                                                // else default to "2PO - (" & getCurrentPOgroup & ")(" & getPONumber & ")"
                                                getCurrentPOgroup = "2PO - (" + getCurrentPOgroup + ")(" + poNumber + ")";
                                            }
                                        }

                                        // update existing group and quantity
                                        SetGridCell(dgvPullListvsPO2, ii, 1, getCurrentPOgroup);
                                        getCurrentQty = getCurrentQty + getQtyPO;
                                        SetGridCell(dgvPullListvsPO2, ii, 5, getCurrentQty);

                                        existingMaterials = true;
                                        break;
                                    }
                                } // end for

                                if (!existingMaterials)
                                {
                                    // Append a new row
                                    int getRow = AddRowIfNeeded(dgvPullListvsPO2);
                                    SetGridCell(dgvPullListvsPO2, getRow, 0, getRow);
                                    // New entry should start with the PO number
                                    SetGridCell(dgvPullListvsPO2, getRow, 1, poNumber);
                                    // column 2 = model (use getProductModel)
                                    SetGridCell(dgvPullListvsPO2, getRow, 2, productModel);
                                    SetGridCell(dgvPullListvsPO2, getRow, 3, getLocalMaterials);
                                    SetGridCell(dgvPullListvsPO2, getRow, 4, getLocalMaterialsDesc);
                                    SetGridCell(dgvPullListvsPO2, getRow, 5, getQtyPO);
                                    SetGridCell(dgvPullListvsPO2, getRow, 6, "N");
                                    //// Ensure trailing empty row
                                    //AddRowIfNeeded(dgvPullListvsPO2);
                                }
                            } // end if product model match
                        } // end if rowStart > 1
                    } // end while
                } // end using sr
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddLocalPhantomSubPO failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// This is equivalent to the VB6: Public Sub AddOverallPackaging(getPONumber As String, getModel As String)
        /// </summary>
        /// <param name="poNumber"></param>
        /// <param name="productModel"></param>
        private void AddOverallPackaging(string poNumber, string productModel)
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (string.IsNullOrWhiteSpace(poNumber) || string.IsNullOrWhiteSpace(productModel))
                return;

            int getPOQty = 0;
            try
            {
                // find PO qty from flxPulledListPO (VB loop ii = 1 .. Rows - 1)
                var pulledList = dgvPulledListPO;
                for (int ii = 1; ii < pulledList.Rows.Count; ii++)
                {
                    if (string.Equals(GetGridCellAsString(pulledList, ii, 1), poNumber, StringComparison.OrdinalIgnoreCase))
                    {
                        getPOQty = (int)ConvertToDoubleSafe(GetGridCellAsString(pulledList, ii, 3));
                        break;
                    }
                }

                using (var cmd = new SqlCommand("SELECT * FROM PackagingMaterialsMatrixvsModel WHERE ProductModel = @model", cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.AddWithValue("@model", productModel);

                    bool openedHere = false;
                    try
                    {
                        if (cnnDLVNDB == null)
                            throw new InvalidOperationException("Database connection (cnnDLVNDB) is not initialized.");

                        if (cnnDLVNDB.State != ConnectionState.Open)
                        {
                            cnnDLVNDB.Open();
                            openedHere = true;
                        }

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader == null)
                                return;

                            // If there is at least one record, read first record to handle outer packaging (matches VB logic)
                            if (reader.Read())
                            {
                                // get outer packaging part and compute outer qty
                                string getOuterPackagingPart = (reader["OuterCartonPackaging"]?.ToString() ?? string.Empty).Trim();
                                double qtyProductPerOuter = ConvertToDoubleSafe(reader["QtyProductPerOuter"]?.ToString() ?? "0");
                                int getOutPackagingQty = 0;
                                if (qtyProductPerOuter > 0)
                                {
                                    // mimic original VB rounding: Round(Val(1 / Val(...)) * getPOQty + 0.5)
                                    double calc = (1.0 / qtyProductPerOuter) * getPOQty + 0.5;
                                    getOutPackagingQty = (int)Math.Floor(calc + 0.0000001); // floor to mimic VB integer result
                                }

                                string getMaterials = getOuterPackagingPart;
                                int accessMaterialsQtyPerPO = GetMaterialTotalQtyInPO(poNumber, getMaterials);

                                if (accessMaterialsQtyPerPO == 0)
                                {
                                    // update or append outer packaging into flxPullListvsPO2
                                    ApplyPackagingToPullList(poNumber, productModel, getMaterials, "OUTER PACKAGING MATERIALS", getOutPackagingQty);
                                }

                                // Now iterate starting from the first record again to process inner packaging for every row
                                // (VB loop used Do While Not EOF after processing outer packaging from the first record)
                                // We already are on the first row; process it and then continue with subsequent rows.
                                do
                                {
                                    string getInnerPackagingPart = (reader["InnerCartonPackaging"]?.ToString() ?? string.Empty).Trim();
                                    int qtyInnerPerOuter = (int)ConvertToDoubleSafe(reader["QtyInnerPerOuter"]?.ToString() ?? "0");
                                    int getInnerPackagingQty = getOutPackagingQty * qtyInnerPerOuter;
                                    getMaterials = getInnerPackagingPart;

                                    accessMaterialsQtyPerPO = GetMaterialTotalQtyInPO(poNumber, getMaterials);
                                    if (accessMaterialsQtyPerPO == 0)
                                    {
                                        ApplyPackagingToPullList(poNumber, productModel, getMaterials, "INNER PACKAGING MATERIALS", getInnerPackagingQty);
                                    }
                                } while (reader.Read());
                            }
                        }
                    }
                    finally
                    {
                        if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
                        {
                            try { cnnDLVNDB.Close(); } catch { /* swallow to match VB behaviour */ }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //// On VB error label, they called UpdateProgramRunningBug
                //try
                //{
                //    //UpdateProgramRunningBug("AddOverallPackaging", ex.HResult.ToString(), ex.Message);
                //}
                //catch { /* swallow */ }
                MessageBox.Show("AddOverallPackaging failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Vb6: Public Function GetMaterialsQtyPerPO(ByVal getPONumber As String, getMaterials As String) As Double;
        /// This will search for all rows in BOM txt file of the input PO number and sum up the qty of all rows of input material number
        /// </summary>
        /// <param name="inputPONumber"></param>
        /// <param name="inputMaterialNumber"></param>
        /// <returns></returns>
        public int GetMaterialTotalQtyInPO(string inputPONumber, string inputMaterialNumber)
        {
            if (string.IsNullOrWhiteSpace(inputPONumber) || string.IsNullOrWhiteSpace(inputMaterialNumber))
                return 0;

            // VB: If Mid(getPONumber, 1, 3) = "SMT" Then getPONumber = Mid(getPONumber, 1, Len(getPONumber) - 3)
            if (inputPONumber.Length >= 3 && inputPONumber.StartsWith("SMT", StringComparison.OrdinalIgnoreCase))
            {
                if (inputPONumber.Length > 3)
                {
                    inputPONumber = inputPONumber.Substring(0, inputPONumber.Length - 3);
                }
                else
                {
                    inputPONumber = string.Empty;
                }    
            }

            // Ensure local copy exists (original VB called this)
            try
            {
                CopyBOMTxtFilesFromSharedDriveToLocalFolder();
            }
            catch (Exception ex)
            {
                string errorMsg = "Error in CopyBOMTxtFilesFromSharedDriveToLocalFolder: " + ex.Message;
                MessageBox.Show(errorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            string targetFile = System.IO.Path.Combine(@"C:\MPH - KANBAN Control Local Data\LocalBOMToPulledList", inputPONumber + ".txt");

            if (!File.Exists(targetFile))
            {
                return 0;
            }

            int total = 0;
            try
            {
                using (var sr = new StreamReader(targetFile))
                {
                    string line;
                    int rowStart = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        // split by tab (VB used Split(..., vbTab))
                        var fields = line.Split('\t');

                        // Skip header lines (VB checks rowStart > 3)
                        if (rowStart > 3)
                        {
                            string infileMaterial = fields.Length > 2 ? fields[2].Trim() : string.Empty;
                            if (string.Equals(infileMaterial, inputMaterialNumber, StringComparison.OrdinalIgnoreCase)) //material-number-in-file matches input-material-number
                            {
                                string infileQtyAsText = fields.Length > 4 ? fields[4] : "0";
                                // Try parsing using invariant culture first, then fallback to replace comma with dot.
                                if (!Int32.TryParse(infileQtyAsText, NumberStyles.Any, CultureInfo.InvariantCulture, out int qty))
                                {
                                    var alt = infileQtyAsText?.Replace(',', '.');
                                    Int32.TryParse(alt, NumberStyles.Any, CultureInfo.InvariantCulture, out qty);
                                }
                                total += Math.Abs(qty);
                            }
                        }
                        rowStart++;
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = "Error in GetMaterialTotalQtyInPO: " + ex.Message;
                MessageBox.Show(errorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return total;
            }

            return total;
        }

        /// <summary>
        /// This is part of the function VB6: Public Sub AddOverallPackaging(getPONumber As String, getModel As String)
        /// </summary>
        /// <param name="getPONumber"></param>
        /// <param name="getModel"></param>
        /// <param name="getMaterials"></param>
        /// <param name="description"></param>
        /// <param name="qtyToAdd"></param>
        private void ApplyPackagingToPullList(string getPONumber, string getModel, string getMaterials, string description, int qtyToAdd)
        {
            var grid = dgvPullListvsPO2;
            bool existingMaterials = false;
            string getCurrentPOgroup = string.Empty;

            // VB loop: For ii = 1 To (.flxPullListvsPO2.Rows - 2)
            int maxIndex = Math.Max(0, grid.Rows.Count - 2);
            for (int ii = 1; ii <= maxIndex; ii++)
            {
                double getCurrentQty = ConvertToDoubleSafe(GetGridCellAsString(grid, ii, 5));
                string getPrevPOgroup = GetGridCellAsString(grid, ii, 1).Trim();
                getCurrentPOgroup = getPrevPOgroup;

                string gridMaterial = GetGridCellAsString(grid, ii, 3).Trim();
                if (string.Equals(getMaterials, gridMaterial, StringComparison.OrdinalIgnoreCase))
                {
                    int getExistingPO = getCurrentPOgroup.IndexOf(getPONumber, StringComparison.OrdinalIgnoreCase);
                    if (getExistingPO == -1)
                    {
                        if (getCurrentPOgroup.Length >= 3 && getCurrentPOgroup.Substring(1, 2).Equals("PO", StringComparison.OrdinalIgnoreCase))
                        {
                            int prefix = 1;
                            if (int.TryParse(getCurrentPOgroup.Substring(0, 1), out int parsedPrefix))
                                prefix = parsedPrefix + 1;
                            string suffix = getCurrentPOgroup.Length > 6 ? getCurrentPOgroup.Substring(6) : string.Empty;
                            getCurrentPOgroup = prefix.ToString() + "PO - " + suffix + "(" + getPONumber + ")";
                        }
                        else
                        {
                            getCurrentPOgroup = "2PO - (" + getCurrentPOgroup + ")(" + getPONumber + ")";
                        }
                    }

                    SetGridCell(grid, ii, 1, getCurrentPOgroup);
                    getCurrentQty = getCurrentQty + qtyToAdd;
                    SetGridCell(grid, ii, 5, getCurrentQty);
                    existingMaterials = true;
                    break;
                }
            }

            if (!existingMaterials)
            {
                int getRow = AddRowIfNeeded(grid);
                SetGridCell(grid, getRow, 0, getRow);

                // getExistingPO = InStr(getCurrentPOgroup, getPONumber) in VB; currentPOgroup here is empty so treat as not found
                SetGridCell(grid, getRow, 1, getPONumber);

                SetGridCell(grid, getRow, 2, getModel);
                SetGridCell(grid, getRow, 3, getMaterials);
                SetGridCell(grid, getRow, 4, description);
                SetGridCell(grid, getRow, 5, qtyToAdd);
                SetGridCell(grid, getRow, 6, "N");
                // mimic AddItem("")
                AddRowIfNeeded(grid);
            }
        }

        private void PhysicalBOMLog(DataGridView dgv, string fileName)
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\PhysicalBOMLogManipulation\" + fileName + ".txt";
            int countOfColumns = dgv.ColumnCount;
            int countOfRows = dgv.RowCount;

            try
            {
                // Ensure folder exists:
                if (!System.IO.Directory.Exists(@"C:\MPH - KANBAN Control Local Data\PhysicalBOMLogManipulation\"))
                {
                    System.IO.Directory.CreateDirectory(@"C:\MPH - KANBAN Control Local Data\PhysicalBOMLogManipulation\");
                }

                // Create txt file if not exists:
                if (!System.IO.File.Exists(targetFile))
                {
                    System.IO.File.Create(targetFile).Close();
                }

                // Build header line from row 0, columns 0 to countOfColumns-1
                string recordData = GetGridCellAsString(dgv, 0, 0);
                for (int ii = 1; ii < countOfColumns - 1; ii++)
                {
                    recordData += "\t" + GetGridCellAsString(dgv, 0, ii);
                }

                // Overwrite file with header (equivalent to ForWriting)
                File.WriteAllText(targetFile, recordData + Environment.NewLine);

                // Append data rows (equivalent to ForAppending)
                using (var sw = new StreamWriter(targetFile, append: true))
                {
                    for (int ii = 1; ii < countOfRows - 1; ii++)
                    {
                        string rowData = GetGridCellAsString(dgv, ii, 0);
                        for (int jj = 1; jj < countOfColumns; jj++)
                        {
                            rowData += "\t" + GetGridCellAsString(dgv, ii, jj);
                        }
                        sw.WriteLine(rowData);
                    }

                    // Write final "End" row (VB used ts.Write to write without newline for the last write)
                    string endRow = "End";
                    for (int jj = 1; jj < countOfColumns; jj++)
                    {
                        endRow += "\tEnd";
                    }
                    sw.Write(endRow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PhysicalBOMLog failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// VB6: Public Sub ClearOldDataBase()
        /// </summary>
        private void ClearOldDatabase()
        {
            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;

            string deletePOPair = @"DELETE FROM POpairunderCO " +
                "WHERE DATEDIFF(DAY, ActiveDateTime, GETDATE()) >= 15 " +
                "AND POaftCO NOT IN (SELECT DISTINCT TopPONumber FROM OpenPOPlanner WITH (NOLOCK))";

            string deletePulledList = @"DELETE FROM PulledListPOStatus " +
                "WHERE DATEDIFF(DAY, InProgressDateTime, GETDATE()) >= 15 " +
                "AND PONumber NOT IN (SELECT DISTINCT TopPONumber FROM OpenPOPlanner WITH (NOLOCK))";

            /*
             * When a query uses WITH (NOLOCK), 
             * it bypasses the standard locking mechanisms that prevent other transactions from modifying data being read, 
             * or that prevent the query from being blocked by other transactions holding exclusive locks. 
             * This can lead to faster data retrieval, especially in environments with high concurrency and heavy read workloads, 
             * as the query does not have to wait for locks to be released.
             */

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = deletePOPair;
                        conn.Open();
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = deletePulledList;
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error clearing old database: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BlockPlannerUserRight()
        {
            string user = "NA";
            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;
            string query1 = @"SELECT TOP 1 * FROM TrackingCurrentPOStatus WHERE Sector = @sector";
            string query2 = @"UPDATE PlanningPOUserRight SET AllowSkipConstraint = 'No' WHERE PlanningPOUserRight = @user";
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    using (var cmd = new SqlCommand(query1, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@sector", cbbPulledListLine.Text.Trim());
                        conn.Open();
                        var result = cmd.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                        {
                            user = result.ToString().Trim();
                            using (var cmd2 = new SqlCommand(query2, conn))
                            {
                                cmd2.CommandType = CommandType.Text;
                                cmd2.Parameters.AddWithValue("@user", user);
                                cmd2.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error blocking planner user rights: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// VB6 equivalent: //VB6: Public Sub CopyBOMServerToLocalFolderProgram()
        /// Copy PO's both BOM txt files (ex. SMT102099984.txt & 102099984.txt) - both are created using FinalAssyPullPlanning
        /// from shared drive ("\\vnmsrv300\pubfiles\ME\Truong_ME\Local_Backup_PO_PullList\") 
        /// to local folder ("C:\MPH - KANBAN Control Local Data\LocalBOMToPulledList\")
        /// </summary>
        private void CopyBOMTxtFilesFromSharedDriveToLocalFolder() 
        {
            string poNumber;
            string fileName;
            string fileSource;
            string fileDestination;

            // Loop through each PO in the added list dgvPulledListPO:
            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                poNumber = dgvPulledListPO.Rows[i].Cells["PONumber"].Value.ToString().Trim(); //Ex: SMT102099984A_G
                
                if (_pulledList_SectorGroup == "SMT_Group")
                {
                    // Copy SMT102099984.txt from shared drive to local folder:
                    fileName = poNumber.Substring(0, poNumber.Length - 3);//Ex: SMT102099984A_G >> Get SMT102099984
                    fileSource = @"\\vnmsrv300\pubfiles\ME\Truong_ME\Local_Backup_PO_PullList\" + fileName + ".txt";
                    fileDestination = @"C:\MPH - KANBAN Control Local Data\LocalBOMToPulledList\" + fileName + ".txt";
                    if (!File.Exists(fileDestination))
                    {
                        try
                        {
                            ExplodeBOMforPCBNoPhantomIfAny(fileSource, fileName);
                            if (File.Exists(fileSource))
                            {
                                File.Copy(fileSource, fileDestination, true);
                            }
                            else
                            {
                                MessageBox.Show("BOM file for PO " + poNumber + " does not exist on server:\r\n" + fileSource, "File Not Found:", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error copying BOM file for PO " + poNumber + ": " + ex.Message, "File Copy Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    // Copy 102099984.txt from shared drive to local folder:
                    fileName = poNumber.Substring(3, poNumber.Length - 6); //Ex: SMT102099984A_G >> Get 102099984
                    fileSource = @"\\vnmsrv300\pubfiles\ME\Truong_ME\Local_Backup_PO_PullList\" + fileName + ".txt";
                    fileDestination = @"C:\MPH - KANBAN Control Local Data\LocalBOMToPulledList\" + fileName + ".txt";
                    if (!File.Exists(fileDestination))
                    {
                        if (File.Exists(fileSource))
                        {
                            File.Copy(fileSource, fileDestination, true);
                        }
                        else
                        {
                            MessageBox.Show("BOM file for PO " + poNumber + " does not exist on server:\r\n" + fileSource, "File Not Found:", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                else
                {
                    //Not touching right now, since we do not use for anything other than SMT
                    MessageBox.Show("Only SMT Line is supported currently.", "Not Supported", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }

        /// <summary>
        /// VB6: Public Sub AccessPCBNoPhantom(ByVal strPOFileName As String, ByVal strPONumber As String);
        /// Input PO, 
        /// Search dgvPulledListPO to get model, side, and PO qty; 
        /// Search BOM txt file for model starting with "65*" or "66*"
        /// if any, query sql and write components to txt file
        /// </summary>
        /// <param name="poFilePath"></param>
        /// <param name="poNumber"></param>
        private void ExplodeBOMforPCBNoPhantomIfAny(string poFilePath, string poNumber)
        {
            bool blnNoPhantom = false;
            string pcbModel = "";
            string pcbSide = "";
            int poQty = 0;

            // Find model, side, and PO qty from dgvPulledListPO:
            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                string existingPO_1 = dgvPulledListPO.Rows[i].Cells["PONumber"].Value.ToString().Trim().Substring(0,12);
                string existingPO_2 = dgvPulledListPO.Rows[i].Cells["PONumber"].Value.ToString().Trim();
                if ((string.Equals(existingPO_1, poNumber, StringComparison.OrdinalIgnoreCase)) || (string.Equals(existingPO_2, poNumber, StringComparison.OrdinalIgnoreCase)))
                {
                    /*
                     dgvPulledListPO.Columns.Add("No", "No.");
                     dgvPulledListPO.Columns.Add("PONumber", "PONumber");
                     dgvPulledListPO.Columns.Add("ModelNumber", "ModelNumber");
                     dgvPulledListPO.Columns.Add("Side", "TOP/BOT");
                     dgvPulledListPO.Columns.Add("POQty", "PO Qty");
                     dgvPulledListPO.Columns.Add("PulledListID", "PulledListID");
                     dgvPulledListPO.Columns.Add("PlannersNotice", "Planner'sNotice");
                     dgvPulledListPO.Columns.Add("POChangeInf", "POChangeInf");
                     */
                    pcbModel = Convert.ToString(dgvPulledListPO.Rows[i].Cells["ModelNumber"].Value).Trim();
                    pcbSide = Convert.ToString(dgvPulledListPO.Rows[i].Cells["Side"].Value).Trim();
                    poQty = Convert.ToInt32(dgvPulledListPO.Rows[i].Cells["POQty"].Value.ToString().Trim());
                }
            }

            // Access the txt file to check if there's model starting with "65*" or "66*", flag blnNoPhantom true:
            if (File.Exists(poFilePath))
            {
                using (var sr = new StreamReader(poFilePath))
                {
                    string line;
                    int rowStart = 1;
                    while ((line = sr.ReadLine()) != null)
                    {
                        rowStart++;
                        var splitedFields = line.Split('\t');
                        if (rowStart > 4)
                        {
                            if (splitedFields.Length > 2)
                            {
                                string getLocalVal = splitedFields[2].Trim();
                                if (getLocalVal.Length >= 2)
                                {
                                    string prefix = getLocalVal.Substring(0, 2);
                                    if (prefix == "66" || prefix == "65")
                                    {
                                        blnNoPhantom = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("BOM file " + poFilePath + " does not exist to check for Phantom parts.", "File Not Found:", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // If blnNoPhantom true, access database to get parts list for the model and side, write to the BOM txt file:
            if (blnNoPhantom)
            {
                try
                {
                    using (var ts = new StreamWriter(poFilePath, append: true))
                    {
                        // Build SQL query
                        MSSQL _sql = new MSSQL();
                        string connectionString = _sql.cnnDLVNDB;
                        
                            string strQuery;
                        if (string.Equals(pcbSide, "A", StringComparison.OrdinalIgnoreCase))
                        {
                            strQuery = "SELECT [Model],[DLVNpn],SUM([PartUsed]) AS QtyUnit " +
                                       "FROM [DLVNDB].[dbo].[PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix] " +
                                       "WHERE MODEL = @pcbModel GROUP BY Model, DLVNpn";
                        }
                        else
                        {
                            strQuery = "SELECT [Model],[DLVNpn],SUM([PartUsed]) AS QtyUnit " +
                                       "FROM [DLVNDB].[dbo].[PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix] " +
                                       "WHERE MODEL = @pcbModel AND ModelSide = @pcbSide GROUP BY Model, DLVNpn"; //VB6: SMTLine not ModelSide
                        }
                        using (var cnnDLVNDB = new SqlConnection(connectionString))
                        {
                            cnnDLVNDB.Open();
                            using (var cmd = new SqlCommand(strQuery, cnnDLVNDB))
                            {
                                cmd.CommandType = System.Data.CommandType.Text;
                                cmd.Parameters.AddWithValue("pcbModel", pcbModel ?? string.Empty);
                                if (!string.Equals(pcbSide, "A", StringComparison.OrdinalIgnoreCase))
                                {
                                    cmd.Parameters.AddWithValue("pcbSide", pcbSide ?? string.Empty);
                                }      

                                using (var reader = cmd.ExecuteReader())
                                {
                                    if (reader != null)
                                    {
                                        while (reader.Read())
                                        {
                                            // VB: Mid(strPONumber,4,9)
                                            string midPo = poNumber.Length >= 4 ? (poNumber.Length >= 3 + 9 ? poNumber.Substring(3, 9) : poNumber.Substring(3)) : poNumber;
                                            string dlvnPn = (reader["DLVNpn"] ?? string.Empty).ToString().Trim();
                                            double qtyUnit = ConvertToDoubleSafe((reader["QtyUnit"] ?? "0").ToString());
                                            double totalQty = poQty * qtyUnit;

                                            string recordSetUpData = string.Join("\t",
                                                midPo,
                                                pcbModel,
                                                dlvnPn,
                                                "Description",
                                                totalQty.ToString(),
                                                "Units");

                                            ts.WriteLine(recordSetUpData);
                                        }
                                    }
                                }
                            }
                        }
                        
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("AccessPCBNoPhantom failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// VB6: Public Function IsPulledListExistinginLocal(getSector As String, getShift As String, getSumPOValID As String) As Boolean
        /// This will check folder "C:\MPH - KANBAN Control Local Data\PulledListLog" to see if a file like "PulledList_SMTG_A_12345_WithoutPK.txt" exists
        /// </summary>
        /// <param name="sector"></param>
        /// <param name="shift"></param>
        /// <param name="sumPOValID"></param>
        /// <returns></returns>
        private bool IsPulledListAlreadyInLocalFolder(string sector, string shift, string sumPOValID)
        {
            bool isPOChangedInf = false;
            bool isPulledListInLocal = false;
            string sector_name_part;
            string shift_name_part;

            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++) // keep consistent with VB loop bounds
            {
                if (Convert.ToString(dgvPulledListPO.Rows[i].Cells["POChangeInf"]) == "YES")
                {
                    isPOChangedInf = true;
                    break;
                }
            }

            // Ex. sector = "SMTG_Line", sector_name_part = "SMTG"
            if (sector != null && sector.Length > 5)
            {
                sector_name_part = sector.Substring(0, sector.Length - 5);
            }
            else
            {
                sector_name_part = "NoSectorInfo";
            }

            // Ex. shift = "Shift_A", shift_name_part = "A"
            if (shift != null && shift.Length >= 7)
            {
                shift_name_part = shift.Substring(6, 1);
            }
            else
            {
                shift_name_part = "NoShiftInfo";
                //MessageBox.Show("Shift info is invalid when checking pulled list in local folder.", "Invalid Shift Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            string getFileName = $"PulledList_{sector_name_part}_{shift_name_part}_{sumPOValID}_WithoutPK";
            string targetDir = @"C:\MPH - KANBAN Control Local Data\PulledListLog";
            string targetFile = System.IO.Path.Combine(targetDir, getFileName + ".txt");

            if (!File.Exists(targetFile))
            {
                isPulledListInLocal = false;
                return isPulledListInLocal;
            }
            else
            {
                if (isPOChangedInf)
                {
                    try 
                    { 
                        File.Delete(targetFile); 
                    } 
                    catch 
                    { 
                        /* log if needed */ 
                        MessageBox.Show("Unable to delete existing pulled list file with PO change info. Please check if the file is open or in use.", "File Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    isPulledListInLocal = false;
                    return isPulledListInLocal;
                }
                else
                {
                    isPulledListInLocal = true;
                    return isPulledListInLocal;
                }
            }
        }

        /// <summary>
        /// VB6: Public Function AccessPastListPO() As Boolean;
        /// this will always returns true for now;
        /// Clear dgvPullListvsPO & dgvMultiUniPhysicalModelPulled;
        /// Do PullPOToPart >> LoadBOMFromSharedDriveTo_dgvPLPhysicalModelPulled >> Populate dgvPLPhysicalModelPulled
        /// Populate dgvUniPhysicalModelPulled from dgvMultiUniPhysicalModelPulled;
        /// </summary>
        /// <returns></returns>
        private bool AccessPastListPO()
        {
            string modelNumber;
            string poNumber;
            bool blnResult = true; //default return as true, without compsapvslocal, this procedure always returns true

            //try
            //{
                // Clear existing rows in dgvPullListvsPO:
                dgvPullListvsPO.Rows.Clear();

                // Clear existing rows in dgvMultiUniOverallModelPulled:
                dgvMultiUniPhysicalModelPulled.Rows.Clear();

                _blnMultiPOExport = true;
                //intGetPrevStartRow = 1;

                for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
                {
                    // Clear existing rows in dgvUniPhysicalModelPulled:
                    dgvUniPhysicalModelPulled.Rows.Clear();
                    dgvUniPhysicalModelPulled.Refresh();

                    // Clear existing rows in dgvQtyvsCountDuplicated and add a new row:
                    dgvQtyvsCountDuplicated.Rows.Clear();
                    dgvQtyvsCountDuplicated.Refresh();

                    // Get values from dgvPulledListPO:
                    poNumber = GetGridCellAsString(dgvPulledListPO, i, 1).Trim(); //PONumber
                    modelNumber = GetGridCellAsString(dgvPulledListPO, i, 2).Trim(); //ModelNumber

                    // If PO file does not exist in local folder, pull model and PO BOM (just like in add plans):
                    if (!IsPOBOMTxtAlreadyInLocalFolder(poNumber))
                    {
                        PullModelAfterCOtoPart(modelNumber, cbbPulledListLine.Text);
                        PullPOAfterCOtoPart(poNumber, modelNumber);
                    }
                    
                    // Populate dgvPLPhysicalModelPulled & dgvUniPhysicalModelPulled & dgvMultiUniPhysicalModelPulled:
                    PullPOtoPart(poNumber, modelNumber);
                }

                // Clear existing rows in dgvPLPhysicalModelPulled:
                dgvPLPhysicalModelPulled.Rows.Clear();
                dgvPLPhysicalModelPulled.Refresh();

                // Clear existing rows in dgvUniPhysicalModelPulled:
                dgvUniPhysicalModelPulled.Rows.Clear();
                dgvUniPhysicalModelPulled.Refresh();

                // Use dgvMultiUniPhysicalModelPulled to populate dgvUniPhysicalModelPulled:
                for (int i = 0; i < dgvMultiUniPhysicalModelPulled.Rows.Count; i++)
                {
                    //var row = dgvMultiUniPhysicalModelPulled.Rows[i];
                    var material_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 3).Trim();
                    if (!string.IsNullOrEmpty(material_in_multi_uni))
                    {
                        //'flxMultiUniPhysicalModelPulled.FormatString = "^No.|^PONumber|^Model|^Part Name|^Part Description|^Qty|^UOM|^PCBABin|^CommonPart|^TopBot"
                        //'flxUniPhysicalModelPulled.FormatString = "^No.|^PONumber|^Model|^Part Name|^Part Description|^Part Qty|^Part UOM|^PCBABin|^TopBot"
                        
                        var po_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 1).Trim();
                        var model_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 2).Trim();
                        var partdesc_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 4).Trim();
                        var qty_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 5).Trim();
                        var uom_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 6).Trim();
                        var topbot_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 8).Trim();

                        var index_in_uni = AddRowIfNeeded(dgvUniPhysicalModelPulled);
                        SetGridCell(dgvUniPhysicalModelPulled, i, 0, index_in_uni);
                        SetGridCell(dgvUniPhysicalModelPulled, i, 1, po_in_multi_uni);
                        SetGridCell(dgvUniPhysicalModelPulled, i, 2, model_in_multi_uni);
                        SetGridCell(dgvUniPhysicalModelPulled, i, 3, material_in_multi_uni);
                        SetGridCell(dgvUniPhysicalModelPulled, i, 4, partdesc_in_multi_uni);
                        SetGridCell(dgvUniPhysicalModelPulled, i, 5, qty_in_multi_uni);
                        SetGridCell(dgvUniPhysicalModelPulled, i, 6, uom_in_multi_uni);
                        SetGridCell(dgvUniPhysicalModelPulled, i, 7, "");
                        SetGridCell(dgvUniPhysicalModelPulled, i, 8, topbot_in_multi_uni);
                    }
                    dgvUniPhysicalModelPulled.Refresh();
                }

                // Sort both grids by column index 3 (if that column exists)
                if (dgvUniPhysicalModelPulled?.Columns.Count > 3)
                {
                    dgvUniPhysicalModelPulled.Sort(dgvUniPhysicalModelPulled.Columns[3], ListSortDirection.Ascending);

                }

                if (dgvMultiUniPhysicalModelPulled?.Columns.Count > 3)
                {
                    dgvMultiUniPhysicalModelPulled.Sort(dgvMultiUniPhysicalModelPulled.Columns[3], ListSortDirection.Ascending);
                }

                // Set "No." column values for Uni grid
                int counter = 1;
                foreach (DataGridViewRow r in dgvUniPhysicalModelPulled?.Rows ?? new DataGridViewRowCollection(new DataGridView()))
                {
                    if (r.IsNewRow)
                    {
                        continue;//skip if new row
                    }
                    if (r.Cells.Count > 0)
                    {
                        r.Cells[0].Value = counter++;
                    }
                }

                // Set "No." column values for Multi-Uni grid
                counter = 1;
                foreach (DataGridViewRow r in dgvMultiUniPhysicalModelPulled?.Rows ?? new DataGridViewRowCollection(new DataGridView()))
                {
                    if (r.IsNewRow)
                    {
                        continue;
                    }//skip if new row
                    if (r.Cells.Count > 0)
                    {
                        r.Cells[0].Value = counter++;
                    }
                }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error accessing past list PO: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}
            //finally
            //{
            //    dgvPLOverallModelPulled?.Refresh();
            //    dgvPLPhysicalModelPulled?.Refresh();
            //}
            return blnResult;//is it always true?
        }

        /// <summary>
        /// VB6: Public Sub PullModelAfterCOtoPart(getModel As String, getSector As String)
        /// </summary>
        /// <param name="model"></param>
        /// <param name="sectorName"></param>
        private void PullModelAfterCOtoPart(string modelNumber, string sectorName)
        {
            string strModelvsPhysicalMaterialsDBTable = "ModelvsPhysicalMaterials_" + sectorName;

            // Clear existing rows in dgvLocalPart and add a new row:
            dgvLocalPart.Rows.Clear();

            // Clear existing rows in dgvPhysicalModelAfterCOPulled and add a new row:
            dgvPhysicalModelAfterCOPulled.Rows.Clear();

            //
            MSSQL _sql = new MSSQL();
            string connectionString = _sql.cnnDLVNDB;
            string query = @"SELECT * FROM " + strModelvsPhysicalMaterialsDBTable + " WHERE Model = @model";

            using (var conn = new SqlConnection(connectionString))
            {
                using (var cmd = new SqlCommand(query, conn))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.AddWithValue("@model", modelNumber);
                    conn.Open();
                    using (var reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        if (reader != null && reader.HasRows)
                        {
                            int ii;
                            while (reader.Read())
                            {
                                string activeMaterials = (reader["Materials"] ?? string.Empty).ToString().Trim();

                                // skip if empty
                                if (string.IsNullOrEmpty(activeMaterials))
                                    continue;

                                // Equivalent of:
                                // If (Mid(activeMaterials, 1, 1) <> "R") And (Mid(activeMaterials, 1, 4) <> "6100") And (Mid(activeMaterials, 1, 2) <> "DR") Then
                                bool startsWithR = activeMaterials.StartsWith("R", StringComparison.OrdinalIgnoreCase);
                                bool startsWith6100 = activeMaterials.Length >= 4 && activeMaterials.Substring(0, 4).Equals("6100", StringComparison.OrdinalIgnoreCase);
                                bool startsWithDR = activeMaterials.Length >= 2 && activeMaterials.Substring(0, 2).Equals("DR", StringComparison.OrdinalIgnoreCase);

                                if (!startsWithR && !startsWith6100 && !startsWithDR)
                                {
                                    // QtyperProduct may be numeric or string
                                    string qtyPerProduct = (reader["QtyperProduct"] ?? string.Empty).ToString();

                                    if (!CheckDuplicatedPhysicalModelAfterCOMaterials(activeMaterials))
                                    {
                                        // last row index (equivalent to VB ii = Rows - 1)
                                        ii = AddRowIfNeeded(dgvPhysicalModelAfterCOPulled);
                                        dgvPhysicalModelAfterCOPulled.Rows[ii].Cells[0].Value = ii;
                                        dgvPhysicalModelAfterCOPulled.Rows[ii].Cells[1].Value = activeMaterials;
                                        dgvPhysicalModelAfterCOPulled.Rows[ii].Cells[3].Value = qtyPerProduct;
                                    }

                                    if (!CheckLocalDuplicatedMaterials(activeMaterials))
                                    {
                                        ii = AddRowIfNeeded(dgvLocalPart);
                                        dgvLocalPart.Rows[ii].Cells[0].Value = ii;
                                        dgvLocalPart.Rows[ii].Cells[1].Value = activeMaterials;
                                        dgvLocalPart.Rows[ii].Cells[3].Value = qtyPerProduct;
                                    }
                                }
                            } // while reader.Read()
                        }
                        else
                        {
                            // No rows found
                            MessageBox.Show(
                                $"Chuong Trinh Khong Tim Thay Model: {modelNumber} Trong Co So Du Lieu!!{Environment.NewLine}Vui Long Lien He Ky Su De Ho Tro!!",
                                "DataBase Not Found: PullModelAfterCOtoPart",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                    } // using reader
                }
            }
        }

        private bool CheckDuplicatedPhysicalModelAfterCOMaterials(string active_material) //VB6: Private Function CheckDuplicatedPhysicalModelAfterCOMaterials(getActiveMaterials As String) As Boolean
        {
            var grid = dgvPhysicalModelAfterCOPulled;

            // VB: For jj = 1 To (Rows - 1)
            // translates to start at index 1 and iterate while jj < Rows.Count
            for (int jj = 1; jj < grid.Rows.Count; jj++)
            {
                var row = grid.Rows[jj];
                if (row.IsNewRow)
                    continue;

                var cellVal = row.Cells[1].Value?.ToString();
                if (string.Equals(active_material, cellVal, StringComparison.OrdinalIgnoreCase))
                {
                    return true; // duplicate found
                }
            }

            return false; // no duplicate found
        }

        private bool CheckLocalDuplicatedMaterials(string active_material) //VB6: Private Function CheckLocalDuplicatedMaterials(getActiveMaterials As String) As Boolean
        {
            var grid = dgvLocalPart;

            // VB: For jj = 1 To (Rows - 1)
            // translate to start at index 1 and iterate while jj < Rows.Count
            for (int jj = 1; jj < grid.Rows.Count; jj++)
            {
                var row = grid.Rows[jj];
                if (row.IsNewRow)
                    continue;

                var cellVal = row.Cells[1].Value?.ToString();
                if (string.Equals(active_material, cellVal, StringComparison.OrdinalIgnoreCase))
                {
                    return true; // duplicate found
                }
            }

            return false; // no duplicate found
        }

        /// <summary>
        /// VB6: Public Sub PullPOAfterCOtoPart(prodorder As String, getModel As String)
        /// </summary>
        /// <param name="po"></param>
        /// <param name="model"></param>
        private void PullPOAfterCOtoPart(string po, string model)
        {
            
        }

        /// <summary>
        /// VB6: Public Sub PullPOtoPart(ByVal prodorder As String, getModel As String)
        /// </summary>
        /// <param name="po"></param>
        /// <param name="model"></param>
        private void PullPOtoPart(string poNumber, string modelNumber)
        {
            int number_of_CommonPO = 1;
            string originalPO = string.Empty;

            // PO input usually in the form: SMT102099984A_G
            if (poNumber.Substring(0, 3) == "SMT")
            {
                originalPO = poNumber.Substring(3, poNumber.Length - 6); //Ex: SMT102099984A_G >> Get 102099984 ; used only for GetBOMSAPToDataGridViews and SaveToLocalPOPhyicalMaterialsPullList
            }

            // Clear existing rows in dgvPLOverallModelPulled:
            dgvPLOverallModelPulled.Rows.Clear();
            dgvPLOverallModelPulled.Refresh();

            // Clear existing rows in dgvPLPhysicalModelPulled:
            dgvPLPhysicalModelPulled.Rows.Clear();
            dgvPLPhysicalModelPulled.Refresh();

            // Get the Qty of PO and paste to dgvQtyvsCountDuplicated:
            GetMaxQtyDuplicated(poNumber);

            // If PO text file does not exist in local folder, get BOM SAP and save to local folder:
            if (!IsPOBOMTxtAlreadyInLocalFolder(poNumber))
            {
                GetBOMSAPToDataGridViews(originalPO, modelNumber);
                SaveToLocalPOPhyicalMaterialsPullList(poNumber, modelNumber, "DicrectPO");
            }
            else
            {
                if (_pulledList_SectorGroup == "SMT_Group")
                {
                    string poSide = GetPOTopBot(poNumber);

                    // Load single PO's BOM to dgvPLPhysicalModelPulled:
                    LoadBOMFromSharedDriveTo_dgvPLPhysicalModelPulled(poNumber, modelNumber, poSide, "DicrectPO");
                }
                else //almost no use anymore
                {
                    MessageBox.Show("Only SMT Line is supported currently.", "Not Supported", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //return;
                }
            }

            // Loop for all materials in PO's BOM (from dgvPLPhysicalModelPulled), copy to dgvUniPhysicalModelPulled
            // if the material did not exist in dgvUniPhysicalModelPulled;
            // if already exist, add the quantity
            for (int i = 0; i < dgvPLPhysicalModelPulled.Rows.Count; i++)
            {
                bool blnMaterialAlreadyInUniGrid = false;

                // If part already exists, set flag to skip adding row:
                for (int j = 0; j < dgvUniPhysicalModelPulled.Rows.Count; j++)
                {
                    //dgvUniPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); PCBABin(7); TopBot(8)
                    //dgvPLPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); TopBot(7)
                    var material_in_uni = GetGridCellAsString(dgvUniPhysicalModelPulled, j, 3); //PartName
                    var material_in_pl = GetGridCellAsString(dgvPLPhysicalModelPulled, i, 3); //PartName
                    
                    if (material_in_uni == material_in_pl)
                    {
                        blnMaterialAlreadyInUniGrid = true;
                        break;
                    }
                }

                // If part does not exist in UniPhysicalModelPulled >> add row to UniPhysicalModelPulled:
                if (!blnMaterialAlreadyInUniGrid)
                {
                    var qty_in_pl = Convert.ToInt32(GetGridCellAsObject(dgvPLPhysicalModelPulled, i, 5)); //Qty
                    if (qty_in_pl > 0) //Col Qty
                    { 
                        var model_in_pl = GetGridCellAsString(dgvPLPhysicalModelPulled, i, 2);
                        var material_in_pl = GetGridCellAsString(dgvPLPhysicalModelPulled, i, 3);
                        var partdesc_in_pl = GetGridCellAsString(dgvPLPhysicalModelPulled, i, 4);
                        var uom_in_pl = GetGridCellAsString(dgvPLPhysicalModelPulled, i, 6);
                        var topbot_in_pl = GetGridCellAsString(dgvPLPhysicalModelPulled, i, 7);

                        int row_index_uni = AddRowIfNeeded(dgvUniPhysicalModelPulled);
                        SetGridCell(dgvUniPhysicalModelPulled, row_index_uni, 0, row_index_uni);
                        SetGridCell(dgvUniPhysicalModelPulled, row_index_uni, 1, poNumber);
                        SetGridCell(dgvUniPhysicalModelPulled, row_index_uni, 2, model_in_pl); //Model
                        SetGridCell(dgvUniPhysicalModelPulled, row_index_uni, 3, material_in_pl); //PartName
                        SetGridCell(dgvUniPhysicalModelPulled, row_index_uni, 4, partdesc_in_pl); //PartDesc
                        SetGridCell(dgvUniPhysicalModelPulled, row_index_uni, 5, "0"); //Qty
                        SetGridCell(dgvUniPhysicalModelPulled, row_index_uni, 6, uom_in_pl); //UOM
                        SetGridCell(dgvUniPhysicalModelPulled, row_index_uni, 7, ""); //PCBABin left blank for now
                        SetGridCell(dgvUniPhysicalModelPulled, row_index_uni, 8, topbot_in_pl); //TopBot
                        dgvUniPhysicalModelPulled.Refresh();
                    } 
                }
            }

            // Loop the dgvUniPhysicalModelPulled (it may contains previously processed POs' BOM), add the quantity:
            for (int i = 0; i < dgvUniPhysicalModelPulled.Rows.Count; i++)
            {
                //dgvUniPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); PCBABin(7); TopBot(8)
                //dgvPLPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); TopBot(7)
                var material_in_uni = GetGridCellAsString(dgvUniPhysicalModelPulled, i, 3);
                var qty_in_uni_as_string = GetGridCellAsString(dgvUniPhysicalModelPulled, i, 5);

                int current_qty_in_uni = Convert.ToInt32(qty_in_uni_as_string);
                for (int j = 0; j < dgvPLPhysicalModelPulled.Rows.Count; j++)
                {
                    var material_in_pl = GetGridCellAsString(dgvPLPhysicalModelPulled, j, 3);
                    var qty_in_pl = GetGridCellAsString(dgvPLPhysicalModelPulled, j, 5);

                    if (material_in_uni == material_in_pl)
                    {
                        current_qty_in_uni = current_qty_in_uni + Convert.ToInt32(qty_in_pl);
                    }
                }

                // Set Material's New Qty to dgvUniPhysicalModelPulled:
                SetGridCell(dgvUniPhysicalModelPulled, i, 5, current_qty_in_uni);
                dgvUniPhysicalModelPulled.Refresh();
            }

            // if we export to multiple POs, add rows to dgvMultiUniPhysicalModelPulled:
            if (_blnMultiPOExport)
            {
                int j = 0;//We need to keep j value for later
                for (int i = 0; i < dgvUniPhysicalModelPulled.Rows.Count; i++)
                {
                    //dgvUniPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); PCBABin(7); TopBot(8)
                    bool blnMaterialAlreadyInMultiUniGrid = false;
                    var material_in_uni = GetGridCellAsString(dgvUniPhysicalModelPulled, i, 3); 
                    var po_in_uni = GetGridCellAsString(dgvUniPhysicalModelPulled, i, 1);
                    var model_in_uni = GetGridCellAsString(dgvUniPhysicalModelPulled, i, 2);
                    var partdesc_in_uni = GetGridCellAsString(dgvUniPhysicalModelPulled, i, 4);
                    var qty_in_uni = GetGridCellAsString(dgvUniPhysicalModelPulled, i, 5);
                    var uom_in_uni = GetGridCellAsString(dgvUniPhysicalModelPulled, i, 6); //ex. ST2
                    var topbot_in_uni = GetGridCellAsString(dgvUniPhysicalModelPulled, i, 8); //ex. T or B or A

                    for (j = 0; j < dgvMultiUniPhysicalModelPulled.Rows.Count; j++)
                    {
                        var material_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 3);
                        if (material_in_uni == material_in_multi_uni)
                        {
                            blnMaterialAlreadyInMultiUniGrid = true;
                            break;
                        }
                    }
                    
                    var class_of_model_in_uni = GetMaterialClass(material_in_uni.Trim());

                    if (!string.IsNullOrEmpty(class_of_model_in_uni) && class_of_model_in_uni.Substring(0,4) == "PCBA" && _pulledList_SectorGroup != "SMT_Group")
                    {
                        blnMaterialAlreadyInMultiUniGrid = false;
                    }
                    
                    if (!blnMaterialAlreadyInMultiUniGrid)
                    {
                        int index_row_multi_uni = AddRowIfNeeded(dgvMultiUniPhysicalModelPulled);
                        SetGridCell(dgvMultiUniPhysicalModelPulled, index_row_multi_uni, 1, po_in_uni);
                        SetGridCell(dgvMultiUniPhysicalModelPulled, index_row_multi_uni, 2, model_in_uni);
                        SetGridCell(dgvMultiUniPhysicalModelPulled, index_row_multi_uni, 3, material_in_uni);
                        SetGridCell(dgvMultiUniPhysicalModelPulled, index_row_multi_uni, 4, partdesc_in_uni); 
                        SetGridCell(dgvMultiUniPhysicalModelPulled, index_row_multi_uni, 5, qty_in_uni);
                        SetGridCell(dgvMultiUniPhysicalModelPulled, index_row_multi_uni, 6, uom_in_uni);
                        SetGridCell(dgvMultiUniPhysicalModelPulled, index_row_multi_uni, 9, topbot_in_uni);

                        var commonpart_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, index_row_multi_uni, 8);
                        string commonPartValue = string.Empty;
                        if (commonpart_in_multi_uni == "" ||string.IsNullOrEmpty(commonpart_in_multi_uni))
                        {
                            number_of_CommonPO = 1;
                            commonPartValue = Convert.ToString(number_of_CommonPO) 
                                + "PO - (" 
                                + po_in_uni
                                + ")"; //ex. 1PO - (SMT102160895B_B)
                            SetGridCell(dgvMultiUniPhysicalModelPulled, index_row_multi_uni, 8, commonPartValue);
                        }
                        dgvUniPhysicalModelPulled.Refresh();
                        //dgvMultiUniPhysicalModelPulled.Rows.Add("", "*xXx*", "*xXx*", "*xXx*", "*xXx*", "", "*xXx*", "*xXx*", "", "*xXx*");
                    }
                    else
                    {
                        // If this material already appears in dgvMultiUniPhysicalModelPulled, we need to update Qty and CommonPart;
                        // Material already exists in Multi-Uni grid at row index j;
                        // We add the Qty from dgvUniPhysicalModelPulled grid to existing row in dgvMultiUniPhysicalModelPulled;
                        var commonpart_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 8);
                        if (commonpart_in_multi_uni != "*xXx*")
                        {
                            if (commonpart_in_multi_uni.Length >= 1)
                            {
                                //Ex: 1PO - ((SMT102160895B_B) >> get "1" >> turn to "2" >> 2PO - (SMT102160895B_B)(SMT102160902A_B)
                                var first_char_commonpart_in_multi_uni = commonpart_in_multi_uni.Substring(0, 1); //get the first char which is the number of common PO
                                number_of_CommonPO = Convert.ToInt32(first_char_commonpart_in_multi_uni) + 1; 
                            }
                            else
                            {
                                // Handle case where commonpart_in_multi_uni is empty or too short
                                MessageBox.Show("CommonPart field is empty or invalid when processing dgvMultiUniPhysicalModelPulled:", "PullPOToPart Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                number_of_CommonPO = -1; //Signal error
                                //continue; // skip to next iteration
                            }
                            
                            var qty_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 5);
                            var total_qty = Convert.ToInt32(qty_in_multi_uni) + Convert.ToInt32(qty_in_uni);
                            var po_number_in_multi_uni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, j, 1);
                            var newCommonPartValue = Convert.ToString(number_of_CommonPO) + "PO - " //this is "2PO - "
                                + commonpart_in_multi_uni.Substring(6, commonpart_in_multi_uni.Length - 6) //this is "(SMT102160895B_B)"
                                + "("
                                + po_number_in_multi_uni
                                + ")";//this is "(SMT102160902A_B)"
                            //result is "2PO - ((SMT102160895B_B)(SMT102160902A_B)"

                            SetGridCell(dgvMultiUniPhysicalModelPulled, j, 5, total_qty); //Set Qty
                            SetGridCell(dgvMultiUniPhysicalModelPulled, j, 8, newCommonPartValue); //Set CommonPart
                            dgvMultiUniPhysicalModelPulled.Refresh();
                        }
                    }
                }
            }

            if (_repeatedProcess == 0)
            {
                PhysicalBOMLog(dgvMultiUniPhysicalModelPulled, "010_PullBOM_MultiUniPhysicalModelPulled");
            }
            else
            {
                PhysicalBOMLog(dgvMultiUniPhysicalModelPulled, "110_PullBOM_MultiUniPhysicalModelPulled");
            }
        }

        /// <summary>
        /// VB6: Public Function AccessMaterialsClass(getMaterials As String);
        /// </summary>
        /// <param name="modelNumber"></param>
        /// <returns></returns>
        private string GetMaterialClass(string modelNumber)
        {
            string strResult = "";
            if (IsFinishedGoods(modelNumber))
            {
                if (IsPCBAModel(modelNumber))
                {
                    string strBin = GetPCBABin(modelNumber);
                    strResult = "SUB-ASSY:" + strBin;
                }
                else
                {
                    strResult = "Finished Goods";
                }
            }
            else if (IsPCBAModel(modelNumber))
            {
                string strBin = GetPCBABin(modelNumber);
                strResult = "PCBA:" + strBin;
            }
            else if (IsPackagingMROPart(modelNumber))
            {
                if (IsMROBox(modelNumber))
                {
                    strResult = "PACKAGING";
                }
                else
                {
                    strResult = "MRO";
                }
            }
            else
            {
                strResult = string.Empty;
            }
            return strResult;
        }

        /// <summary>
        /// Public Function IsFinishedGood(ByVal getMaterials As String) As Boolean
        /// </summary>
        /// <param name="partNumber"></param>
        /// <returns></returns>
        private bool IsFinishedGoods(string partNumber)
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\MPHAllSector.txt";

            if (string.IsNullOrEmpty(partNumber))
            {
                MessageBox.Show("Part number is null or empty when checking Finished Goods.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }        

            if (!File.Exists(targetFile))
            {
                MessageBox.Show("Finished Goods data file not found: " + targetFile, "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Update_MPHAllSectorTxt();
            }    

            try
            {
                int rowStart = 0;
                foreach (var line in File.ReadLines(targetFile))
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        var fields = line.Split('\t');

                        // VB logic: only evaluate when rowStart > 0 and first field <> "End"
                        if (rowStart > 0 && fields.Length > 0 && !string.Equals(fields[0].Trim(), "End", StringComparison.Ordinal))
                        {
                            var localModel = fields[0].Trim();
                            var sectorType = fields[4].Trim();
                            if (string.Equals(partNumber, localModel, StringComparison.Ordinal))
                            {
                                if (sectorType == "FA")
                                {
                                    return true;
                                }
                                else
                                {
                                    return false;
                                }   
                            }
                        }
                    }

                    rowStart++;
                }
            }
            catch
            {
                // Mirror VB permissive behavior: swallow errors and return false.
                return false;
            }
            return false;
        }

        /// <summary>
        /// Public Function IsMROBox(getPartNumber As String) As Boolean
        /// </summary>
        /// <param name="partNumber"></param>
        /// <returns></returns>
        private bool IsMROBox(string partNumber)
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\MPHMROList.txt";

            if (string.IsNullOrEmpty(partNumber))
            {
                MessageBox.Show("Part number is null or empty when checking if MRO Box.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!File.Exists(targetFile))
            {
                MessageBox.Show("MRO List data file not found: " + targetFile, "File Not Found;\r\nUpdate C:\\MPH - KANBAN Control Local Data\\MPHMROList.txt", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Update_MPHMROListTxt();
            }
            else
            {
                try
                {
                    using (var sr = new StreamReader(targetFile))
                    {
                        string line;
                        int rowStart = 1;
                        while ((line = sr.ReadLine()) != null)
                        {
                            rowStart++;
                            var fields = line.Split('\t');

                            string getLocalMROPartName = fields.Length > 0 ? fields[0].Trim() : string.Empty;
                            string getLocalIsBox = fields.Length > 4 ? fields[4].Trim() : string.Empty;

                            if (string.Equals(getLocalMROPartName, partNumber, StringComparison.OrdinalIgnoreCase))
                            {
                                if (getLocalMROPartName == partNumber)
                                {
                                    if (getLocalIsBox == "NO")
                                    {
                                        return false;
                                    }
                                    else
                                    {
                                        return true;
                                    }
                                }     
                            }
                        }
                    }
                }
                catch
                {
                    // Mirror VB permissive behavior: swallow errors and return false.
                    return false;
                }
            }    
            return false;
        }

        /// <summary>
        /// Public Function IsPackagingMROPart(getMaterials As String) As Boolean
        /// </summary>
        /// <param name="partNumber"></param>
        /// <returns></returns>
        private bool IsPackagingMROPart(string partNumber)
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\MPHMROList.txt";

            if (string.IsNullOrEmpty(partNumber))
            {
                MessageBox.Show("Part number is null or empty when checking if MRO Box.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!File.Exists(targetFile))
            {
                MessageBox.Show("MRO List data file not found: " + targetFile, "File Not Found;\r\nUpdate C:\\MPH - KANBAN Control Local Data\\MPHMROList.txt", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Update_MPHMROListTxt();
            }
            else
            {
                try
                {
                    using (var sr = new StreamReader(targetFile))
                    {
                        string line;
                        int rowStart = 1;
                        while ((line = sr.ReadLine()) != null)
                        {
                            rowStart++;
                            var fields = line.Split('\t');

                            string getLocalMROPartName = fields.Length > 0 ? fields[0].Trim() : string.Empty;
                            string getLocalIsBox = fields.Length > 4 ? fields[4].Trim() : string.Empty;

                            if (string.Equals(getLocalMROPartName, partNumber, StringComparison.OrdinalIgnoreCase))
                            {
                                // VB logic: If getLocalIsBox = "NO" Then IsMROBox = False Else True
                                return !string.Equals(getLocalIsBox, "NO", StringComparison.OrdinalIgnoreCase);
                            }
                        }
                    }
                }
                catch
                {
                    // Mirror VB permissive behavior: swallow errors and return false.
                    return false;
                }
            }
            return false;
        }

        /// <summary>
        /// //Public Function IsPreAssyComponent(getPartNumber As String) As Boolean
        /// </summary>
        /// <param name="partNumber"></param>
        /// <returns></returns>
        private bool IsPCBComponent(string partNumber) 
        {
            string TargetFile = @"C:\MPH - KANBAN Control Local Data\WareHouseMaterialsSource.txt";

            bool isPostAssyComponent = false;
            // Try to call a local implementation if available; if not, assume false.
            try 
            { 
                isPostAssyComponent = IsPCBAComponent(partNumber); 
            } 
            catch (Exception e)
            { 
                MessageBox.Show("Error checking PCBA component: " + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isPostAssyComponent = false; 
            }
            
            if (isPostAssyComponent)
            {
                return false;
            }
                

            if (!File.Exists(TargetFile))
            {
                Update_WareHouseMaterialsSourceTxt();
            }  

            try
            {
                int rowStart = 0;
                foreach (var line in File.ReadLines(TargetFile))
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        rowStart++;
                        continue;
                    }

                    var fields = line.Split('\t');

                    // VB logic: only evaluate lines where rowStart > 0 and first field <> "End"
                    if (rowStart > 0 && fields.Length > 0 && !string.Equals(fields[0].Trim(), "End", StringComparison.Ordinal))
                    {
                        var localPart = fields[0].Trim();
                        var groupField = fields.Length > 2 ? fields[2] : string.Empty;

                        if (string.Equals(partNumber, localPart, StringComparison.Ordinal) && string.Equals(groupField, "SMT_Group", StringComparison.Ordinal))
                        {
                            return true;
                        }
                    }

                    rowStart++;
                }
            }
            catch
            {
                // Mirror VB's silent failure behavior: return false on IO/parsing errors.
                return false;
            }
            return false;
        }

        /// <summary>
        /// //Public Function IsPostAssyComponent(getPartNumber As String) As Boolean
        /// </summary>
        /// <param name="partNumber"></param>
        /// <returns></returns>
        private bool IsPCBAComponent(string partNumber) 
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\WareHouseMaterialsSource.txt";

            if (string.IsNullOrEmpty(partNumber))
                return false;

            if (!File.Exists(targetFile))
            {
                Update_WareHouseMaterialsSourceTxt();
            }

            try
            {
                int rowStart = 0;
                foreach (var line in File.ReadLines(targetFile))
                {
                    // Preserve VB behavior: skip processing when rowStart == 0 (header/first line)
                    // and ignore blank lines.
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        var fields = line.Split('\t');

                        if (rowStart > 0 && fields.Length > 0 && !string.Equals(fields[0].Trim(), "End", StringComparison.Ordinal))
                        {
                            var localPart = fields[0].Trim();
                            var groupField = fields.Length > 2 ? fields[2] : string.Empty;

                            if (string.Equals(partNumber, localPart, StringComparison.Ordinal) &&
                                string.Equals(groupField, "POSTASSY_Group", StringComparison.Ordinal))
                            {
                                return true;
                            }
                        }
                    }

                    rowStart++;
                }
            }
            catch
            {
                // Mirror VB's silent-failure style: return false on IO/parsing errors.
                return false;
            }

            return false;
        }

        /// <summary>
        /// //Public Function IsPreAssyModel(getPartNumber As String) As Boolean
        /// </summary>
        /// <param name="partNumber"></param>
        /// <returns></returns>
        private bool IsPCBModel(string partNumber) 
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\PreAssyModel.txt";


            if (string.IsNullOrEmpty(partNumber))
                return false;

            if (!File.Exists(targetFile))
                return false;

            try
            {
                int rowStart = 0;
                foreach (var line in File.ReadLines(targetFile))
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        var fields = line.Split('\t');

                        // VB logic: only evaluate when rowStart > 0 and first field <> "End"
                        if (rowStart > 0 && fields.Length > 0 && !string.Equals(fields[0].Trim(), "End", StringComparison.Ordinal))
                        {
                            var localModel = fields[0].Trim();
                            if (string.Equals(partNumber, localModel, StringComparison.Ordinal))
                            {
                                return true;
                            }
                        }
                    }

                    rowStart++;
                }
            }
            catch
            {
                // Mirror VB's On Error Resume Next: swallow errors and return false.
                return false;
            }

            return false;
        }

        /// <summary>
        /// //Public Function isPostAssyModel(getPartNumber As String) As Boolean; Public Function isPCBAModel(ByVal getPCBAModel As String) As Boolean
        /// </summary>
        /// <param name="partNumber"></param>
        /// <returns></returns>
        private bool IsPCBAModel(string partNumber) 
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\PostAssyModel.txt";
            //string targetFile  = "C:\MPH - KANBAN Control Local Data\PCBAPackagedInf.txt" //we can also use this.

            if (string.IsNullOrEmpty(partNumber))
                return false;

            if (!File.Exists(targetFile))
                return false;

            try
            {
                int rowStart = 0;
                foreach (var line in File.ReadLines(targetFile))
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        var fields = line.Split('\t');

                        // VB logic: only evaluate when rowStart > 0 and first field <> "End"
                        if (rowStart > 0 && fields.Length > 0 && !string.Equals(fields[0].Trim(), "End", StringComparison.Ordinal))
                        {
                            var localModel = fields[0].Trim();
                            if (string.Equals(partNumber, localModel, StringComparison.Ordinal))
                            {
                                return true;
                            }
                        }
                    }

                    rowStart++;
                }
            }
            catch
            {
                // Mirror VB permissive behavior: swallow errors and return false.
                return false;
            }

            return false;
        }

        private void ClearAndSeedGrid(DataGridView grid, int ensureColumns)
        {
            grid.Rows.Clear();
            // Ensure at least ensureColumns columns exist
            while (grid.Columns.Count < ensureColumns)
            {
                grid.Columns.Add("c" + grid.Columns.Count, "c" + grid.Columns.Count);
            }
            grid.Rows.Add(); // seeded empty row (VB AddItem "")
        }

        /// <summary>
        /// return index of row to write into
        /// </summary>
        /// <param name="grid"></param>
        /// <returns></returns>
        private int AddRowIfNeeded(DataGridView grid)
        {
            // Ensures there is at least one empty row at the end and returns index to write into.
            if (grid.Rows.Count == 0)
            {
                grid.Rows.Add();
                return 0;
            }
            // Use last row as writable (VB used Rows - 1)
            int index = grid.Rows.Count - 1;
            // If last row has any non-null content in its cells, append a new row and use that
            bool isEmpty = true;
            foreach (DataGridViewCell cell in grid.Rows[index].Cells)
            {
                if (cell.Value != null && cell.Value.ToString() != "")
                {
                    isEmpty = false;
                    break;
                }
            }
            if (!isEmpty)
            {
                grid.Rows.Add();
                index = grid.Rows.Count - 1;
            }
            return index;
        }

        private void SetGridCell(DataGridView grid, int rowIndex, int colIndex, object value)
        {
            // Ensure column exists
            while (grid.Columns.Count <= colIndex)
            {
                grid.Columns.Add("c" + grid.Columns.Count, "c" + grid.Columns.Count);
            }
            while (grid.Rows.Count <= rowIndex)
            {
                grid.Rows.Add();
            }
            grid.Rows[rowIndex].Cells[colIndex].Value = value;
        }

        private string GetGridCellAsString(DataGridView grid, int rowIndex, int colIndex)
        {
            try
            {
                if (rowIndex >= 0 && rowIndex < grid.Rows.Count && colIndex >= 0 && colIndex < grid.Columns.Count)
                {
                    var v = grid.Rows[rowIndex].Cells[colIndex].Value;
                    return v?.ToString() ?? string.Empty;
                }
            }
            catch { }
            return string.Empty;
        }

        private object GetGridCellAsObject(DataGridView grid, int rowIndex, int colIndex)
        {
            try
            {
                if (rowIndex >= 0 && rowIndex < grid.Rows.Count && colIndex >= 0 && colIndex < grid.Columns.Count)
                {
                    return grid.Rows[rowIndex].Cells[colIndex].Value;
                }
            }
            catch { }
            return null;
        }

        private double ConvertToDoubleSafe(string s)
        {
            double d;
            if (double.TryParse(s, out d)) return d;
            return 0.0;
        }

        /// <summary>
        /// VB6: Public Function AccessTopBottomRunning(getPONumber As String) As String;
        /// </summary>
        /// <param name="poNumber"></param>
        /// <returns></returns>
        private string GetPOTopBot(string poNumber)
        {
            string poSide = "A";
            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                if (GetGridCellAsString(dgvPulledListPO, i, 1).Trim() == poNumber)
                {
                    poSide = GetGridCellAsString(dgvPulledListPO, i, 3).Trim(); //Side
                    break;
                }
            }
            return poSide;
        }

        private bool IsMaterialInvalidForProcessing (string materialNumber)
        {
            if (materialNumber.StartsWith("R", StringComparison.OrdinalIgnoreCase)
                                            || (materialNumber.Length >= 4 && materialNumber.Substring(0, 4).Equals("6100", StringComparison.OrdinalIgnoreCase))
                                            || (materialNumber.Length >= 3 && materialNumber.Substring(0, 3).Equals("621", StringComparison.OrdinalIgnoreCase))
                                            || (materialNumber.Length >= 2 && materialNumber.Substring(0, 2).Equals("DR", StringComparison.OrdinalIgnoreCase))
                                            || (materialNumber.Length >= 7 && materialNumber.Substring(0, 7).Equals("6000000", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }
            return false;
        }



        /// <summary>
        /// VB6: Public Sub LoadLocalBackupPOPhysicalMaterialsPulledList(ByVal getPONumber As String, ByVal getModel As String, ByVal getModelSide As String, ByVal typePulledList As String);
        /// Load the VOM file (ẽ. SMT102166545.txt) to dgvPLPhysicalModelPulled;
        /// </summary>
        /// <param name="poNumber"></param>
        /// <param name="modelNumber"></param>
        /// <param name="poSide"></param>
        /// <param name="pulledListType"></param>
        private void LoadBOMFromSharedDriveTo_dgvPLPhysicalModelPulled(string poNumber, string modelNumber, string poSide, string pulledListType)
        {
            if (string.IsNullOrEmpty(poNumber)) return;

            if (poNumber.Length > 12)
            {
                poNumber = poNumber.Substring(0, 12);
            }

            string targetFile = System.IO.Path.Combine(@"\\vnmsrv305\pubfiles\ME\Truong_ME\Local_Backup_PO_PullList", poNumber + ".txt");
            if (!File.Exists(targetFile))
            {
                // Nothing to load
                MessageBox.Show(targetFile, "File does not exist:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string[] splittedFields;
            string dataLine;
            int rowStart = 0;
            bool isPCBModel = false;
            bool isPCBComponent = false;
            bool isPCBAComponent = false;
            bool isPCBAModel = false; // VB used string here but logic expects boolean
            string strSAPModel = string.Empty;

            try
            {
                using (var sr = new StreamReader(targetFile))
                {
                    // Read all lines and process starting after the header (skip first 4 lines)
                    while ((dataLine = sr.ReadLine()) != null)
                    {
                        // Use fields split by tab
                        splittedFields = dataLine.Split('\t');

                        // We want to skip the first 4 lines (VB logic effectively processed when rowStart > 3)
                        if (rowStart >= 4)
                        {
                            // On the first processed data row, store strSAPModel from column index 1 (if present)
                            if (rowStart == 4)
                            {
                                if (splittedFields.Length > 1)
                                {
                                    strSAPModel = splittedFields[1].Trim();
                                }  
                            }

                            // Ensure we have at least 5 columns to check qty at index 4
                            if (splittedFields.Length > 4)
                            {
                                double qtyVal = Math.Abs(Convert.ToDouble(splittedFields[4])); //Qty
                                qtyVal = Math.Round(qtyVal);
                                if (qtyVal > 0)
                                {
                                    // Handle DicrectPO (original spelling)
                                    if (string.Equals(pulledListType, "DicrectPO", StringComparison.OrdinalIgnoreCase))
                                    {
                                        // Material code is column index 2
                                        string materialNumber = splittedFields.Length > 2 ? splittedFields[2].Trim() : string.Empty;
                                        // Filter out unwanted prefixes
                                        bool skipByPrefix = IsMaterialInvalidForProcessing(materialNumber);   

                                        if (!skipByPrefix)
                                        {
                                            if (!IsCommonMaterials(materialNumber))
                                            {
                                                if (!IsPartRemovedFromBOM(materialNumber))
                                                {
                                                    // tag checks
                                                    isPCBModel = IsPCBAModel(materialNumber);
                                                    isPCBAModel = IsPCBAModel(materialNumber);
                                                    isPCBComponent = IsPCBComponent(materialNumber);
                                                    isPCBAComponent = IsPCBAComponent(materialNumber);

                                                    // Filtering based on PulledListSectorGroup and tag flags
                                                    bool allow =
                                                        (_pulledList_SectorGroup == "POSTASSY_Group" && !isPCBComponent && isPCBAComponent && !isPCBModel && !isPCBAModel)
                                                        || (_pulledList_SectorGroup == "SMT_Group" && isPCBComponent && !isPCBAComponent && !isPCBAModel)
                                                        || (_pulledList_SectorGroup == "SMT_Group" && isPCBComponent && !isPCBAComponent && isPCBAModel)
                                                        || (_pulledList_SectorGroup != "SMT_Group" && _pulledList_SectorGroup != "POSTASSY_Group");

                                                    if (allow)
                                                    {
                                                        //var grid = dgvPLPhysicalModelPulled;
                                                        int rowIndex = AddRowIfNeeded(dgvPLPhysicalModelPulled);

                                                        // VB: .TextMatrix(ii, 0) = ii  -> we use rowIndex as the index value
                                                        SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 0, rowIndex);

                                                        if (_pulledList_SectorGroup == "SMT_Group")
                                                        {
                                                            // column 1 = "SMT" & splittedFields(0)
                                                            SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 1, "SMT" + (splittedFields.Length > 0 ? splittedFields[0].Trim() : string.Empty));
                                                            // column 2 = getModel
                                                            SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 2, modelNumber);
                                                            // column 7 = getModelSide
                                                            SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 7, poSide);
                                                        }
                                                        else
                                                        {
                                                            SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 1, (splittedFields.Length > 0 ? splittedFields[0].Trim() : string.Empty));
                                                            SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 2, (splittedFields.Length > 1 ? splittedFields[1].Trim() : string.Empty));
                                                            SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 7, "A");
                                                        }

                                                        SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 3, materialNumber);
                                                        // column 4 is description with single quotes removed
                                                        string desc = splittedFields.Length > 3 ? splittedFields[3].Trim().Replace("'", "") : string.Empty;
                                                        SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 4, desc);
                                                        SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 5, qtyVal);
                                                        SetGridCell(dgvPLPhysicalModelPulled, rowIndex, 6, (splittedFields.Length > 5 ? splittedFields[5].Trim() : string.Empty));
                                                    }
                                                    else
                                                    {

                                                    }
                                                }
                                            }
                                        }
                                    }
                                    // AfterCO handling
                                    else if (string.Equals(pulledListType, "AfterCO", StringComparison.OrdinalIgnoreCase))
                                    {
                                        string material = splittedFields.Length > 2 ? splittedFields[2].Trim() : string.Empty;
                                        bool skipByPrefix =
                                            material.StartsWith("R", StringComparison.OrdinalIgnoreCase)
                                            || (material.Length >= 4 && material.Substring(0, 4).Equals("6100", StringComparison.OrdinalIgnoreCase))
                                            || (material.Length >= 3 && material.Substring(0, 3).Equals("621", StringComparison.OrdinalIgnoreCase))
                                            || (material.Length >= 2 && material.Substring(0, 2).Equals("DR", StringComparison.OrdinalIgnoreCase))
                                            || (material.Length >= 7 && material.Substring(0, 7).Equals("6000000", StringComparison.OrdinalIgnoreCase));

                                        if (!skipByPrefix)
                                        {
                                            if (!IsPartRemovedFromBOM(material))
                                            {
                                                var grid = dgvPhysicalSAPModelAfterCOPulled;
                                                int rowIndex = AddRowIfNeeded(grid);
                                                SetGridCell(dgvPhysicalSAPModelAfterCOPulled, rowIndex, 0, rowIndex);
                                                SetGridCell(dgvPhysicalSAPModelAfterCOPulled, rowIndex, 1, material);

                                                if (_pulledList_SectorGroup == "SMT_Group")
                                                {
                                                    SetGridCell(dgvPhysicalSAPModelAfterCOPulled, rowIndex, 2, modelNumber);
                                                }
                                                else
                                                {
                                                    string desc = splittedFields.Length > 3 ? splittedFields[3].Trim().Replace("'", "") : string.Empty;
                                                    SetGridCell(dgvPhysicalSAPModelAfterCOPulled, rowIndex, 2, desc);
                                                }

                                                SetGridCell(dgvPhysicalSAPModelAfterCOPulled, rowIndex, 3, Math.Abs(ConvertToDoubleSafe(splittedFields[4])));
                                                SetGridCell(dgvPhysicalSAPModelAfterCOPulled, rowIndex, 4, (splittedFields.Length > 5 ? splittedFields[5].Trim() : string.Empty));
                                            }
                                        }
                                    }
                                    // BeforeCO handling
                                    else if (string.Equals(pulledListType, "BeforeCO", StringComparison.OrdinalIgnoreCase))
                                    {
                                        string material = splittedFields.Length > 2 ? splittedFields[2].Trim() : string.Empty;
                                        bool skipByPrefix =
                                            material.StartsWith("R", StringComparison.OrdinalIgnoreCase)
                                            || (material.Length >= 4 && material.Substring(0, 4).Equals("6100", StringComparison.OrdinalIgnoreCase))
                                            || (material.Length >= 3 && material.Substring(0, 3).Equals("621", StringComparison.OrdinalIgnoreCase))
                                            || (material.Length >= 2 && material.Substring(0, 2).Equals("DR", StringComparison.OrdinalIgnoreCase))
                                            || (material.Length >= 7 && material.Substring(0, 7).Equals("6000000", StringComparison.OrdinalIgnoreCase));

                                        if (!skipByPrefix)
                                        {
                                            if (!IsPartRemovedFromBOM(material))
                                            {
                                                var grid = dgvPhysicalModelRunningPulled;
                                                int rowIndex = AddRowIfNeeded(dgvPhysicalModelRunningPulled);
                                                SetGridCell(dgvPhysicalModelRunningPulled, rowIndex, 0, rowIndex);
                                                SetGridCell(dgvPhysicalModelRunningPulled, rowIndex, 1, material);

                                                if (_pulledList_SectorGroup == "SMT_Group")
                                                {
                                                    SetGridCell(dgvPhysicalModelRunningPulled, rowIndex, 2, modelNumber);
                                                }
                                                else
                                                {
                                                    string desc = splittedFields.Length > 3 ? splittedFields[3].Trim().Replace("'", "") : string.Empty;
                                                    SetGridCell(dgvPhysicalModelRunningPulled, rowIndex, 2, desc);
                                                }

                                                SetGridCell(dgvPhysicalModelRunningPulled, rowIndex, 3, Math.Abs(ConvertToDoubleSafe(splittedFields[4])));
                                                SetGridCell(dgvPhysicalModelRunningPulled, rowIndex, 4, (splittedFields.Length > 5 ? splittedFields[5].Trim() : string.Empty));
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        rowStart++;
                    } // end while
                } // end using sr

                // After processing, call the copy routine:
                CopyBOMTxtFilesFromSharedDriveToLocalFolder();
            }
            catch (Exception ex)
            {
                MessageBox.Show("LoadBOMFromSharedDriveTo_dgvPLPhysicalModelPulled failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// This is part of Public Sub PullPOtoPart(ByVal prodorder As String, getModel As String)
        /// </summary>
        /// <param name="po"></param>
        /// <param name="model"></param>
        private void GetBOMSAPToDataGridViews(string po, string model)
        {
            po = po.Trim();
            string result = "";
            int intPhysicalRow = 0;

            string strDateTime = System.DateTime.Now.ToString("yyyyMMdd-hhmm");
            string SAPModel = "";
            string SAPModelDesc = "";
            string partNumber = "";
            string partDesc = "";
            string motherPart = "";
            string qty = "";
            string UOM = "";
            string xDoc = "";

            //string xmlFileName = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\" + PO + "_" + strDateTime + ".xml";
            //string txtFileName = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\" + PO + "_" + strDateTime + ".txt";
            //string txtFileName2 = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\" + PO + "_BOMPulled_" + strDateTime + ".txt";

            sapLink.clsSAPLink sap = new sapLink.clsSAPLink();
            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList oNodes;

            // Get the XML document of the BOM
            xDoc = "<?xml version=\"1.0\"?><ZRFC_SEND_PODATA_ACS><AUFNR>" + po + "</AUFNR></ZRFC_SEND_PODATA_ACS>";

            try
            {
                result = sap.postMessageNew(xDoc, "home", 10, "PRD");
                xmlDoc.LoadXml(result);
                //xmlDoc.Save(xmlFileName);
                //xmlDoc.GetElementsByTagName("RETURN_CODE");

                // Create txt File:
                //if (!File.Exists(txtFileName))
                //{
                //    File.Create(txtFileName).Close();
                //}

                // Get Top Model Info:
                SAPModel = xmlDoc.GetElementsByTagName("MATERIAL").Item(0).InnerText.Trim().ToUpper();
                SAPModelDesc = xmlDoc.GetElementsByTagName("DESCRIPTION").Item(0).InnerText.Trim().ToUpper();
                oNodes = xmlDoc.SelectNodes("//PODATA_ACS/*"); //list of XML nodes that are the BOM items
                //using (StreamWriter sw = File.AppendText(txtFileName))
                //{
                //    sw.WriteLine("Product Name: " + "\t" + SAPModel + "\t" + SAPModelDesc);
                //    sw.WriteLine("PO Number" + "\t" + "Product Model" + "\t" + "Materials Number" + "\t" + "Material Description" + "\t" + "Qty/PO" + "\t" + "UOM");//write header
                //}

                //Components:
                int j = 0;

                foreach (XmlNode xn in oNodes) //for each BOM item
                {
                    partNumber = xn.SelectSingleNode("IDNRK").InnerText.Trim().ToUpper();
                    partDesc = xn.SelectSingleNode("MAKTX").InnerText.Trim().ToUpper();
                    motherPart = xn.SelectSingleNode("MATNR").InnerText.Trim().ToUpper();
                    qty = xn.SelectSingleNode("MENGE").InnerText.Trim().ToUpper();
                    UOM = xn.SelectSingleNode("MEINS").InnerText.Trim().ToUpper();

                    ////write to txt #1:
                    //using (StreamWriter sw = File.AppendText(txtFileName))
                    //{
                    //    sw.WriteLine(PO + "\t" + motherPart + "\t" + partNumber + "\t" + partDesc + "\t" + qty + "\t" + UOM);
                    //}

                    // Populate into datagridviews:
                    if (Convert.ToDouble(qty) > 0)
                    {
                        // Filter materials:
                        string pnTrim = partNumber?.Trim() ?? string.Empty;
                        bool startsWithR = pnTrim.Length >= 1 && pnTrim.Substring(0, 1).Equals("R", StringComparison.OrdinalIgnoreCase);
                        bool startsWith6100 = pnTrim.Length >= 4 && pnTrim.Substring(0, 4).Equals("6100", StringComparison.OrdinalIgnoreCase);
                        bool startsWithDR = pnTrim.Length >= 2 && pnTrim.Substring(0, 2).Equals("DR", StringComparison.OrdinalIgnoreCase);
                        bool startsWith6000000 = pnTrim.Length >= 7 && pnTrim.Substring(0, 7).Equals("6000000", StringComparison.OrdinalIgnoreCase);

                        if (!startsWithR && !startsWith6100 && !startsWithDR && !startsWith6000000)
                        {
                            if (!IsCommonMaterials(pnTrim) && Convert.ToInt32(qty) > 0)
                            {
                                if (!IsPartRemovedFromBOM(pnTrim))
                                {
                                    j++;
                                    dgvPLOverallModelPulled.Rows.Add(j, pnTrim, partDesc, qty ?? string.Empty, UOM, motherPart);
                                }
                            }
                        }
                    }
                    // Remove the mother level from BOM to get physical materials:
                    for (int i = 0; i < dgvPLOverallModelPulled.Rows.Count; i++)
                    {
                        int k = 0;
                        for (int jj = 0; jj < dgvPLOverallModelPulled.Rows.Count; jj++)
                        {
                            if (dgvPLOverallModelPulled.Rows[i].Cells[1].Value == dgvPLOverallModelPulled.Rows[jj].Cells[5].Value)
                            {
                                k++;
                                break;
                            }
                        }
                        if (k == 0)
                        {
                            intPhysicalRow = AddRowIfNeeded(dgvPLPhysicalModelPulled);
                            dgvPLPhysicalModelPulled.Rows.Add(intPhysicalRow,
                                po,
                                model,
                                dgvPLOverallModelPulled.Rows[i].Cells[1].Value,
                                dgvPLOverallModelPulled.Rows[i].Cells[2].Value,
                                dgvPLOverallModelPulled.Rows[i].Cells[3].Value,
                                dgvPLOverallModelPulled.Rows[i].Cells[4].Value
                                );
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting BOM from SAP for PO " + po + ": " + ex.Message, "SAP Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// VB6: Public Sub SaveToLocalPOPhyicalMaterialsPullList(getPONumber As String, getModel As String, typePulledList As String);
        /// </summary>
        /// <param name="poNumber"></param>
        /// <param name="modelNumber"></param>
        /// <param name="pulledListType"></param>
        private void SaveToLocalPOPhyicalMaterialsPullList(string poNumber, string modelNumber, string pulledListType) //VB6: Public Sub SaveToLocalPOPhyicalMaterialsPullList(getPONumber As String, getModel As String, typePulledList As String)
        {
            string TargetFile = @"\\vnmsrv300\pubfiles\ME\Truong_ME\Local_Backup_PO_PullList\" + poNumber + ".txt";

            CopyBOMTxtFilesFromSharedDriveToLocalFolder(); 
        }

        private void Update_PartRemovedFromBOMTxt()
        {
            MSSQL _sql = new MSSQL();
            string _connection_string = _sql.cnnDLVNDB;
            SqlConnection connection = new SqlConnection(_connection_string);

            string targetFile = @"C:\MPH - KANBAN Control Local Data\PartRemovedFromBOM.txt";

            if (connection == null) throw new ArgumentNullException(nameof(connection));

            // Ensure directory exists
            string dir = System.IO.Path.GetDirectoryName(targetFile);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            // Ensure file exists; we'll overwrite the top (header) then append rows
            if (!File.Exists(targetFile))
            {
                // Create an empty file
                using (File.Create(targetFile)) { }
            }

            // Write header (overwrite file)
            File.WriteAllText(targetFile, "Materials\tSector" + Environment.NewLine, Encoding.UTF8);

            // Prepare and run query
            bool openedHere = false;
            try
            {
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                    openedHere = true;
                }

                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM PartRemovedFromBOM";
                    cmd.CommandType = CommandType.Text;

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader != null)
                        {
                            var sb = new StringBuilder();
                            while (reader.Read())
                            {
                                string materials = reader.IsDBNull(0) ? string.Empty : reader.GetValue(0).ToString().Trim();
                                string sector = reader.IsDBNull(1) ? string.Empty : reader.GetValue(1).ToString().Trim();
                                sb.AppendLine(materials + "\t" + sector);
                            }

                            if (sb.Length > 0)
                            {
                                // Append all collected data rows
                                File.AppendAllText(targetFile, sb.ToString(), Encoding.UTF8);
                            }
                        }
                    }
                }

                // Append final "End\tEnd" (no trailing newline to mirror original ts.Write)
                File.AppendAllText(targetFile, "End\tEnd", Encoding.UTF8);
            }
            finally
            {
                if (openedHere && connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
        }

        /// <summary>
        /// VB6: Public Function IsCommonMaterials(getMaterialsNumber As String) As Boolean
        /// </summary>
        /// <param name="part_number"></param>
        /// <returns></returns>
        private bool IsCommonMaterials(string part_number)
        {
            for (int i = 0; i < dgvPullListvsPO2.Rows.Count; i++)
            {
                var cellVal = dgvPullListvsPO2.Rows[i].Cells["Model"].Value?.ToString();
                if (string.Equals(part_number, cellVal, StringComparison.OrdinalIgnoreCase))
                {
                    return true; // part is common
                }
            }
            return false;
        }

        /// <summary>
        /// VB6: Public Function IsPartRemovedFromBOM(getMaterials As String) As Boolean;
        /// </summary>
        /// <param name="part_number"></param>
        /// <returns></returns>
        private bool IsPartRemovedFromBOM(string part_number)
        {
            string TargetFile = @"C:\MPH - KANBAN Control Local Data\PartRemovedFromBOM.txt";


            if (!System.IO.File.Exists(TargetFile))
            {
                Update_PartRemovedFromBOMTxt();
                //return false; //should we?
            }

            // Read file line by line, skip the header if present (the VB code started rowStart = 1 then incremented before processing lines,
            // but it effectively checked all lines; we'll check all lines and compare the first tab-separated field).
            foreach (var line in File.ReadLines(TargetFile))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;

                var fields = line.Split('\t');
                if (fields.Length == 0) continue;

                var localMaterials = fields[0].Trim();
                if (part_number.Trim() == localMaterials)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// VB6: Private Sub GetMaxQtyDuplicated(getPONumber As String);
        /// </summary>
        /// <param name="po"></param>
        private void GetMaxQtyDuplicated(string po)
        {
            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                //dgvPulledListPO - PONumber = SMT102160895B_B >> po = 102160895;
                if (Convert.ToString(dgvPulledListPO.Rows[i].Cells["PONumber"].Value).Trim() == po)
                {
                    //dgvQtyvsCountDuplicated.Rows[1].Cells["Qty"].Value = Convert.ToInt32(dgvPulledListPO.Rows[i].Cells["Qty"].Value);
                    SetGridCell(dgvQtyvsCountDuplicated, 1, 2, GetGridCellAsObject(dgvPulledListPO, i, 4));
                    break;
                }
            }
        }

        /// <summary>
        /// VB6: Public Function IsPOExistInLocal(ByVal getPONumber As String) As Boolean;
        /// Check if PO txt file (ex.SMT102099984.txt) already exists in local folder "C:\MPH - KANBAN Control Local Data\LocalBOMToPulledList\
        /// </summary>
        /// <param name="po"></param>
        /// <returns></returns>
        private bool IsPOBOMTxtAlreadyInLocalFolder(string po)
        {
            if (po.Substring(0,3) == "SMT")
            {
                po = po.Substring(0, po.Length - 3); //Ex: SMT102099984A_G >> Get SMT102099984
            }
            string TargetFile = @"C:\MPH - KANBAN Control Local Data\LocalBOMToPulledList\" + po + ".txt";
            if (System.IO.File.Exists(TargetFile))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void RC_Add1PO_Click(object sender, EventArgs e)
        {
            // Lock selection controls to avoid changes during processing:
            LockSelection();

            // If Line chosen is SMT >> modify PO Number to match SMT format:
            if (cbbPulledListLine.Text.Trim().Length < 3)
            {
                MessageBox.Show("Please select a valid Line before adding PO Numbers!", "Line Not Selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Input POs - single or from excel file:
            string result = "ERROR";

            // Variables declaration:
            string strImportedPO = string.Empty;

            // Initiate input form:
            frmInputForm inputForm = new frmInputForm();
            inputForm.SetForm("Add Single PO", "Enter PO Number:");
            if (inputForm.ShowDialog() == DialogResult.OK)
            {
                strImportedPO = inputForm.InputText.Trim();
                inputForm.Close();
            }
            else
            {
                inputForm.Close();
                result = "Cancelled input PO!";
                MessageBox.Show(result, "Input PO Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                UnlockSelection();
                return;
            }

            // Process the input PO number:
            strImportedPO = strImportedPO.Replace(((char)160).ToString(), "").Trim();

            result = AddSinglePO(strImportedPO);

            if (result != "OK")
            {
                UnlockSelection();
                return;
            }

            // Release selection controls:
            UnlockSelection();
        }

        private void LockSelection()
        {
            cbbPulledListLine.Enabled = false;
            cbbActiveDate.Enabled = false;
            cbbPulledListShift.Enabled = false;
            RC_Add1PO.Enabled = false;
            RC_AddPOsKitting.Enabled = false;
            RC_Clear.Enabled = false;
        }

        private void UnlockSelection()
        {
            cbbPulledListLine.Enabled = true;
            cbbActiveDate.Enabled = true;
            cbbPulledListShift.Enabled = true;
            RC_Add1PO.Enabled = true;
            RC_AddPOsKitting.Enabled = true;
            RC_Clear.Enabled = true;
        }

        /// <summary>
        /// VB6: Private Sub mAddPO_Click();
        /// </summary>
        /// <returns></returns>
        private string AddSinglePO(string addedPOwithSide)
        {
            string pulledPO = "No PO";
            string result = "OK";
            string addedPO = "No PO";
            string topBot = "A"; //default to A
            string poSector = "No sector";
            string poModel = "No model";
            int poQty = -1;

            //Ex: addedPOwithSide = "102099984A" >> addedPO = "102099984"
            //if (addedPOwithSide.Length == 10)
            //{
            //    addedPO = addedPOwithSide.Substring(0, 9);
            //    topBot = addedPOwithSide.Substring(8, 1).ToUpper(); //A/T/B
            //}

            if (cbbPulledListLine.Text.Trim().Substring(0, 3).ToUpper() == "SMT")
            {
                //Ex: pulled_po = "SMTA102099984A_G"
                pulledPO = "SMT" + addedPOwithSide + "_" + cbbPulledListLine.Text.Trim().ToUpper().Substring(3, 1); //cbbLine must have format SMTA_Line >> get "A";
                if (IsValidPO(pulledPO))
                {
                    if (IsPOUnique(pulledPO))
                    {
                        poSector = GetPOSector(pulledPO);
                        if (poSector == cbbPulledListLine.Text.Trim() || cbbPulledListLine.Text.Trim().ToUpper().Substring(0, 3) == "SMT")
                        {
                            poModel = GetPOModelFromOpenPOPlanner(pulledPO);
                            if (poModel != "NA")
                            {
                                //dgvPulledListPO: No(0); PONumber(1); ModelNumber(2); Side(3); POQty(4); PulledListID(5); PlannersNotice(6); POChangeInf(7)
                                int po_row_index = AddRowIfNeeded(dgvPulledListPO);
                                SetGridCell(dgvPulledListPO, po_row_index, 1, pulledPO);
                                SetGridCell(dgvPulledListPO, po_row_index, 2, poModel);

                                if (cbbPulledListLine.Text.Trim().Substring(0, 3).ToUpper() == "SMT")
                                {
                                    // Get TOP/BOT info: A/T/B
                                    SetGridCell(dgvPulledListPO, po_row_index, 3, topBot);

                                    // Get PO Quantity:
                                    poQty = GetPOQty(pulledPO); // changed to pulledPO
                                    if (poQty == 0)
                                    {
                                        MessageBox.Show("PO Number: " + pulledPO + " has zero quantity!", "Zero PO Quantity", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    SetGridCell(dgvPulledListPO, po_row_index, 4, poQty);

                                    // Check if PO has change quantity info:
                                    if (IsPOQtyChanged(pulledPO)) // changed to pulledPO
                                    {
                                        SetGridCell(dgvPulledListPO, po_row_index, 7, "YES");
                                    }
                                    else
                                    {
                                        SetGridCell(dgvPulledListPO, po_row_index, 7, "NO");
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Not support none SMT PO anymore", "Error:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show("PO Number: " + addedPO + " does not exist in table OpenPOPlanner!", "PO Not Found:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return "PO not in OpenPOPlanner";
                            }
                        }
                        else
                        {
                            MessageBox.Show("PO Number: " + addedPO + " does not belong to sector " + cbbPulledListLine.Text.Trim() + "!", "Sector Mismatch", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return "PO not in sector";
                        }
                    }
                    else
                    {
                        MessageBox.Show("PO Number: " + addedPO + " is already in the list!", "Duplicate PO Number", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return "PO in list already";
                    }
                }
                else
                {
                    MessageBox.Show("PO Number: " + addedPO + "\r\nis invalid format!", "Invalid PO Number", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return "Invalid PO Format";
                }

                //Check if all models have layout setup:
            }
            return result;
        }


        private void btnShow_Click(object sender, EventArgs e)
        {
            if (btnShow.Text == "Hide") //Hide rows 3-9
            {
                TableLayoutPanel tlp = this.tbllayout_bot;
                tlp.SuspendLayout();
                //// Save previous RowStyle if you need to restore it later (optional)
                //RowStyle prev = tlp.RowStyles[rowIndex];

                // Set row to absolute 0 height
                for (int rowIndex = 3; rowIndex <= 10; rowIndex++)
                {
                    tlp.RowStyles[rowIndex].SizeType = SizeType.Absolute;
                    tlp.RowStyles[rowIndex].Height = 0F;
                    // Hide every control assigned to that row
                    foreach (System.Windows.Forms.Control c in tlp.Controls)
                    {
                        if (tlp.GetRow(c) == rowIndex)
                            c.Visible = false;
                    }
                }
                tlp.ResumeLayout();
                tlp.PerformLayout();
                btnShow.Text = "Show";
            }
            else //"Show"
            {
                TableLayoutPanel tlp = this.tbllayout_bot;
                tlp.SuspendLayout();

                // Restore desired height and size type
                tlp.RowStyles[3].SizeType = SizeType.Absolute;
                tlp.RowStyles[3].Height = 20F;
                tlp.RowStyles[4].SizeType = SizeType.Percent;
                tlp.RowStyles[4].Height = 20F;
                tlp.RowStyles[5].SizeType = SizeType.Absolute;
                tlp.RowStyles[5].Height = 25F;
                tlp.RowStyles[6].SizeType = SizeType.Percent;
                tlp.RowStyles[6].Height = 20F;
                tlp.RowStyles[7].SizeType = SizeType.Absolute;
                tlp.RowStyles[7].Height = 25F;
                tlp.RowStyles[8].SizeType = SizeType.Percent;
                tlp.RowStyles[8].Height = 20F;
                tlp.RowStyles[9].SizeType = SizeType.Absolute;
                tlp.RowStyles[9].Height = 25F;
                tlp.RowStyles[10].SizeType = SizeType.Percent;
                tlp.RowStyles[10].Height = 25F;

                // Show controls in that row
                for (int rowIndex = 3; rowIndex <= 10; rowIndex++)
                {
                    foreach (System.Windows.Forms.Control c in tlp.Controls)
                    {
                        if (tlp.GetRow(c) == rowIndex)
                            c.Visible = true;
                    }
                }
                tlp.ResumeLayout();
                tlp.PerformLayout();
                btnShow.Text = "Hide";
            }
        }

        /// <summary>
        /// Update C:\MPH - KANBAN Control Local Data\MPHMROList.txt
        /// </summary>
        private void Update_MPHMROListTxt()
        {
            //string targetFile = @"C:\MPH - KANBAN Control Local Data\MPHMROList.txt";
            const string dir = @"C:\MPH - KANBAN Control Local Data";
            string targetFile = System.IO.Path.Combine(dir, "MPHMROList.txt");

            // Ensure directory exists
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            // Ensure file exists
            if (!System.IO.File.Exists(targetFile))
            {
                System.IO.File.Create(targetFile).Close();
            }

            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            try
            {
                // Write header line (overwrite existing file)
                string header = "MROMaterials\tMtDes\tUOM\tType\tIsBOX";
                File.WriteAllText(targetFile, header + Environment.NewLine);

                // Open StreamWriter for appending the data rows
                using (var sw = new StreamWriter(targetFile, append: true))
                {
                    using (var cmd = new SqlCommand("SELECT * FROM MROlst", cnnDLVNDB))
                    {
                        cmd.CommandType = CommandType.Text;
                        bool openedHere = false;

                        try
                        {
                            if (cnnDLVNDB == null)
                                throw new InvalidOperationException("Database connection (cnnDLVNDB) is null.");

                            if (cnnDLVNDB.State != ConnectionState.Open)
                            {
                                cnnDLVNDB.Open();
                                openedHere = true;
                            }

                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader != null)
                                {
                                    while (reader.Read())
                                    {
                                        string pn = reader["PN"]?.ToString().Trim() ?? string.Empty;
                                        string mtDes = reader["MtDes"]?.ToString().Trim() ?? string.Empty;
                                        string uom = reader["UOM"]?.ToString().Trim() ?? string.Empty;
                                        string box = reader["Box"]?.ToString().Trim() ?? string.Empty;

                                        // VB wrote: PN \t MtDes \t UOM \t "MAT" \t Box
                                        string record = string.Join("\t", pn, mtDes, uom, "MAT", box);
                                        sw.WriteLine(record);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
                            {
                                try { cnnDLVNDB.Close(); } catch { /* swallow to match original VB behavior */ }
                            }
                        }
                    }

                    // Write final End marker without an extra newline (matches original VB ts.Write)
                    sw.Write("End\tEnd\tEnd\tEnd\tEnd");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("UpdateLocalMROList failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Update C:\MPH - KANBAN Control Local Data\MPHAllSector.txt
        /// </summary>
        private void Update_MPHAllSectorTxt() //VB6: Public Sub UpdateLocalAllSector()
        {
            //string targetFile = @"C:\MPH - KANBAN Control Local Data\MPHAllSector.txt";
            string dir = @"C:\MPH - KANBAN Control Local Data";
            string targetFile = System.IO.Path.Combine(dir, @"MPHAllSector.txt");

            // Ensure directory exists
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            // Ensure file exists
            if (!System.IO.File.Exists(targetFile))
            {
                System.IO.File.Create(targetFile).Close();
            }

            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            try
            {
                // Write header line (overwrite existing file)
                string header = "Series\tModel\tThresholdFG\tSector\tPSA\tSectorType";
                File.WriteAllText(targetFile, header + Environment.NewLine);

                // Open StreamWriter for appending the data rows
                using (var sw = new StreamWriter(targetFile, append: true))
                {
                    using (var cmd = new SqlCommand("SELECT * FROM KANBANModelSeries", cnnDLVNDB))
                    {
                        cmd.CommandType = CommandType.Text;
                        bool openedHere = false;

                        try
                        {
                            if (cnnDLVNDB == null)
                                throw new InvalidOperationException("Database connection (cnnDLVNDB) is null.");

                            if (cnnDLVNDB.State != ConnectionState.Open)
                            {
                                cnnDLVNDB.Open();
                                openedHere = true;
                            }

                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader != null)
                                {
                                    while (reader.Read())
                                    {
                                        string series = reader["Series"]?.ToString().Trim() ?? string.Empty;
                                        string model = reader["Model"]?.ToString().Trim() ?? string.Empty;
                                        string thresholdFG = reader["ThresholdFG"]?.ToString() ?? string.Empty;
                                        string sect = reader["Sect"]?.ToString().Trim() ?? string.Empty;
                                        string psa = reader["PSA"]?.ToString().Trim() ?? string.Empty;
                                        string sectorType = reader["SectorType"]?.ToString().Trim() ?? string.Empty;

                                        string record = string.Join("\t", series, model, thresholdFG, sect, psa, sectorType);
                                        sw.WriteLine(record);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
                            {
                                try { cnnDLVNDB.Close(); } catch { /* swallow to match VB behavior */ }
                            }
                        }
                    }

                    // Write final End marker without an extra newline (matches original VB ts.Write)
                    sw.Write("End\tEnd\tEnd\tEnd\tEnd\tEnd");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("UpdateLocalAllSector failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// VB6:Public Sub UpdateLocalPCBABin()
        /// Update C:\MPH - KANBAN Control Local Data\MPHKANBANFixLoc.txt
        /// </summary>
        private void Update_MPHKANBANFixLocTxt() 
        {
            const string dir = @"C:\MPH - KANBAN Control Local Data";
            string targetFile = System.IO.Path.Combine(dir, "MPHKANBANFixLoc.txt");

            // Ensure directory exists
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            // Ensure file exists
            if (!System.IO.File.Exists(targetFile))
            {
                System.IO.File.Create(targetFile).Close();
            }

            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            try
            {
                // Ensure file exists and write the header (VB used ForWriting first to overwrite)
                File.WriteAllText(targetFile, "TypeBoxConstraint\tInventoryLocation\tPCBAPartName" + Environment.NewLine);

                // Open file for appending and write records from DB
                using (var sw = new StreamWriter(targetFile, append: true))
                {
                    // Query DB
                    using (var cmd = new SqlCommand("SELECT * FROM WareHouseMaterialsInventoryCurrentStock", cnnDLVNDB))
                    {
                        cmd.CommandType = CommandType.Text;

                        bool openedHere = false;
                        try
                        {
                            if (cnnDLVNDB == null)
                                throw new InvalidOperationException("Database connection (cnnDLVNDB) is not initialized.");

                            if (cnnDLVNDB.State != ConnectionState.Open)
                            {
                                cnnDLVNDB.Open();
                                openedHere = true;
                            }

                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader != null)
                                {
                                    while (reader.Read())
                                    {
                                        string inventoryLocation = reader["InventoryLocation"]?.ToString().Trim() ?? string.Empty;
                                        string materials = reader["Materials"]?.ToString().Trim() ?? string.Empty;

                                        string recordSetUpData = $"Big Box\t{inventoryLocation}\t{materials}";
                                        sw.WriteLine(recordSetUpData);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
                            {
                                try { cnnDLVNDB.Close(); } catch { /* swallow to match original VB behavior */ }
                            }
                        }
                    }

                    // Write final End line without extra newline (VB used ts.Write)
                    sw.Write("End\tEnd\tEnd");
                }
            }
            catch (Exception ex)
            {
                // Mirror VB style (no thrown error) but report in UI; adapt to your logging if desired.
                MessageBox.Show("UpdateLocalPCBABin failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //private void Update_PartRemovedFromBOMTxt() //VB6: Public Sub UpdatePartRemovedFromBOM()
        //{
        //    const string dir  = @"C:\MPH - KANBAN Control Local Data";
        //    string targetFile = System.IO.Path.Combine(dir, "PartRemovedFromBOM.txt");

        //    // Ensure directory exists
        //    if (!System.IO.Directory.Exists(dir))
        //    {
        //        System.IO.Directory.CreateDirectory(dir);
        //    }

        //    // Ensure file exists
        //    if (!System.IO.File.Exists(targetFile))
        //    {
        //        System.IO.File.Create(targetFile).Close();
        //    }

        //    MSSQL _sql = new MSSQL();
        //    SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

        //    try
        //    {
        //        // Write header line (overwrite existing file)
        //        string header = "Materials\tSector";
        //        File.WriteAllText(targetFile, header + Environment.NewLine);

        //        // Append data rows from DB
        //        using (var sw = new StreamWriter(targetFile, append: true))
        //        using (var cmd = new SqlCommand("SELECT * FROM PartRemovedFromBOM", cnnDLVNDB))
        //        {
        //            cmd.CommandType = CommandType.Text;
        //            bool openedHere = false;

        //            try
        //            {
        //                if (cnnDLVNDB == null)
        //                    throw new InvalidOperationException("Database connection (cnnDLVNDB) is null.");

        //                if (cnnDLVNDB.State != ConnectionState.Open)
        //                {
        //                    cnnDLVNDB.Open();
        //                    openedHere = true;
        //                }

        //                using (var reader = cmd.ExecuteReader())
        //                {
        //                    if (reader != null)
        //                    {
        //                        while (reader.Read())
        //                        {
        //                            string materials = reader["Materials"]?.ToString().Trim() ?? string.Empty;
        //                            string sector = reader["Sector"]?.ToString().Trim() ?? string.Empty;

        //                            string record = string.Join("\t", materials, sector);
        //                            sw.WriteLine(record);
        //                        }
        //                    }
        //                }
        //            }
        //            finally
        //            {
        //                if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
        //                {
        //                    try { cnnDLVNDB.Close(); } catch { /* swallow to match VB behavior */ }
        //                }
        //            }

        //            // Write final End marker without extra newline (VB used ts.Write)
        //            sw.Write("End\tEnd");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("UpdatePartRemovedFromBOM failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        /// <summary>
        /// Update C:\MPH - KANBAN Control Local Data\PhantomSubMaterialsvsModel.txt
        /// </summary>
        private void Update_PhantomSubMaterialsvsModelTxt()//Vb6: Public Sub UpdateLocalPhantomSubMaterialsvsModel()
        {
            const string dir = @"C:\MPH - KANBAN Control Local Data";
            string targetFile = System.IO.Path.Combine(dir, "PhantomSubMaterialsvsModel.txt");

            // Ensure directory exists
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            // Ensure file exists
            if (!System.IO.File.Exists(targetFile))
            {
                System.IO.File.Create(targetFile).Close();
            }

            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            try
            {

                // Overwrite file with header line (equivalent to ForWriting + WriteLine header in VB)
                string header = "ProductModel\tSubModel\tMaterials\tMaterialsDesc\tQtyMaterialsPerProduct";
                File.WriteAllText(targetFile, header + Environment.NewLine);

                // Append rows from DB
                using (var sw = new StreamWriter(targetFile, append: true))
                using (var cmd = new SqlCommand("SELECT * FROM [DLVNDB].[dbo].[PhantomSubMaterialsvsModel]", cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;
                    bool openedHere = false;

                    try
                    {
                        if (cnnDLVNDB == null)
                            throw new InvalidOperationException("Database connection (cnnDLVNDB) is null.");

                        if (cnnDLVNDB.State != ConnectionState.Open)
                        {
                            cnnDLVNDB.Open();
                            openedHere = true;
                        }

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader != null)
                            {
                                while (reader.Read())
                                {
                                    string productModel = reader["ProductModel"]?.ToString().Trim() ?? string.Empty;
                                    string subModel = reader["SubModel"]?.ToString().Trim() ?? string.Empty;
                                    string materials = reader["Materials"]?.ToString().Trim() ?? string.Empty;
                                    string materialsDesc = reader["MaterialsDesc"]?.ToString().Trim() ?? string.Empty;
                                    string qtyPerProduct = reader["QtyMaterialsPerProduct"]?.ToString().Trim() ?? string.Empty;

                                    string recordSetUpData = string.Join("\t", productModel, subModel, materials, materialsDesc, qtyPerProduct);

                                    if (recordSetUpData.Length >= 9)
                                    {
                                        sw.WriteLine(recordSetUpData);
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
                        {
                            try { cnnDLVNDB.Close(); } catch { /* swallow to match VB behaviour */ }
                        }
                    }

                    // Write final End marker without extra newline (VB used ts.Write)
                    sw.Write("End\tEnd\tEnd\tEnd\tEnd");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("UpdateLocalPhantomSubMaterialsvsModel failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Vb6 equivalent: Public Sub UpdateLocalTrackingCurrentPOStatus()
        /// Update C:\MPH - KANBAN Control Local Data\SectorGeneralInfor.txt
        /// </summary>
        private void Update_SectorGeneralInforTxt()
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            string dir = @"C:\MPH - KANBAN Control Local Data";
            string targetFile = System.IO.Path.Combine(dir, "SectorGeneralInfor.txt");

            // Ensure directory exists
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            // Ensure file exists
            if (!System.IO.File.Exists(targetFile))
            {
                System.IO.File.Create(targetFile).Close();
            }

            string header = "Sector\tSectorType\tMaxCapacity\tPSA\tKITTINGActive\tSectorGroup\tSectorPlanning\tMaterialsPulledListType";

            try
            {
                // Write header (overwrite)
                File.WriteAllText(targetFile, header + Environment.NewLine, Encoding.UTF8);

                // Prepare DB command and reader
                if (cnnDLVNDB.State != ConnectionState.Open)
                {
                    cnnDLVNDB.Open();
                }

                using (IDbCommand cmd = cnnDLVNDB.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM TrackingCurrentPOStatus";
                    cmd.CommandType = CommandType.Text;

                    using (IDataReader reader = cmd.ExecuteReader())
                    {
                        // Append rows from the data reader
                        using (var sw = new StreamWriter(targetFile, append: true, encoding: Encoding.UTF8))
                        {
                            if (reader != null && reader.FieldCount > 0)
                            {
                                while (reader.Read())
                                {
                                    string record = string.Join("\t",
                                        SafeTrim(reader, "Sector"),
                                        SafeTrim(reader, "SectorType"),
                                        SafeTrim(reader, "MaxCapacity"),
                                        SafeTrim(reader, "PSA"),
                                        SafeTrim(reader, "KITTINGActive"),
                                        SafeTrim(reader, "SectorGroup"),
                                        SafeTrim(reader, "SectorPlanning"),
                                        SafeTrim(reader, "MaterialsPulledListType")
                                    );

                                    sw.WriteLine(record);
                                }
                            }

                            // Write End row (no newline in VB used Write, but here we append it consistently)
                            string endRow = "End\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd\tEnd";
                            sw.Write(endRow);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Mirror VB6 behavior of surfacing an error; include exception text for easier diagnosis
                throw new ApplicationException("Error updating local tracking current PO status: " + ex.Message, ex);
            }
            finally
            {
                // do not close connection if caller expects to reuse; optionally close if we opened it here.
                // (In this implementation we don't explicitly close the connection to avoid side-effects)
            }
        }

        // Helper to safely get a trimmed string from reader by column name, handling DBNull.
        private static string SafeTrim(IDataRecord reader, string columnName)
        {
            try
            {
                int ordinal = -1;
                try { ordinal = reader.GetOrdinal(columnName); } catch { return string.Empty; }
                object value = reader.GetValue(ordinal);
                return value == null || value == DBNull.Value ? string.Empty : value.ToString().Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string NormalizeSheetName(string s, int maxLength = 31)
        {
            if (string.IsNullOrEmpty(s)) return "Sheet1";
            foreach (var c in System.IO.Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');
            if (s.Length > maxLength) s = s.Substring(0, maxLength);
            return s;
        }

        private static string TrimSafe(string s) => s?.Trim() ?? string.Empty;

        private string GetPulledListLogID()
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            bool mustClose = false;
            if (cnnDLVNDB.State != ConnectionState.Open)
            {
                cnnDLVNDB.Open();
                mustClose = true;
            }

            try
            {
                using (var cmd = new SqlCommand("dlsvn_PulledListLogID", cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    var outParam = new SqlParameter("@result", SqlDbType.NVarChar, 20)
                    {
                        Direction = ParameterDirection.Output
                    };
                    cmd.Parameters.Add(outParam);

                    cmd.ExecuteNonQuery();

                    var value = outParam.Value;
                    return (value == DBNull.Value || value == null) ? null : value.ToString();
                }
            }
            finally
            {
                if (mustClose) cnnDLVNDB.Close();
            }
        }

        private void CheckPulledListPotentialIssues() //VB6: Public Sub PotentialIssuePulledList()
        {
            dgvPotentialIssues.Rows.Clear();
            string poNumber;
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);
            string query = "SELECT TOP 3 * FROM [DLVNDB].[dbo].[PlanningPotentialPending] where PotentialPONumber = @poNumber";

            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                poNumber = GetGridCellAsString(dgvPulledListPO, i, 1); // PONumber column
                if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));
                
                using (var cmd = new SqlCommand(query, cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.AddWithValue("@poNumber", poNumber);

                    bool openedHere = false;
                    try
                    {
                        if (cnnDLVNDB == null)
                            throw new InvalidOperationException("Database connection (cnnDLVNDB) is not initialized.");

                        if (cnnDLVNDB.State != ConnectionState.Open)
                        {
                            cnnDLVNDB.Open();
                            openedHere = true;
                        }

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader != null)
                            {
                                while (reader.Read())
                                {
                                    int intRow = AddRowIfNeeded(dgvPotentialIssues);
                                    string potentialPONumber = reader["PotentialPONumber"]?.ToString().Trim() ?? string.Empty;
                                    string potentialPendingIssue = reader["PotentialPendingIssue"]?.ToString().Trim() ?? string.Empty;  
                                    SetGridCell(dgvPotentialIssues, intRow, 0, intRow);
                                    SetGridCell(dgvPotentialIssues, intRow, 1, potentialPONumber);
                                    SetGridCell(dgvPotentialIssues, intRow, 2, potentialPendingIssue);
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
                        {
                            try { cnnDLVNDB.Close(); } catch { /* swallow to match original VB behavior */ }
                        }
                    }
                }

            }
        }

        /// <summary>
        /// Vb6 equivalent: Public Function MaterialsAftConversionProgram(ByVal strMaterialsBefProgramming As String, ByVal strPCBModel As String) As String
        /// Find the converted material after programming based on material before programming and PCB model from WareHouseMaterialsProgrammingControl.txt
        /// </summary>
        /// <param name="inputMaterialBeforeProgramming"></param>
        /// <param name="inputPCBModel"></param>
        /// <returns></returns>
        private string GetMaterialAfterProgramming(string inputMaterialBeforeProgramming, string inputPCBModel)
        {
            string aftProgMaterialNumber = "NA"; //Default value
            string targetFile = @"C:\MPH - KANBAN Control Local Data\WareHouseMaterialsProgrammingControl.txt";
            
            if (!System.IO.File.Exists(targetFile))
            {
                Update_WareHouseMaterialsProgrammingControlTxt();
            }

            try
            {
                using (var reader = new StreamReader(targetFile))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        // split by tab
                        var fields = line.Split('\t');
                        if (fields.Length < 3)
                            continue; // not enough columns

                        var file_bfrProgMaterialNumber = fields[0].Trim();
                        var file_pcbModel = fields[2].Trim();

                        // match both material and PCB model (ordinal comparison to mimic VB6 binary compare)
                        if (string.Equals(file_bfrProgMaterialNumber, inputMaterialBeforeProgramming, StringComparison.Ordinal) &&
                            string.Equals(file_pcbModel, inputPCBModel, StringComparison.Ordinal))
                        {
                            aftProgMaterialNumber = fields[1].Trim();
                            break; // found the match, exit loop
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "Error in GetMaterialAfterProgramming: " + ex.Message;
                MessageBox.Show(message, "Exception Caught:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return aftProgMaterialNumber;
        }

        /// <summary>
        /// VB6 equivalent: Public Sub UpdateLocalWareHouseProgrammingMaterialsControl()
        /// </summary>
        private void Update_WareHouseMaterialsProgrammingControlTxt()
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            string dir = @"C:\MPH - KANBAN Control Local Data";
            string targetFile = System.IO.Path.Combine(dir, @"WareHouseMaterialsProgrammingControl.txt");
            string header = "MaterialsBefProgramming\tMaterialsAftProgramming\tPCBModel";

            // Ensure directory exists
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            // Ensure file exists
            if (!System.IO.File.Exists(targetFile))
            {
                System.IO.File.Create(targetFile).Close();
            }

            // Overwrite file with header (this matches VB6 Open For Writing then WriteLine header)
            File.WriteAllText(targetFile, header + Environment.NewLine);

            // Append data from DB
            using (var writer = new StreamWriter(targetFile, append: true))
            {
                bool mustClose = false;
                if (cnnDLVNDB.State != ConnectionState.Open)
                {
                    cnnDLVNDB.Open();
                    mustClose = true;
                }

                try
                {
                    using (var cmd = new SqlCommand("SELECT * FROM WarehouseMaterialsProgrammingControl", cnnDLVNDB))
                    {
                        cmd.CommandType = CommandType.Text;
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string materialsBefore = reader["MaterialsBefProgramming"] != DBNull.Value
                                    ? reader["MaterialsBefProgramming"].ToString().Trim()
                                    : string.Empty;

                                string materialsAfter = reader["MaterialsAftProgramming"] != DBNull.Value
                                    ? reader["MaterialsAftProgramming"].ToString().Trim()
                                    : string.Empty;

                                string model = reader["Model"] != DBNull.Value
                                    ? reader["Model"].ToString().Trim()
                                    : string.Empty;

                                string recordSetUpData = $"{materialsBefore}\t{materialsAfter}\t{model}";

                                // VB6 wrote the line only if Len(recordSetUpData) >= 9
                                if (recordSetUpData.Length >= 9)
                                {
                                    writer.WriteLine(recordSetUpData);
                                }
                            }
                        }
                    }
                }
                finally
                {
                    if (mustClose) cnnDLVNDB.Close();
                }

                // Write the End marker (VB used ts.Write, i.e. no newline)
                writer.Write("End\tEnd\tEnd");
            }
        }

        private void Update_WareHouseMaterialsOnTrayNonProgramControlTxt()
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            string dir = @"C:\MPH - KANBAN Control Local Data";
            string targetFile = dir + @"\WareHouseMaterialsOnTrayNonProgramControl.txt";

            // Ensure directory exists
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            // Ensure file exists
            if (!System.IO.File.Exists(targetFile))
            {
                System.IO.File.Create(targetFile).Close();
            }

            // Overwrite file with header:
            string header = "MaterialsOnTrayNonProgram\tOnTray";
            File.WriteAllText(targetFile, header + Environment.NewLine);

            // Prepare SQL query (mirrors VB6 logic):
            string sql = @"SELECT * FROM WareHouseMaterialsOnTrayControl " +
                "WHERE (Materials NOT IN (SELECT DISTINCT MaterialsBefProgramming FROM [DLVNDB].[dbo].[WarehouseMaterialsProgrammingControl]))" +
                "AND (Materials NOT IN (SELECT DISTINCT MaterialsAftProgramming FROM [DLVNDB].[dbo].[WarehouseMaterialsProgrammingControl]))";

            bool mustClose = false;
            if (cnnDLVNDB.State != ConnectionState.Open)
            {
                cnnDLVNDB.Open();
                mustClose = true;
            }

            try
            {
                using (var writer = new StreamWriter(targetFile, append: true))
                using (var cmd = new SqlCommand(sql, cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string materials = reader["Materials"] != DBNull.Value
                                ? reader["Materials"].ToString().Trim()
                                : string.Empty;

                            string onTray = reader["OnTray"] != DBNull.Value
                                ? reader["OnTray"].ToString().Trim()
                                : string.Empty;

                            // VB6 wrote only if Len(recordSetUpData) >= 9 for the Materials string
                            if (materials.Length >= 9)
                            {
                                writer.WriteLine(materials + "\t" + onTray);
                            }
                        }
                    }

                    // Write the End marker without adding a newline (matches VB6 ts.Write)
                    writer.Write("End\tEnd");
                }
            }
            finally
            {
                if (mustClose) cnnDLVNDB.Close();
            }
        }

        /// <summary>
        /// Public Function MaterialsOnTrayNonProgram(getMaterials As String) As String
        /// </summary>
        /// <param name="inputMaterialNumber"></param>
        /// <returns></returns>
        private string IsMaterialsOnTrayNonProgram(string inputMaterialNumber)
        {
            string result = "NO"; // Default value if not found
            string targetFile = @"C:\MPH - KANBAN Control Local Data\WareHouseMaterialsOnTrayNonProgramControl.txt";
            if (!System.IO.File.Exists(targetFile))
            {
                // Attempt to create/update the local file (implement this method as needed)
                Update_WareHouseMaterialsOnTrayNonProgramControlTxt();
            }

            try
            {
                using (var reader = new StreamReader(targetFile))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        // Split by tab, like VB6 Split(Dataline, vbTab)
                        var fields = line.Split('\t');
                        if (fields.Length < 2)
                            continue;

                        var infileMaterialNumber = fields[0].Trim();
                        var infileYESNO = fields[1].Trim(); //YES or NO

                        // If match found, check OnTray value
                        if (string.Equals(infileMaterialNumber, inputMaterialNumber, StringComparison.Ordinal))
                        {
                            if (string.Equals(infileYESNO, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                return "YES";
                            }
                            // found but not YES -> return default (mirrors VB6 which exits loop)
                            return result;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "Error in IsMaterialsOnTrayNonProgram: " + ex.Message;
                MessageBox.Show(message, "Exception Caught:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// Update C:\MPH - KANBAN Control Local Data\PreAssyModelLayout.txt
        /// VB6 equivalent: Public Sub UpdateLocalModelWithSideAndSector()
        /// </summary>
        private void Update_PreAssyModelLayoutTxt()
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            string dir = @"C:\MPH - KANBAN Control Local Data";
            string targetFile = System.IO.Path.Combine(dir, @"PreAssyModelLayout.txt");
            string header = "Sector\tModel\tModelSide";

            // Ensure directory exists
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            // Ensure file exists
            if (!System.IO.File.Exists(targetFile))
            {
                System.IO.File.Create(targetFile).Close();
            }

            // Overwrite file with header (matches VB6 Open For Writing then WriteLine header)
            File.WriteAllText(targetFile, header + Environment.NewLine);

            // SQL to fetch distinct SMTLine, Model, ModelSide
            const string sql = "SELECT DISTINCT SMTLine, Model, ModelSide FROM PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix";

            bool mustClose = false;
            if (cnnDLVNDB.State != ConnectionState.Open)
            {
                cnnDLVNDB.Open();
                mustClose = true;
            }

            try
            {
                using (var writer = new StreamWriter(targetFile, append: true))
                using (var cmd = new SqlCommand(sql, cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string smtLine = reader["SMTLine"] != DBNull.Value ? reader["SMTLine"].ToString().Trim() : string.Empty;
                            string model = reader["Model"] != DBNull.Value ? reader["Model"].ToString().Trim() : string.Empty;
                            string modelSide = reader["ModelSide"] != DBNull.Value ? reader["ModelSide"].ToString().Trim() : string.Empty;

                            string recordSetUpData = $"{smtLine}\t{model}\t{modelSide}";

                            // VB6 wrote the line only if Len(recordSetUpData) >= 9
                            if (recordSetUpData.Length >= 9)
                            {
                                writer.WriteLine(recordSetUpData);
                            }
                        }
                    }

                    // Write the End marker without adding a newline (matches VB6 ts.Write)
                    writer.Write("End\tEnd\tEnd");
                }
            }
            finally
            {
                if (mustClose) cnnDLVNDB.Close();
            }
        }

        /// <summary>
        /// VB6 equivalent: Public Function IsPreAssyModelAllSideInLayout(getPreAssyModel As String, getSector As String) As Boolean
        /// </summary>
        /// <param name="preAssyModel"></param>
        /// <param name="sector"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        private bool IsPreAssyModelAllSideInLayout(string preAssyModel, string sector)
        {
            if (string.IsNullOrEmpty(preAssyModel)) throw new ArgumentNullException(nameof(preAssyModel));
            if (string.IsNullOrEmpty(sector)) throw new ArgumentNullException(nameof(sector));

            string FilePath = @"C:\MPH - KANBAN Control Local Data\PreAssyModelLayout.txt";
            if (!File.Exists(FilePath))
            {
                // Attempt to (re)create or update the local file, matching VB6 behavior
                Update_PreAssyModelLayoutTxt();
            }

            try
            {
                using (var reader = new StreamReader(FilePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var fields = line.Split('\t');
                        if (fields.Length < 3)
                            continue;

                        var getLocalSector = fields[0].Trim();
                        var getLocalPreAssyModel = fields[1].Trim();
                        var getLocalModelSide = fields[2].Trim();

                        if (string.Equals(getLocalPreAssyModel, preAssyModel, StringComparison.Ordinal)
                            && string.Equals(getLocalSector, sector, StringComparison.Ordinal)
                            && string.Equals(getLocalModelSide, "A", StringComparison.Ordinal))
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Preserve VB6-like silent failure behavior: return false on IO errors.
                // Optionally log the exception in your application.
                return false;
            }

            return false;
        }

        /// <summary>
        /// VB6 equivalent: Public Sub AccessPCB(ByVal getPreAssyPONumber As String, uniTopBot As String)
        /// This will access the BOM txt file (copied to local folder) and get the PCB part number (ex. 100110991) from it
        /// </summary>
        /// <param name="preAssyPONumber"></param>
        /// <param name="topBot"></param>
        private void GetPCBFromPOBOM(string preAssyPONumber, string topBot)
        {
            if (string.IsNullOrEmpty(preAssyPONumber)) return;

            // Convert SMTxxx PO numbers like VB6 Mid(getPreAssyPONumber,1,Len-3)
            string trimmedPONumber = preAssyPONumber;
            if (preAssyPONumber.StartsWith("SMT", StringComparison.OrdinalIgnoreCase) && preAssyPONumber.Length > 3)
            {
                trimmedPONumber = preAssyPONumber.Substring(0, preAssyPONumber.Length - 3);
            }

            // Ensure the local BOM is available (user must implement this to copy from server)
            CopyBOMTxtFilesFromSharedDriveToLocalFolder();

            string targetFile = System.IO.Path.Combine(@"C:\MPH - KANBAN Control Local Data\LocalBOMToPulledList", trimmedPONumber + ".txt");
            if (!File.Exists(targetFile))
            {
                MessageBox.Show("GetPCBFromPOBOM: target file = " + targetFile, "File not found:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                using (var reader = new StreamReader(targetFile))
                {
                    int rowStart = 1; // matches VB6 initialization
                    string strLine;
                    while ((strLine = reader.ReadLine()) != null)
                    {
                        rowStart = rowStart + 1;
                        // VB6 processed lines when rowStart >= 5 (skipping first 4 lines)
                        if (rowStart >= 5)
                        {
                            if (string.IsNullOrWhiteSpace(strLine)) continue;

                            var fields = strLine.Split('\t');
                            // guard against short lines
                            if (fields.Length < 3) continue;

                            string inFileMaterialNumber = fields[2].Trim();

                            // VB6: If Mid(getLocalMaterials,1,3) = "100" Then
                            if (!string.IsNullOrEmpty(inFileMaterialNumber) &&
                                inFileMaterialNumber.Length >= 3 &&
                                inFileMaterialNumber.Substring(0, 3).Equals("100", StringComparison.Ordinal))
                            {
                                // Reference to the grid on your form. Adjust if your control is different.
                                //var grid = dgvMaterialsOnTrayNonProgramMatrix as DataGridView;
                                //if (grid == null)
                                //{
                                //    // If your control is not a DataGridView, adapt the code accordingly.
                                //    return;
                                //}

                                // Add a new row and populate columns similar to VB6 TextMatrix/AddItem behavior
                                int newRowIndex = AddRowIfNeeded(dgvMaterialsOnTrayNonProgramMatrix);
                                SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 0, newRowIndex);
                                SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 1, inFileMaterialNumber);
                                SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 2, preAssyPONumber);
                                SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 3, fields.Length > 1 ? fields[1].Trim() : string.Empty);
                                SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 4, fields.Length > 4 ? fields[4] : string.Empty);
                                SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 5, "PCB");
                                SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 6, string.Empty);
                                SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 7, 1);

                                // Cells 8 may be unused in the VB code; set only 9 and 10 as original
                                if (dgvMaterialsOnTrayNonProgramMatrix.ColumnCount > 9)
                                {
                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 9, GetLocationInStock(inFileMaterialNumber));
                                }
                                if (dgvMaterialsOnTrayNonProgramMatrix.ColumnCount > 10)
                                {
                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, newRowIndex, 10, topBot);
                                }

                                // VB6 did AddItem("") which effectively prepared a blank row for next insert.
                                // DataGridView.Rows.Add() already added the row; nothing further needed.

                                // Exit after first matching material, same as VB6 Exit Do
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "GETPCBFromPOBOM: " + ex.ToString();
                MessageBox.Show(message, "Exception Caught:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// VB6 equivalent: Public Sub UpdateLocalWareHouseMaterialsSource()
        /// Update "C:\MPH - KANBAN Control Local Data\WareHouseMaterialsSource.txt"
        /// </summary>
        private void Update_WareHouseMaterialsSourceTxt()
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            const string dir = @"C:\MPH - KANBAN Control Local Data";
            const string targetFile = dir + @"\WareHouseMaterialsSource.txt";
            const string header = "Materials\tWareHouseLoc.\tSourceSector\tSSP";

            // Ensure directory exists
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            // Ensure file exists
            if (!System.IO.File.Exists(targetFile))
            {
                System.IO.File.Create(targetFile).Close();
            }

            // Overwrite file with header (matches VB6 Open For Writing then WriteLine header)
            File.WriteAllText(targetFile, header + Environment.NewLine);

            // SQL to fetch rows
            const string sql = "SELECT * FROM WareHouseMaterialsInventoryCurrentStock";

            bool mustClose = false;
            if (cnnDLVNDB.State != ConnectionState.Open)
            {
                cnnDLVNDB.Open();
                mustClose = true;
            }

            try
            {
                using (var writer = new StreamWriter(targetFile, append: true))
                using (var cmd = new SqlCommand(sql, cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string materials = reader["Materials"] != DBNull.Value ? reader["Materials"].ToString().Trim() : string.Empty;
                            string loc = reader["InventoryLocation"] != DBNull.Value ? reader["InventoryLocation"].ToString().Trim() : string.Empty;
                            string sector = reader["InventorySectorGroup"] != DBNull.Value ? reader["InventorySectorGroup"].ToString().Trim() : string.Empty;
                            string ssp = reader["InventorySSP"] != DBNull.Value ? reader["InventorySSP"].ToString().Trim() : string.Empty;

                            string recordSetUpData = $"{materials}\t{loc}\t{sector}\t{ssp}";

                            // VB6 wrote the line only if Len(recordSetUpData) >= 9
                            if (recordSetUpData.Length >= 9)
                            {
                                writer.WriteLine(recordSetUpData);
                            }
                        }
                    }

                    // Write the End marker without adding a newline (matches VB6 ts.Write)
                    writer.Write("End\tEnd\tEnd\tEnd");
                }
            }
            finally
            {
                if (mustClose) cnnDLVNDB.Close();
            }
        }

        /// <summary>
        /// VB6 equivalent: Public Function AccessLocationInStock(getMaterials As String) As String
        /// Public Function AccessWareHouseMaterialsSource(getPartNumber As String, getSourceSector As String) As String
        /// </summary>
        /// <param name="partNumber"></param>
        /// <returns></returns>
        private string GetLocationInStock(string inputPartNumber, string inputSector = "")
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\WareHouseMaterialsSource.txt";

            if (string.IsNullOrEmpty(inputPartNumber))
            {
                return "EmptyInput"; ;
            }
                

            if (!File.Exists(targetFile))
            {
                // Attempt to create/update the local file (implement as needed).
                Update_WareHouseMaterialsSourceTxt();
            }

            try
            {
                using (var reader = new StreamReader(targetFile))
                {
                    string line;
                    int rowStart = 1;
                    while ((line = reader.ReadLine()) != null)
                    {
                        rowStart++;

                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var fields = line.Split('\t');
                        if (fields.Length < 1)
                            continue;

                        string material_number_in_file = fields[0].Trim();
                        string sector_in_file = fields.Length > 2 ? fields[2].Trim() : string.Empty;

                        if (string.IsNullOrEmpty(inputSector))
                        {
                            if (string.Equals(material_number_in_file, inputPartNumber, StringComparison.Ordinal))
                            {
                                // return second column if present, otherwise empty string
                                return fields.Length > 1 ? fields[1].Trim() : string.Empty;
                            }
                        }
                        else
                        {
                            if (string.Equals(material_number_in_file, inputPartNumber, StringComparison.Ordinal) && string.Equals(sector_in_file, inputSector, StringComparison.Ordinal))
                            {
                                // return second column if present, otherwise empty string
                                return fields.Length > 1 ? fields[1].Trim() : string.Empty;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "GetLocationInStock: " + ex.ToString();
                MessageBox.Show(message, "Exception Caught:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return string.Empty;
        }

        private string GetSeries(string modelNumber)
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\MPHAllSector.txt";
            string result = "None";
            blnGoodKANBANinf = false;

            if (!File.Exists(targetFile))
            {
                // Attempt to create/update the local file (implement this method)
                Update_MPHAllSectorTxt();
            }

            try
            {
                using (var reader = new StreamReader(targetFile))
                {
                    string line;
                    int rowStart = 1;
                    while ((line = reader.ReadLine()) != null)
                    {
                        rowStart++;
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var fields = line.Split('\t');
                        if (fields.Length < 2)
                            continue;

                        string getLocalModel = fields[1].Trim();
                        string getLocalSeries = fields[0].Trim();

                        if (string.Equals(getLocalModel, modelNumber, StringComparison.Ordinal))
                        {
                            result = getLocalSeries;
                            blnGoodKANBANinf = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Preserve VB6-like silent behavior on error: return "None".
                // Optionally log the exception in your application.
            }

            if (result == "None")
            {
                MessageBox.Show(
                    $"Chuong Trinh Khong Tim Thay Thong Tin Series Cua Model #{modelNumber} Trong Local DataBase! Vui Long Click Vao Menu **Update Local DataBase** De Tien Hanh Cap Nhat!",
                    "Required Information Not Found!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// VB6 equivalent: Public Sub UpdateLocalProductionMaterialsMissingRateControl()
        /// Update C:\MPH - KANBAN Control Local Data\ProductionMaterialsMissingRateControl.txt
        /// </summary>
        private void Update_ProductionMaterialsMissingRateControlTxt()
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);
            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            const string dir = @"C:\MPH - KANBAN Control Local Data";
            const string targetFile = dir + @"\ProductionMaterialsMissingRateControl.txt";
            const string header = "Materials\tSectorType\tMissingRate";

            // Ensure directory exists
            Directory.CreateDirectory(dir);

            // Overwrite file with header (VB6 opened for writing then wrote header)
            File.WriteAllText(targetFile, header + Environment.NewLine);

            const string sql = "SELECT * FROM WarehouseMaterialsMissingRateControl";

            bool mustClose = false;
            if (cnnDLVNDB.State != ConnectionState.Open)
            {
                cnnDLVNDB.Open();
                mustClose = true;
            }

            try
            {
                using (var writer = new StreamWriter(targetFile, append: true))
                using (var cmd = new SqlCommand(sql, cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string materials = reader["Materials"] != DBNull.Value
                                ? reader["Materials"].ToString().Trim()
                                : string.Empty;

                            string sectorType = reader["SectorType"] != DBNull.Value
                                ? reader["SectorType"].ToString().Trim()
                                : string.Empty;

                            // VB6 used Val(...) to convert to numeric. Preserve numeric formatting using invariant culture.
                            string missingRate;
                            if (reader["MissingRate"] == DBNull.Value)
                            {
                                missingRate = "0";
                            }
                            else
                            {
                                var mrRaw = reader["MissingRate"].ToString().Trim();
                                if (decimal.TryParse(mrRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var mr))
                                    missingRate = mr.ToString(CultureInfo.InvariantCulture);
                                else
                                    missingRate = mrRaw; // fallback to raw text if parsing fails
                            }

                            string recordSetUpData = $"{materials}\t{sectorType}\t{missingRate}";

                            // Preserve VB6 behavior: write only if length >= 9
                            if (recordSetUpData.Length >= 9)
                            {
                                writer.WriteLine(recordSetUpData);
                            }
                        }
                    }

                    // Write the End marker without adding a newline (matches VB6 ts.Write)
                    writer.Write("End\tEnd\tEnd");
                }
            }
            finally
            {
                if (mustClose)
                {
                    try { cnnDLVNDB.Close(); } catch { /* swallow to mimic VB6 behavior */ }
                }
            }
        }

        /// <summary>
        /// VB6 equivalent: Public Function AccessProductionMaterialsMissingRateControl(getMaterials As String, getSectorType As String) As Double
        /// </summary>
        /// <param name="partNumber"></param>
        /// <param name="sector"></param>
        /// <returns></returns>
        private double GetMaterialsMissingRateControl(string partNumber, string sector)
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\ProductionMaterialsMissingRateControl.txt";
            double result = 0.0;
            if (string.IsNullOrEmpty(partNumber))
                return result;
            if (!File.Exists(targetFile))
            {
                // Attempt to create/update the local file (implement as needed).
                Update_ProductionMaterialsMissingRateControlTxt();
            }
            try
            {
                using (var reader = new StreamReader(targetFile))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;
                        var fields = line.Split('\t');
                        if (fields.Length < 3)
                            continue;
                        string getLocalMaterials = fields[0].Trim();
                        string getLocalSectorType = fields.Length > 2 ? fields[1].Trim() : string.Empty;
                        if (string.Equals(getLocalMaterials, partNumber, StringComparison.Ordinal))
                        {
                            if (double.TryParse(fields[2].Trim(), out double rate))
                            {
                                result = rate;
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Mirror VB6 silent behavior on error: return 0.0.
                // Optionally log exception here.
            }
            return result;
        }

        /// <summary>
        /// VB6 equivalent: Public Sub UpdateLocalPCBModelvsQtyPerPanelMatrixDBSetUp()
        /// </summary>
        private void Update_PCBModelPanelvsQtyperPanelTxt()
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);
            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            const string dir = @"C:\MPH - KANBAN Control Local Data";
            const string targetFile = dir + @"\PCBModelPanelvsQtyperPanel.txt";
            const string header = "PCBModel\tQuantity of Board per Panel";

            // Ensure directory exists
            Directory.CreateDirectory(dir);

            // Overwrite file with header (matches VB6 Open For Writing then WriteLine header)
            File.WriteAllText(targetFile, header + Environment.NewLine);

            // SQL to fetch rows
            const string sql = "SELECT * FROM PreAssyModelvsPostAssyQty";

            bool mustClose = false;
            if (cnnDLVNDB.State != ConnectionState.Open)
            {
                cnnDLVNDB.Open();
                mustClose = true;
            }

            try
            {
                using (var writer = new StreamWriter(targetFile, append: true))
                using (var cmd = new SqlCommand(sql, cnnDLVNDB))
                {
                    cmd.CommandType = CommandType.Text;

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string panelModel = reader["PanelModel"] != DBNull.Value ? reader["PanelModel"].ToString().Trim() : string.Empty;
                            string panelPcba = reader["PanelPCBA"] != DBNull.Value ? reader["PanelPCBA"].ToString().Trim() : string.Empty;

                            string recordSetUpData = $"{panelModel}\t{panelPcba}";

                            // Write the record (VB wrote unconditionally)
                            writer.WriteLine(recordSetUpData);
                        }
                    }

                    // Write the End marker without adding a newline (matches VB6 ts.Write)
                    writer.Write("End\tEnd");
                }
            }
            finally
            {
                if (mustClose)
                {
                    try { cnnDLVNDB.Close(); } catch { /* swallow to mimic VB6 behavior */ }
                }
            }
        }

        /// <summary>
        /// VB6 equivalent: Public Function AccessPreAssyQTyvsPostAssy(getPreAssyModel As String) As Integer
        /// </summary>
        private int GetPanelPerPO(string preAssyModel)
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\PCBModelPanelvsQtyperPanel.txt";

            if (!System.IO.File.Exists(targetFile))
            {
                Update_PCBModelPanelvsQtyperPanelTxt();
            }

            int defaultValue = 1;
            if (string.IsNullOrEmpty(preAssyModel)) return defaultValue;

            try
            {
                using (var reader = new StreamReader(targetFile))
                {
                    string line;
                    int rowStart = 0;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var fields = line.Split('\t');

                        if (rowStart > 0) // skip header (VB checked rowStart > 0)
                        {
                            if (fields.Length >= 1 && !string.Equals(fields[0].Trim(), "End", StringComparison.Ordinal))
                            {
                                var getLocalPCBModel = fields[0].Trim();
                                if (string.Equals(preAssyModel, getLocalPCBModel, StringComparison.Ordinal))
                                {
                                    if (fields.Length > 1 && int.TryParse(fields[1].Trim(), out int qty))
                                    {
                                        return qty;
                                    }
                                    // if parsing fails, return default (VB would coerce)
                                    return defaultValue;
                                }
                            }
                        }

                        rowStart++;
                    }
                }
            }
            catch (Exception)
            {
                // Mirror VB6 silent behavior: return default on errors.
            }

            return defaultValue;
        }

        /// <summary>
        /// VB6 equivalent: Public Function AccessNumOfPanelByPO(ByVal getPONumber As String, getTopBot As String) As Integer
        /// </summary>
        /// <param name="poNumber"></param>
        /// <param name="topBot"></param>
        /// <returns></returns>
        private int GetNumOfPanelByPO(string poNumber, string topBot)
        {
            string preAssyModel;
            int qtyOfPanel = 0;
            string qtyOfPO;

            for (int i = 0; i < dgvPulledListPO.Rows.Count; i++)
            {
                if (GetGridCellAsString(dgvPulledListPO, i, 1) == poNumber)
                {
                    qtyOfPO = GetGridCellAsString(dgvPulledListPO, i, 4);
                    preAssyModel = GetGridCellAsString(dgvPulledListPO, i, 2);
                    qtyOfPanel = Convert.ToInt32(qtyOfPO) / GetPanelPerPO(preAssyModel);
                }
            }
            return qtyOfPanel;
        }

        /// <summary>
        /// VB6: Public Sub UpdateLocalMaterialsMSLnFLControl()
        /// </summary>
        private void Update_MaterialsMSLnFLControlTxt()
        {
            string dir = @"C:\MPH - KANBAN Control Local Data";
            string targetFile = System.IO.Path.Combine(dir, "MaterialsMSLnFLControl.txt");

            try
            {
                // Ensure directory exists
                if (!System.IO.Directory.Exists(dir))
                {
                    System.IO.Directory.CreateDirectory(dir);
                }

                // Ensure file exists
                if (!System.IO.File.Exists(targetFile))
                {
                    System.IO.File.Create(targetFile).Close();
                }
            }
            catch (Exception ex)
            {
                string message = "Update_MaterialsMSLnFLControlTxt: " + ex.ToString();
                MessageBox.Show(message, "Exception Caught:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Overwrite file with header (equivalent to opening ForWriting and writing header)
            string header = "DLPartName\tMoistureSensitiveLevel\tFloorLifeHour";
            // Using default system encoding to be closer to VB/ASCII behaviour; change if you need UTF-8.
            File.WriteAllText(targetFile, header + Environment.NewLine, Encoding.Default);

            // Append rows from DB (equivalent to opening ForAppending)
            const string query = "SELECT DLPartName, MoistureSensitiveLevel, FloorLifeHour FROM MaterialsMSLnFLControl";

            try
            {
                
                using (var writer = new StreamWriter(targetFile, append: true, encoding: Encoding.Default))
                using (var cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB))
                {
                    cnnDLVNDB.Open();
                    using (var cmd = new SqlCommand(query, cnnDLVNDB))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            int rowStart = 0;

                            if (reader != null)
                            {
                                while (reader.Read())
                                {
                                    rowStart++;

                                    string dlPartName = reader["DLPartName"] == DBNull.Value ? string.Empty : Convert.ToString(reader["DLPartName"]).Trim();
                                    string moistureSensitiveLevel = reader["MoistureSensitiveLevel"] == DBNull.Value ? string.Empty : Convert.ToString(reader["MoistureSensitiveLevel"]).Trim();
                                    string floorLifeHour = reader["FloorLifeHour"] == DBNull.Value ? string.Empty : Convert.ToString(reader["FloorLifeHour"]).Trim();

                                    string record = string.Concat(dlPartName, "\t", moistureSensitiveLevel, "\t", floorLifeHour);

                                    // replicate VB's Len(record) >= 9 check
                                    if (record.Length >= 9)
                                    {
                                        writer.WriteLine(record);
                                    }
                                }
                            }

                            // Write final "End\tEnd\tEnd" without trailing newline (VB used ts.Write)
                            writer.Write("End\tEnd\tEnd");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "Update_MaterialsMSLnFLControlTxt (during DB read/write): " + ex.ToString();
                MessageBox.Show(message, "Exception Caught:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                
            }
        }

        /// <summary>
        /// VB6 equivalent: Public Function AccessMSL(ByVal getDLPartName As String) As String
        /// </summary>
        /// <returns></returns>
        private string GetMSL(string dlvnPartNumber)
        {
            string targetFile = @"C:\MPH - KANBAN Control Local Data\MaterialsMSLnFLControl.txt";

            if (string.IsNullOrEmpty(dlvnPartNumber))
                return string.Empty;

            if (!File.Exists(targetFile))
            {
                // Optional: recreate/populate the local file like the VB6 app did.
                Update_MaterialsMSLnFLControlTxt();
            }

            try
            {
                using (var reader = new StreamReader(targetFile))
                {
                    string line;
                    // VB6 used a row counter but only to iterate lines; we can simply read each line.
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var fields = line.Split('\t');
                        if (fields.Length < 3)
                            continue;

                        var localPart = fields[0].Trim();
                        if (string.Equals(localPart, dlvnPartNumber, StringComparison.Ordinal))
                        {
                            var msl = fields[1].Trim();
                            var fl = fields[2].Trim();
                            return $"(MSL: {msl} - FL: {fl})";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "GetMSL: " + ex.ToString();
                MessageBox.Show(message, "Exception Caught:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return string.Empty;
        }

        /// <summary>
        /// VB6 equivalent: Public Sub LogPulledListvsSector(getSector As String, getShift As String, getPONumber As String, getMaterials As String, getQtyperShift As Double, getPriority As Integer, pulledListID As String, getSumPOVal As Double, getSiplace As String, getTrackDiv As String)
        /// </summary>
        private void LogPulledListvsSector(string sector, string shift, string poNumber, string materialNumber, 
            double qtyPerShift, int priority, string pulledListID, double sumPOVal, string siplace, string trackDiv)
        {
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            // Build the textual "record" string for diagnostics (closely mirrors the VB version)
            System.DateTime plannedDate;
            if (string.IsNullOrWhiteSpace(cbbActiveDate.Text.Trim()))
            {
                plannedDate = System.DateTime.Now;
            }
            else if (!System.DateTime.TryParse(cbbActiveDate.Text.Trim(), out plannedDate))
            {
                // If parsing fails, fall back to Now (mirrors VB's CDate risk behavior more safely)
                plannedDate = System.DateTime.Now;
            }

            // format date as dd/Mon/yyyy to be similar to VB's "dd/mmm/yyyy"
            string plannedDateFormatted = plannedDate.ToString("dd/MMM/yyyy", CultureInfo.InvariantCulture);

            string strPulledListRecord =
                "('" + sector +
                "','" + poNumber +
                "','" + materialNumber +
                "','" + qtyPerShift +
                "','" + plannedDateFormatted +
                "','" + priority +
                "','" + pulledListID +
                "','" + sumPOVal +
                "','" + System.DateTime.Now.ToString("s", CultureInfo.InvariantCulture) + // ISO-like datetime
                "','" + siplace +
                "','" + trackDiv +
                "')";

            // Parameterized INSERT using OleDb (uses positional ? parameters)
            string strQuery =
                "INSERT INTO PulledListShiftvsSectorLog" +
                "(Sector,PONumber,Materials,QtyperShift,PulledListPlannedDate,Priority,pulledListID,SumPOValID,CreatedDateTime,WorkstationCode,MaterialsLocOnWSCode) " +
                "VALUES (@sector, @poNumber , @materialNumber, @qtyPerShift, @plannedDate, @priority, @pulledListID, @sumPOValID, GETDATE(), @workstationCode, @MaterialsLocOnWSCode)";

            try
            {
                bool mustClose = false;
                if (cnnDLVNDB.State == ConnectionState.Closed)
                {
                    cnnDLVNDB.Open();
                    mustClose = true;
                }

                using (var cmd = new SqlCommand(strQuery, cnnDLVNDB))
                {
                    // Add parameters in the exact order of the ? placeholders
                    cmd.Parameters.AddWithValue("sector", sector ?? string.Empty);
                    cmd.Parameters.AddWithValue("poNumber", poNumber ?? string.Empty);
                    cmd.Parameters.AddWithValue("materialNumber", materialNumber ?? string.Empty);
                    cmd.Parameters.AddWithValue("qtyPerShift", qtyPerShift);
                    cmd.Parameters.AddWithValue("plannedDate", plannedDate);           // pass as DateTime
                    cmd.Parameters.AddWithValue("priority", priority);
                    cmd.Parameters.AddWithValue("pulledListID", pulledListID ?? string.Empty);
                    cmd.Parameters.AddWithValue("sumPOValID", sumPOVal);
                    //cmd.Parameters.AddWithValue("createdDate", System.DateTime.Now);         // CreatedDateTime
                    cmd.Parameters.AddWithValue("workstationCode", siplace ?? string.Empty);
                    cmd.Parameters.AddWithValue("MaterialsLocOnWSCode", trackDiv ?? string.Empty);

                    cmd.ExecuteNonQuery();
                }

                if (mustClose && cnnDLVNDB.State != ConnectionState.Closed)
                {
                    cnnDLVNDB.Close();
                }
            }
            catch (Exception ex)
            {
                // Show a message similar to the VB MsgBox in the original routine.
                string msg = "INSERT into PulledListShiftvsSectorLog(Sector,Shift,PONumber,Materials,QtyperShift,PulledListPlannedDate,Priority,pulledListID,SumPOValID,CreatedDateTime) VALUES " +
                             strPulledListRecord +
                             Environment.NewLine + Environment.NewLine +
                             "Error: " + ex.Message;
                MessageBox.Show(msg, "LogPulledListvsSector Module Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// VB6 equivalent: Public Sub ClearPulledListIDLabel()
        /// This will DELETE FROM PulledListShiftControl
        /// </summary>
        /// <exception cref="ArgumentNullException"></exception>
        private void ClearPulledListIDLabel()
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            const string strQueryCommand = "DELETE FROM PulledListShiftControl";

            try
            {
                bool mustClose = false;
                if (cnnDLVNDB.State == ConnectionState.Closed)
                {
                    cnnDLVNDB.Open();
                    mustClose = true;
                }

                using (var cmd = new SqlCommand(strQueryCommand, cnnDLVNDB))
                {
                    cmd.ExecuteNonQuery();
                }

                if (mustClose && cnnDLVNDB.State != ConnectionState.Closed)
                {
                    cnnDLVNDB.Close();
                }
            }
            catch (Exception ex)
            {
                // Mirror VB behavior by showing an error dialog; adjust as needed for logging/throwing instead.
                MessageBox.Show($"Error executing '{strQueryCommand}': {ex.Message}", "ClearPulledListIDLabel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// VB6 equivalent: Public Sub ReleasePulledListInfor()
        /// This will create 2 txt files on \\vnmsrv601.dl.net\Program SharedFolder
        /// </summary>
        private void CreateTxtFilesToTriggerBartender()
        {
            try
            {
                string getPathFile = @"\\vnmsrv601.dl.net\Program SharedFolder\ReleasePulledListID.txt";
                if (!System.IO.File.Exists(getPathFile))
                {
                    // CreateTextFile in VB creates the file only (it does not create missing directories).
                    // File.Create returns a FileStream that must be disposed to close the handle.
                    using (System.IO.File.Create(getPathFile)) { }
                }

                getPathFile = @"\\vnmsrv601.dl.net\Program SharedFolder\ReleasePulledListPCBLabel.txt";
                if (!System.IO.File.Exists(getPathFile))
                {
                    using (System.IO.File.Create(getPathFile)) { }
                }
            }
            catch (Exception ex)
            {
                // Match prior conversions' behavior by showing an error dialog.
                // In production you may prefer logging or rethrowing instead.
                MessageBox.Show("Error creating release files: " + ex.Message, "ReleasePulledListInfor Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Vb6 equivalent: Public Sub ReleasePulledListIDLabel(ByVal getPulledListSector As String, ByVal getPulledListID As String, ByVal getPulledListUser As String, ByVal getPulledListSeries As String, ByVal getPONumber As String, ByVal getPCBNumber As String, ByVal getPCBQty As Double)
        /// This will update infor into [DLVNDB].[dbo].[PulledListShiftControl] table
        /// This information will be use by the PulledListIDLabel.btw and PulledListPCBLabel.btw on vnmsrv601.dl.net server computer;
        /// </summary>
        private void InsertDataForBartenderLabels(string pulledListSector, string pulledListID, string pulledListUser, string pulledListSeries, string poNumber, string pcbNumber, double pcbQty)
        {
            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);

            if (cnnDLVNDB == null) throw new ArgumentNullException(nameof(cnnDLVNDB));

            try
            {
                // Resolve the materials location (assumes this helper exists in your codebase)
                string strMaterialsLoc = GetLocationInStock(pcbNumber, "SMT_Group");

                // VB Mid(getPONumber, 4, 9) -> substring starting at index 3, up to 9 chars (safe against short strings)
                if (!string.IsNullOrEmpty(poNumber) && poNumber.Length > 12)
                {
                    poNumber = poNumber.Substring(3, 9); //ex SMT12345678 -> 12345678
                }

                const string insertSql =
                    "INSERT INTO PulledListShiftControl " +
                    "(PulledListSector, PulledListID, PulledListDateTime, PulledListSeries, PulledListUser, SampleMaterials, SampleMaterialsQty, SamplePO, SampleLoc) " +
                    "VALUES (@sector, @pulledListID, @pulledListDateTime, @series, @user, @materials, @qty, @po, @loc)";

                bool mustClose = false;
                if (cnnDLVNDB.State == ConnectionState.Closed)
                {
                    cnnDLVNDB.Open();
                    mustClose = true;
                }

                using (var cmd = new SqlCommand(insertSql, cnnDLVNDB))
                {
                    // Add parameters in the same order as the placeholders (OleDb uses positional parameters)
                    cmd.Parameters.AddWithValue("sector", pulledListSector ?? string.Empty);
                    cmd.Parameters.AddWithValue("pulledListID", pulledListID ?? string.Empty);
                    cmd.Parameters.AddWithValue("pulledListDateTime", System.DateTime.Now);            // PulledListDateTime
                    cmd.Parameters.AddWithValue("series", pulledListSeries ?? string.Empty);
                    cmd.Parameters.AddWithValue("user", pulledListUser ?? string.Empty);
                    cmd.Parameters.AddWithValue("materials", pcbNumber ?? string.Empty); // SampleMaterials
                    cmd.Parameters.AddWithValue("qty", pcbQty);                // SampleMaterialsQty
                    cmd.Parameters.AddWithValue("po", poNumber ?? string.Empty); // SamplePO
                    cmd.Parameters.AddWithValue("loc", strMaterialsLoc ?? string.Empty); // SampleLoc

                    cmd.ExecuteNonQuery();
                }

                if (mustClose && cnnDLVNDB.State != ConnectionState.Closed)
                {
                    cnnDLVNDB.Close();
                }
            }
            catch (Exception ex)
            {
                // Mirror prior conversions by showing an error dialog; replace with logging/throwing if desired.
                string msg = "Error inserting into PulledListShiftControl: " + ex.Message;
                MessageBox.Show(msg, "ReleasePulledListIDLabel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Vb6 equivalent: Public Function AccessPONumber(ByVal getMaterials As String) As String
        /// </summary>
        /// <param name="materialNumber"></param>
        /// <returns></returns>
        private string GetPONumberFrom_dgvMultiUniPhysicalModelPulled(string materialNumber)
        {
            string result = string.Empty;

            for (int i = 0; i < dgvMultiUniPhysicalModelPulled.Rows.Count; i++)
            {
                if (GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 3) == materialNumber)
                {
                    result = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, i, 1);
                }
            }
            return result;
        }

        /// <summary>
        /// Public Function AccessQtySMTCompBefProgInSetupSheet(getCompBefProg As String, getPCBModel As String, getModelSide As String) As Integer
        /// </summary>
        /// <returns></returns>
        private int GetPartUsedFrom_PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix(string compBefProg, string pcbModel, string modelSide)
        {
            int parts_used = 0;

            MSSQL _sql = new MSSQL();
            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);
            string query;
            string sector = cbbPulledListLine.Text.Trim();

            if (modelSide == "A")
            {
                query = "SELECT TOP 1 * FROM PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix " +
                "WHERE Model = @pcbModel " +
                "AND DLVNpn = @compBefProg " +
                "AND SMTLine = @sector";
            }
            else
            {
                query = "SELECT TOP 1 * FROM PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix " +
                "WHERE Model = @pcbModel " +
                "AND ModelSide = @modelSide " +
                "AND DLVNpn = @compBefProg " +
                "AND SMTLine = @sector";
            }

            using (var cmd = new SqlCommand(query, cnnDLVNDB))
            {
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("pcbModel", pcbModel ?? string.Empty);
                cmd.Parameters.AddWithValue("compBefProg", compBefProg ?? string.Empty);
                if (modelSide != "A")
                {
                    cmd.Parameters.AddWithValue("modelSide", modelSide ?? string.Empty);
                }
                cmd.Parameters.AddWithValue("sector", sector ?? string.Empty);

                bool openedHere = false;
                try
                {
                    if (cnnDLVNDB == null)
                        throw new InvalidOperationException("Database connection (cnnDLVNDB) is not initialized.");

                    if (cnnDLVNDB.State != ConnectionState.Open)
                    {
                        cnnDLVNDB.Open();
                        openedHere = true;
                    }

                    using (var rdr = cmd.ExecuteReader())
                    {
                        if (rdr.HasRows)
                        {
                            while (rdr.Read())
                            {
                                object val = rdr["PartUsed"];
                                if (val == DBNull.Value)
                                {
                                    // skip nulls
                                    continue;
                                }
                                else
                                {
                                    // fallback: try parsing
                                    if (!Int32.TryParse(val.ToString(), out parts_used))
                                        continue;
                                }
                            }
                        }
                    }
                    return parts_used;
                }
                finally
                {
                    if (openedHere && cnnDLVNDB.State == ConnectionState.Open)
                    {
                        try { cnnDLVNDB.Close(); } catch { /* swallow to match original VB behavior */ }
                    }
                }
            }
        }

        /// <summary>
        /// VB6: Public Sub ExportPRAPulledListFromServerByLocation()
        /// </summary>
        private void ExportPRAPulledListFromServerByLocation()
        {
            // Declare variables:
            int intNumOfAddedPO = dgvPulledListPO.Rows.Count;
            //int getStartSpreadSheet = 1;
            string pulledListID = null;
            string strIsKittingActive = null;
            //double sumPOVal = 0;
            int rowNormalMaterials = 0; //int newRow
            int rowProgrammingMaterials = 0; //int rowMaterials
            int rowOnTrayMaterials = 0; //int rowOnTrayMaterials
            int oo = 0;//indexers
            int aa = 0;//indexers
            int bb = 0;//indexers - newly added
            int getRow = 0;
            //int getLastPCBARowInExcel = 0;
            int m = 0;
            int ii = 0, xx = 0, yy = 0, zz = 0, tt = 0, hh = 0;
            string material_in_dgvMultiUni = null;
            string ponumber_in_dgvPulledListPO = null;
            string model_in_dgvPulledListPO = null; 
            string topbot_in_dgvPulledListPO = null;
            double total_qty_in_this_po = 0;
            string strMaterialAfterProgramming = null;
            string strIsMaterialOnTray = null;
            string DLVNPN_of_material = null;
            string strQuery = null;
            string local_siplace_code = null; //getWorkstationCode
            string local_track_div_value = null;//getMaterialsLocOnWSCode
            double part_used_value = 0; //getQtyperProduct
            string pulledListLineText = null;
            string outputFile = null;
            string siplace_value_from_grid; //getSiplace
            string track_div_value_from_grid; //getTrackDiv
            string strPostAssyPOnumber;
            int intNumOfKBDivided = 0;
            double dblRoundToQty = 0.0;

            try
            {
                // Prepare state
                intNumOfAddedPO = dgvPulledListPO.Rows.Count;
                pulledListID = TrimSafe(GetPulledListLogID());
                pulledListLineText = TrimSafe(cbbPulledListLine.Text);

                // Create new workbook:
                using (var workbook = new XLWorkbook())
                {
                    // Create worksheet: (& Name worksheet:)
                    var ws = workbook.AddWorksheet(NormalizeSheetName(pulledListLineText, 31) + "_PullList");

                    // initial formatting (header row background, fonts, etc.)
                    ws.Style.Font.FontName = "Times New Roman";

                    // Sort dgvUniPhysicalModelPulled by column "PartName", ascending order:
                    //dgvUniPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); PCBABin(7); TopBot(8)
                    dgvUniPhysicalModelPulled.Sort(dgvUniPhysicalModelPulled.Columns[3], ListSortDirection.Ascending);

                    // Sort dgvMultiUniPhysicalModelPulled by column "PartName", ascending order:
                    //dgvMultiUniPhysicalModelPulled: No(0); PONumber(1); Model(2); PartName(3); PartDescription(4); Qty(5); UOM(6); PCBABin(7); CommonPart(8); TopBot(9)
                    dgvMultiUniPhysicalModelPulled.Sort(dgvMultiUniPhysicalModelPulled.Columns[3], ListSortDirection.Ascending);

                    // Sort dgvPullListvsPO by column "PONumber", descending order:
                    //dgvPullListvsPO: No(0); PONumber(1); Model(2); PartName(3); PartDesc(4); Qty(5); UOM(6); NextPOUsed(7); MatClass(8); TopBot(9)
                    dgvPullListvsPO.Sort(dgvPullListvsPO.Columns[1], ListSortDirection.Descending);

                    // Set No. column in dgvPullListvsPO:
                    for (int i = 0; i < dgvPullListvsPO.Rows.Count; i++)
                    {
                        SetGridCell(dgvPullListvsPO, i, 0, i + 1);
                    }

                    m = 1;
                    getRow = 1; //Excel starts at row 1

                    // Check for planner notes if there's any change:
                    CheckPulledListPotentialIssues();
                    if (dgvPotentialIssues.Rows.Count > 2)
                    {
                        // Merge column A:C - row = getRow:
                        ws.Range(getRow, 1, getRow, 3).Merge();
                        ws.Cell(getRow, 1).Value = "IMPORTANT NOTES FROM PLANNERS:";
                        ws.Range(getRow, 1, getRow, 3).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 255); // example
                        ws.Range(getRow, 1, getRow, 3).Style.Font.Bold = true;

                        // Merge column C:H - row = getRow:
                        for (ii = 0; ii < dgvPotentialIssues.Rows.Count; ii++)
                        {
                            getRow++;
                            ws.Range(getRow, 3, getRow, 8).Merge();
                            ws.Cell(getRow, 1).Value = ii;
                            ws.Cell(getRow, 2).Value = "PO: " + GetGridCellAsString(dgvPotentialIssues, ii, 1);
                            ws.Cell(getRow, 3).Value = GetGridCellAsString(dgvPotentialIssues, ii, 2);
                            ws.Range(getRow, 3, getRow, 8).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);
                            ws.Range(getRow, 3, getRow, 8).Style.Font.Bold = true;
                        }
                    }

                    strIsKittingActive = IsKittingActive(cbbPulledListLine.Text);
                    //getStartSpreadSheet = getRow + 2;

                    // Clear dgvUniPOnQtyMaterials:
                    dgvUniPOnQtyMaterials.Rows.Clear();

                    // Clear dgvMaterialsConversionMatrix:
                    dgvMaterialsConversionMatrix.Rows.Clear();

                    // Clear dgvMaterialsOnTrayNonProgramMatrix:
                    dgvMaterialsOnTrayNonProgramMatrix.Rows.Clear();

                    //--------------------------------------------------------------------------------------------------------------------------------------
                    //--------------------------------- Loop for all materials in the grid at top right corner:---------------------------------------------
                    //--------------------------------Get data from dgvMultiUniPhysicalModelPulled to populate into:----------------------------------------
                    //----------------------dgvUniPOnQtyMaterials, dgvMaterialsConversionMatrix, dgvMaterialsOnTrayNonProgramMatrix-------------------------
                    //--------------------------------------------------------------------------------------------------------------------------------------
                    for (xx = 0; xx < dgvMultiUniPhysicalModelPulled.Rows.Count; xx++)
                    {
                        // Get material model from grid:
                        material_in_dgvMultiUni = GetGridCellAsString(dgvMultiUniPhysicalModelPulled, xx, 3);

                        // Check and set flags for this material:
                        bool isPCBModel = IsPCBModel(material_in_dgvMultiUni);
                        bool isPCBAModel = IsPCBAModel(material_in_dgvMultiUni);
                        bool isPCBComponent = IsPCBComponent(material_in_dgvMultiUni);
                        bool isPCBAComponent = IsPCBAComponent(material_in_dgvMultiUni);

                        // For some kind of models...
                        if (((_pulledList_SectorGroup == "POSTASSY_Group") && (!isPCBComponent) && (isPCBAComponent) && (!isPCBModel) && (!isPCBAModel))
                            || ((_pulledList_SectorGroup == "SMT_Group") && (isPCBComponent) && (!isPCBAComponent) && (!isPCBAModel))
                            || ((_pulledList_SectorGroup == "SMT_Group") && (isPCBComponent) && (!isPCBAComponent) && (isPCBAModel))
                            || ((_pulledList_SectorGroup != "SMT_Group") && (_pulledList_SectorGroup != "POSTASSY_Group")))
                        {
                            // Get material class:
                            string class_of_this_material = GetMaterialClass(material_in_dgvMultiUni);

                            // If material class is not PCBA or (is PCB component and is PCBA model) then...
                            if ((class_of_this_material.Length < 4 || class_of_this_material.Substring(0, Math.Min(4, class_of_this_material.Length)) != "PCBA") || (isPCBComponent && isPCBAModel))
                            {
                                // Material not 100*:
                                if ((material_in_dgvMultiUni.Length >= 6) && (material_in_dgvMultiUni.Substring(0, 3) != "100"))
                                {
                                    // Loop material not 100* in dgvPulledListPO:
                                    for (zz = 0; zz < dgvPulledListPO.Rows.Count; zz++)
                                    {
                                        //dgvPulledListPO: No(0); PONumber(1); ModelNumber(2); Side(3); POQty(4); PulledListID(5); PlannersNotice(6); POChangeInf(7)

                                        ponumber_in_dgvPulledListPO = GetGridCellAsString(dgvPulledListPO, zz, 1);
                                        model_in_dgvPulledListPO = GetGridCellAsString(dgvPulledListPO, zz, 2);
                                        topbot_in_dgvPulledListPO = GetGridCellAsString(dgvPulledListPO, zz, 3);
                                        total_qty_in_this_po = GetMaterialTotalQtyInPO(ponumber_in_dgvPulledListPO, material_in_dgvMultiUni);

                                        if (total_qty_in_this_po > 0)
                                        {
                                            strMaterialAfterProgramming = GetMaterialAfterProgramming(material_in_dgvMultiUni, model_in_dgvPulledListPO);
                                            strIsMaterialOnTray = IsMaterialsOnTrayNonProgram(material_in_dgvMultiUni); //YES or NO

                                            bool blnPreAssyModelAllSideInLayout = IsPreAssyModelAllSideInLayout(model_in_dgvPulledListPO, cbbPulledListLine.Text.Substring(3, 1));

                                            if ((topbot_in_dgvPulledListPO == "A") && (!blnPreAssyModelAllSideInLayout))
                                            {
                                                if (strMaterialAfterProgramming != "NA")
                                                {
                                                    strQuery =
                                                        "SELECT * FROM PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix" +
                                                        " WHERE Model = '" + model_in_dgvPulledListPO + "'" +
                                                        " AND DLVNpn IN " +
                                                        "(SELECT DISTINCT MaterialsAftProgramming" +
                                                        " FROM DLVNDB.dbo.WarehouseMaterialsProgrammingControl" +
                                                        " WHERE MaterialsBefProgramming = '" + material_in_dgvMultiUni + "' AND Model = '" + model_in_dgvPulledListPO + "')" +
                                                        " AND SMTLine = '" + cbbPulledListLine.Text.Substring(3, 1) + "'";
                                                }
                                                else
                                                {
                                                    strQuery =
                                                        "SELECT * FROM PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix" +
                                                        " WHERE Model = '" + model_in_dgvPulledListPO + "'" +
                                                        " AND DLVNpn = '" + material_in_dgvMultiUni + "'" +
                                                        " AND SMTLine = '" + cbbPulledListLine.Text.Substring(3, 1) + "'";
                                                }
                                            }
                                            else
                                            {
                                                if (strMaterialAfterProgramming != "NA")
                                                {
                                                    strQuery =
                                                        "SELECT * FROM PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix" +
                                                        " WHERE Model = '" + model_in_dgvPulledListPO + "'" +
                                                        " AND ModelSide = '" + topbot_in_dgvPulledListPO + "'" +
                                                        " AND DLVNpn IN " +
                                                        "(SELECT DISTINCT MaterialsAftProgramming" +
                                                        " FROM DLVNDB.dbo.WarehouseMaterialsProgrammingControl" +
                                                        " WHERE MaterialsBefProgramming = '" + material_in_dgvMultiUni + "' AND Model = '" + model_in_dgvPulledListPO + "')" +
                                                        " AND SMTLine = '" + cbbPulledListLine.Text.Substring(3, 1) + "'";
                                                }
                                                else
                                                {
                                                    strQuery =
                                                        "SELECT * FROM PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix" +
                                                        " WHERE Model = '" + model_in_dgvPulledListPO + "'" +
                                                        " AND ModelSide = '" + topbot_in_dgvPulledListPO + "'" +
                                                        " AND DLVNpn = '" + material_in_dgvMultiUni + "'" +
                                                        " AND SMTLine = '" + cbbPulledListLine.Text.Substring(3, 1) + "'";
                                                }
                                            }
                                            MSSQL _sql = new MSSQL();
                                            SqlConnection cnnDLVNDB = new SqlConnection(_sql.cnnDLVNDB);
                                            
                                            using (cnnDLVNDB)
                                            {
                                                cnnDLVNDB.Open();
                                                using (SqlCommand cmd = new SqlCommand(strQuery, cnnDLVNDB))
                                                {
                                                    cmd.CommandType = CommandType.Text;
                                                    var reader = cmd.ExecuteReader();

                                                    // Query returns at least 1 row:
                                                    if (reader.HasRows)
                                                    {
                                                        while (reader.Read())
                                                        {
                                                            local_siplace_code = TrimSafe(Convert.ToString(reader["LocalSiplace"])); //getWorkstationCode
                                                            local_track_div_value = TrimSafe(Convert.ToString(reader["LocalTrackDiv"]));
                                                            part_used_value = ToDoubleSafe(Convert.ToString(reader["PartUsed"]));

                                                            if (strMaterialAfterProgramming == "NA")
                                                            {
                                                                if (strIsMaterialOnTray == "NO")
                                                                {
                                                                    //---------------------------------------------------------------------------------------
                                                                    //------------------------Populate dgvUniPOnQtyMaterials:--------------------------------
                                                                    //---------------------------------------------------------------------------------------
                                                                    //dgvUniPOnQtyMaterials: No(0); Materials(1); PONumber(2); POModel(3); POQty(4); WSCode(5); WSLocCode(6); QtyPerLoc(7); QtyPerPOLoc(8); InStockLoc(9); TopBot(10); POMat(11)
                                                                    int rowIndex = AddRowIfNeeded(dgvUniPOnQtyMaterials);
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 0, rowIndex);  // value is boxed and passed as object
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 1, material_in_dgvMultiUni);
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 2, ponumber_in_dgvPulledListPO);
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 3, model_in_dgvPulledListPO);
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 4, total_qty_in_this_po);
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 5, local_siplace_code);
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 6, local_track_div_value);
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 7, part_used_value);
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 9, GetLocationInStock(material_in_dgvMultiUni)); //
                                                                    SetGridCell(dgvUniPOnQtyMaterials, rowIndex, 10, topbot_in_dgvPulledListPO);
                                                                }
                                                                else
                                                                {
                                                                    //---------------------------------------------------------------------------------------
                                                                    //--------------------Populate dgvMaterialsOnTrayNonProgramMatrix:-----------------------
                                                                    //---------------------------------------------------------------------------------------
                                                                    int rowIndex = AddRowIfNeeded(dgvMaterialsOnTrayNonProgramMatrix);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 0, rowIndex);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 1, material_in_dgvMultiUni);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 2, ponumber_in_dgvPulledListPO);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 3, model_in_dgvPulledListPO);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 4, total_qty_in_this_po);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 5, local_siplace_code);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 6, local_track_div_value);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 7, part_used_value);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 9, GetLocationInStock(material_in_dgvMultiUni));
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex, 10, topbot_in_dgvPulledListPO);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                DLVNPN_of_material = TrimSafe(Convert.ToString(reader["DLVNpn"]));
                                                                int returnedAccessQtySMTCompBefProgInSetupSheet = GetPartUsedFrom_PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix(material_in_dgvMultiUni, model_in_dgvPulledListPO, topbot_in_dgvPulledListPO);

                                                                //---------------------------------------------------------------------------------------
                                                                //------------------------Populate dgvMaterialsConversionMatrix:-------------------------
                                                                //---------------------------------------------------------------------------------------
                                                                //dgvMaterialsConversionMatrix: No(0), MaterialsBef(1), MaterialsAft(2), PONumber(3), POModel(4), POQty(5), WSCode(6), WSLocCode(7), QtyPerLoc(8), QtyPerPOLoc(9), InStockLoc(10), TopBot(11)
                                                                int rowIndex = AddRowIfNeeded(dgvMaterialsConversionMatrix);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 0, rowIndex);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 1, material_in_dgvMultiUni);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 2, DLVNPN_of_material);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 3, ponumber_in_dgvPulledListPO);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 4, model_in_dgvPulledListPO);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 5, total_qty_in_this_po);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 6, local_siplace_code);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 7, local_track_div_value);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 8, part_used_value);
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 10, "(" + material_in_dgvMultiUni + ")" + GetLocationInStock(material_in_dgvMultiUni));
                                                                SetGridCell(dgvMaterialsConversionMatrix, rowIndex, 11, topbot_in_dgvPulledListPO);

                                                                if (returnedAccessQtySMTCompBefProgInSetupSheet > 0)
                                                                {
                                                                    int rowIndex2 = AddRowIfNeeded(dgvMaterialsOnTrayNonProgramMatrix);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 0, rowIndex2);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 1, material_in_dgvMultiUni);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 2, ponumber_in_dgvPulledListPO);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 3, model_in_dgvPulledListPO);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 4, total_qty_in_this_po);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 5, local_siplace_code);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 6, local_track_div_value);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 7, part_used_value);
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 9, GetLocationInStock(material_in_dgvMultiUni));
                                                                    SetGridCell(dgvMaterialsOnTrayNonProgramMatrix, rowIndex2, 10, topbot_in_dgvPulledListPO);
                                                                }
                                                            }
                                                        } // foreach row
                                                    }
                                                    else // Query returns 0 row
                                                    {
                                                        if (topbot_in_dgvPulledListPO == "A")
                                                        {
                                                            bool adminNo = false;
                                                            if (_strUserName.ToUpper() == "ALAM" || _strUserName.ToUpper() == "ADMINISTRATOR")
                                                            {
                                                                var dr = MessageBox.Show(
                                                                    "Chuong Trinh Khong Tim Thay Thong Tin Setup Sheet Link Kien Sau Khi Nap ROM Cua Linh Kien #" + material_in_dgvMultiUni + " Cua Model #" + model_in_dgvPulledListPO + " Trong Database [DLVNDB].[dbo].[PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix]" +
                                                                    Environment.NewLine + Environment.NewLine + "- Is Programming ROM: " + strMaterialAfterProgramming +
                                                                    Environment.NewLine + "- Data Query: " + strQuery,
                                                                    "Missing Setup Sheet",
                                                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                                                if (dr == DialogResult.No)
                                                                {
                                                                    adminNo = true;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                MessageBox.Show("Chuong Trinh Khong Tim Thay Thong Tin Setup Sheet Link Kien Sau Khi Nap ROM Cua Linh Kien #" + material_in_dgvMultiUni + " Cua Model #" + model_in_dgvPulledListPO + " Trong Database [DLVNDB].[dbo].[PreAssyComponentsTracking_ModelvsMaterialsLayoutMatrix]" +
                                                                    Environment.NewLine + Environment.NewLine + "- Is Programming ROM: " + strMaterialAfterProgramming +
                                                                    Environment.NewLine + "- Data Query: " + strQuery,
                                                                    "Missing Setup Sheet", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                            }

                                                            if (adminNo)
                                                            {
                                                                // Save the workbook and open it for debugging if admin chooses No
                                                                outputFile = SaveWorkbookAndOpen(workbook, pulledListLineText);
                                                                return;
                                                            }
                                                        }
                                                    } // End of Query returns at least 1 row!
                                                }
                                            } 
                                        } // End of If material has quantity greater than 0!
                                    } // End of Loop material not 100* in dgvPulledListPO!
                                } // End of Material not 100*!
                            } // End of If material class is not PCBA or (is PCB component and is PCBA model) then...!
                            else
                            {
                                MessageBox.Show("Material: " + material_in_dgvMultiUni + " is of PCBA class or is PCB component and is PCBA model, skipped processing!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        } // End of For some kind of models...!
                        else
                        {
                            string message = "Material: " + material_in_dgvMultiUni + " does not match the model type criteria, skipped processing!";
                            MessageBox.Show(message, "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    } // End of Loop for all materials in the grid at top right corner!

                    //--------------------------------------------------------------------------------------------------------------
                    // Find the PCB of PO:
                    for (xx = 0; xx < intNumOfAddedPO; xx++)
                    {
                        strPostAssyPOnumber = GetGridCellAsString(dgvPulledListPO, xx, 1);
                        GetPCBFromPOBOM(strPostAssyPOnumber, GetGridCellAsString(dgvPulledListPO, xx, 3));
                    }



                    if ((_pulledList_SectorGroup == "SMT_Group") || (_pulledList_SectorGroup == "POSTASSY_Group"))
                    {
                        oo = 1; // starting row for header area in earlier implementation

                        //--------------------------------------------------------------------------------------------------------------------------------
                        //-----------------------------------------NORMAL PHYSICAL MATERIALS SECTION------------------------------------------------------
                        //----------------------------------Fill data to excel file table using dgvUniPOnQtyMaterials ------------------------------------
                        //--------------------------------------------------------------------------------------------------------------------------------

                        // Set the rowNormalMaterials equals 3, then as 3++ = 4 to account for table's header;
                        // As the for (yy = 0..) loop goes on, rowNormalMaterials will increase with yy:
                        rowNormalMaterials = 3; //rowNormalMaterials increases with yy
                        oo = 3; //oo increases with yy; note indexes are 1 based in ClosedXML

                        // Fill in Excel file with information from dgvUniPOnQtyMaterials - starts from row 4, column Z:
                        for (yy = 0; yy < dgvUniPOnQtyMaterials.Rows.Count; yy++)
                        {
                            oo++; //starts at 4
                            //dgvUniPOnQtyMaterials: No(0); Materials(1); PONumber(2); POModel(3); POQty(4); WSCode(5); WSLocCode(6); QtyPerLoc(7); QtyPerPOLoc(8); InStockLoc(9); TopBot(10); POMat(11)
                            string wsPONumber = GetGridCellAsString(dgvUniPOnQtyMaterials, yy, 2);
                            string wsModel = GetGridCellAsString(dgvUniPOnQtyMaterials, yy, 3);
                            string wsMaterial = GetGridCellAsString(dgvUniPOnQtyMaterials, yy, 1);
                            string wsTopBot = GetGridCellAsString(dgvUniPOnQtyMaterials, yy, 10);
                            string wsQtyByPanel = GetGridCellAsString(dgvUniPOnQtyMaterials, yy, 7);
                            siplace_value_from_grid = GetGridCellAsString(dgvUniPOnQtyMaterials, yy, 5);
                            track_div_value_from_grid = GetGridCellAsString(dgvUniPOnQtyMaterials, yy, 6);
                            
                            double materialMissRate = GetMaterialsMissingRateControl(wsMaterial, "PRA");
                            var wsNumOfPanelInPO = GetNumOfPanelByPO(wsPONumber, wsTopBot);

                            ws.Cell(oo, 26).Value = oo - 3;//row oo, column 26(Z) = No.
                            ws.Cell(oo, 27).Value = wsPONumber;//row oo, column 27(AA) = PONumber
                            ws.Cell(oo, 28).Value = "'" + wsModel;//row oo, column 28(AB) =  Model
                            ws.Cell(oo, 29).Value = wsNumOfPanelInPO;// row oo, column 29(AC) = NumOfPanelByPO
                            ws.Cell(oo, 30).Value = wsTopBot;// row oo, column 30(AD) = TOP ? BOT
                            ws.Cell(oo, 31).Value = "'" + wsMaterial;// row oo, column 31(AE) =  Materials
                            ws.Cell(oo, 32).Value = GetGridCellAsString(dgvUniPOnQtyMaterials, yy, 4);// row oo, column 32(AF) = Materials Qty By PO Without MissingRate
                            ws.Cell(oo, 33).Value = "*" + siplace_value_from_grid + "*";// row oo, column 33(AG) = SIPLACE
                            ws.Cell(oo, 34).Value = "*" + track_div_value_from_grid + "*";// row oo, column 34(AH) = TRACK-Div
                            ws.Cell(oo, 35).Value = ws.Cell(oo, 33).Value.ToString() + "(" + ws.Cell(oo, 34).Value.ToString() + ")";// row oo, column 35(AI) = SIPLACE + (TRACK-Div)
                            ws.Cell(oo, 36).Value = wsQtyByPanel;// row oo, column 36(AJ) = Qty By Panel

                            // row oo, column 37(AK) = Qty In Total (accounted for missing rate) = Panels per PO * Boards per Panel * Missing Rate:
                            // row oo, column 29(AC) = NumOfPanelInPO = Panels per PO
                            // row oo, column 36(AJ) = Qty By Panel = Boards per Panel

                            var num_of_panels = ToDoubleSafe(ws.Cell(oo, 29).Value);
                            var qty_of_pcb_per_panel = ToDoubleSafe(ws.Cell(oo, 36).Value);
                            if (materialMissRate == 0)
                            {
                                ws.Cell(oo, 37).Value = num_of_panels * qty_of_pcb_per_panel;
                            }
                            else
                            {
                                ws.Cell(oo, 37).Value = Math.Round(num_of_panels * qty_of_pcb_per_panel * (materialMissRate + 1) + 0.5, 0); //missrate added
                            }

                            // row oo, column 38(AL) = Storage Bin
                            ws.Cell(oo, 38).SetValue(GetGridCellAsString(dgvUniPOnQtyMaterials, yy, 9));

                            // row oo, column 40(AN) = Std. REEL Qty
                            ws.Cell(oo, 40).Value = 0;

                            // Get PO order (which will run first):
                            string strPostAssyPONumber = ws.Cell(oo, 27).GetString();
                            int intOrderOfThisPO = 0;
                            // Loop through dgvPulledListPO and find the order to set intOrderOfThisPO:
                            for (tt = 0; tt < intNumOfAddedPO; tt++)
                            {
                                if (strPostAssyPONumber == GetGridCellAsString(dgvPulledListPO, tt, 1))
                                {
                                    intOrderOfThisPO = tt + 1;
                                    break;
                                }
                            }

                            // Find duplicate Item (Material) in list of POs:
                            // oo is indexer of current row in Excel file being processed
                            // xx is indexer running from row 4 (first input) to rowNormalMaterials + 1
                            // meaning, if the newly added row is a material that exists already in the list, we mark the existing row index as dupRow
                            bool dupItem = false;
                            int dupRow = 0;
                            for (xx = 4; xx <= rowNormalMaterials + 1; xx++) //note indexes are 1 based in ClosedXML
                            {
                                //column 46 (AT - PulledList form - SIPLACE + (TRACK-Div))
                                //column 35 (AI - left side draft - SIPLACE + (TRACK-Div))
                                //column 44 (AR - PulledList form - Materials)
                                //column 31 (AE - left side draft - Materials)
                                if (ws.Cell(xx, 46) != null && !ws.Cell(xx, 46).IsEmpty() &&
                                    ws.Cell(xx, 44) != null && !ws.Cell(xx, 44).IsEmpty() &&
                                    ws.Cell(oo, 35) != null && !ws.Cell(oo, 35).IsEmpty() &&
                                    ws.Cell(oo, 31) != null && !ws.Cell(oo, 31).IsEmpty())
                                {
                                    if (ws.Cell(xx, 46).GetString() == ws.Cell(oo, 35).GetString()
                                        && ws.Cell(xx, 44).GetString() == ws.Cell(oo, 31).GetString())
                                    {
                                        dupItem = true;
                                        dupRow = xx;
                                        break;
                                    }
                                }
                            }

                            // If no duplicate >> add new row to the official pulled list:
                            if (!dupItem)
                            {
                                var mslInfo = GetMSL(ws.Cell(oo, 31).Value.GetText());

                                rowNormalMaterials++;//starts at row 4
                                ws.Cell(rowNormalMaterials, 43).Value = rowNormalMaterials - 3; //Column 43(AQ) = No.
                                ws.Cell(rowNormalMaterials, 44).Value = ws.Cell(oo, 31).Value;
                                ws.Cell(rowNormalMaterials, 45).Value = ws.Cell(oo, 38).GetString() + mslInfo;
                                ws.Cell(rowNormalMaterials, 46).Value = ws.Cell(oo, 35).Value;
                                ws.Cell(rowNormalMaterials, 47 + intNumOfAddedPO).Value = ws.Cell(oo, 37).Value;
                                ws.Cell(rowNormalMaterials, 46 + intOrderOfThisPO).Value = ToDoubleSafe(ws.Cell(oo, 37).Value);
                            }
                            else // there's a duplicate row (material), add the quantity to that row, new column (AU, AV, ...) instead of adding a new row:
                            {
                                ws.Cell(dupRow, 47 + intNumOfAddedPO).Value = ToDoubleSafe(ws.Cell(oo, 37).Value) + ToDoubleSafe(ws.Cell(dupRow, 47 + intNumOfAddedPO).Value);
                                ws.Cell(dupRow, 46 + intOrderOfThisPO).Value = ToDoubleSafe(ws.Cell(dupRow, 46 + intOrderOfThisPO).Value) + ToDoubleSafe(ws.Cell(oo, 37).Value);
                            }
                        } // End of Fill in Excel file with information from dgvUniPOnQtyMaterials - starts from row 4, column Z:

                        // Add input POs to Normal Material Table's headers; Start from column 47 = AU:
                        //We must put this paragraph here, otherwise LogPulledListvsSector will fail because PO headers are not there yet!
                        for (xx = 0; xx < intNumOfAddedPO; xx++)
                        {
                            string postAssyPONumber = GetGridCellAsString(dgvPulledListPO, xx, 1);
                            string model = GetGridCellAsString(dgvPulledListPO, xx, 2);
                            string topbot = GetGridCellAsString(dgvPulledListPO, xx, 3);
                            var cell = ws.Cell(3, 46 + xx + 1);
                            cell.Value = postAssyPONumber + "\n" + model + "\n(" + topbot + ")";
                            cell.Style.Alignment.WrapText = true;
                            cell.Style.Font.Bold = true;
                        }

                        // Looping through the excel table just created to log the information (PO - PulledListID - Materials) that we have processed onto server vnmsrv601, table PulledListShiftvsSectorLog (1):
                        for (ii = 1; ii <= intNumOfAddedPO; ii++) //1 PO = 1 column AU; 2 PO = AU, AV; ...
                        {
                            for (int jj = 4; jj <= rowNormalMaterials; jj++)
                            {
                                string siplace_and_trackdiv = ws.Cell(jj, 46).GetString();
                                int index_of_position_in_col_AT = InStrSafe(siplace_and_trackdiv, "(");
                                string poNumber = ws.Cell(3, 46 + ii).GetString();//POs Number added to row 3 (headers), columns AU, AV, AW, ...
                                string materialNumber = ws.Cell(jj, 44).GetString();
                                double qtyByPO = ToDoubleSafe(ws.Cell(jj, 46 + ii).Value); //Qty of POs added to row jj, columns AU, AV, AW, ...
                                siplace_value_from_grid = MidSafe(siplace_and_trackdiv, 2, index_of_position_in_col_AT - 3);
                                track_div_value_from_grid = MidSafe(siplace_and_trackdiv, index_of_position_in_col_AT + 2, siplace_and_trackdiv.Length - index_of_position_in_col_AT - 3);
                                if (qtyByPO > 0)
                                {
                                    LogPulledListvsSector(cbbPulledListLine.Text, cbbPulledListShift.Text, poNumber, materialNumber, qtyByPO, ii, pulledListID, _sumPOValue, siplace_value_from_grid, track_div_value_from_grid);
                                }
                            }
                        }

                        //----------------------------------------------------------------------------------------------------------------------------------------------------
                        //---------------------------------MATERIALS REQUIRED PROGRAMMING SECTION - FILL IN FROM dgvMaterialsConversionMatrix --------------------------------
                        //----------------------------------------------------------------------------------------------------------------------------------------------------

                        // Set the rowProgrammingMaterials equals last row of materials table + 6 rows spared for the "DANH SACH LINH KIEN NAP ROM" table's header:
                        rowProgrammingMaterials = rowNormalMaterials + 6;
                        aa = oo + 6;

                        // Fill in Excel file with information from dgvMaterialsConversionMatrix, starting from row below last physical materials row) (column Z to AN first):
                        for (yy = 0; yy < dgvMaterialsConversionMatrix.Rows.Count; yy++)
                        {
                            aa++;
                            //dgvMaterialsConversionMatrix: No(0), MaterialsBef(1), MaterialsAft(2), PONumber(3), POModel(4), POQty(5), WSCode(6), WSLocCode(7), QtyPerLoc(8), QtyPerPOLoc(9), InStockLoc(10), TopBot(11)
                            // Get values in advance:
                            string wsPONumber = GetGridCellAsString(dgvMaterialsConversionMatrix, yy, 3); //PONumber
                            string wsModel = GetGridCellAsString(dgvMaterialsConversionMatrix, yy, 4); //POModel
                            string wsMaterial = GetGridCellAsString(dgvMaterialsConversionMatrix, yy, 2); //MaterialsAft
                            string wsTopBot = GetGridCellAsString(dgvMaterialsConversionMatrix, yy, 11); //TopBot
                            string wsQtyByPanel = GetGridCellAsString(dgvMaterialsConversionMatrix, yy, 7); //QtyPerLoc
                            siplace_value_from_grid = GetGridCellAsString(dgvMaterialsConversionMatrix, yy, 5); //WSCode
                            track_div_value_from_grid = GetGridCellAsString(dgvMaterialsConversionMatrix, yy, 6); //WSLocCode
                            var wsStorageBin = GetGridCellAsString(dgvMaterialsConversionMatrix, yy, 10); //InStockLoc 
                            var wsMaterialQtyByPOWithoutMissRate = GetGridCellAsString(dgvMaterialsConversionMatrix, yy, 5); //POQty

                            double materialMissRate = GetMaterialsMissingRateControl(wsMaterial, "PRA");
                            var wsNumOfPanelInPO = GetNumOfPanelByPO(wsPONumber, wsTopBot); //NumOfPanelInPO
                            

                            // Set worksheet cells:
                            ws.Cell(aa, 26).Value = aa - 6;//row aa, column 26(Z) = No. ; starts at oo + 1
                            ws.Cell(aa, 27).Value = wsPONumber;//row aa, column 27(AA) = PONumber - add "'" to avoid Excel auto-formatting
                            ws.Cell(aa, 28).Value = "'" + wsModel;//row aa, column 28(AB) =  POModel; force excel convert to string
                            ws.Cell(aa, 29).Value = wsNumOfPanelInPO;// row aa, column 29(AC) = NumOfPanelInPO
                            ws.Cell(aa, 30).Value = wsTopBot;// row aa, column 30(AD) = TOP ? BOT
                            ws.Cell(aa, 31).Value = "'" + wsMaterial;// row aa, column 31(AE) =  Materials; force excel convert to string
                            ws.Cell(aa, 32).Value = wsMaterialQtyByPOWithoutMissRate;// row aa, column 32(AF) = Materials Qty By PO Without MissingRate
                            ws.Cell(aa, 33).Value = "*" + siplace_value_from_grid + "*";// row aa, column 33(AG) = SIPLACE
                            ws.Cell(aa, 34).Value = "*" + track_div_value_from_grid + "*";// row aa, column 34(AH) = TRACK-Div
                            ws.Cell(aa, 35).Value = ws.Cell(aa, 33).Value.ToString() + "(" + ws.Cell(aa, 34).Value.ToString() + ")";// row aa, column 35(AI) = SIPLACE + (TRACK-Div)
                            ws.Cell(aa, 36).Value = wsQtyByPanel;// row aa, column 36(AJ) = Qty By Panel

                            // row aa, column 37(AK) = Qty In Total (accounted for missing rate) = Panels per PO * Boards per Panel * Missing Rate:
                            // row aa, column 29(AC) = NumOfPanelInPO = Panels per PO
                            // row aa, column 36(AJ) = Qty By Panel = Boards per Panel
                            var num_of_panels = ToDoubleSafe(ws.Cell(aa, 29).Value);
                            var qty_of_pcb_per_panel = ToDoubleSafe(ws.Cell(aa, 36).Value);

                            if (materialMissRate == 0)
                            {
                                ws.Cell(aa, 37).SetValue(num_of_panels * qty_of_pcb_per_panel);
                            }
                            else
                            {
                                ws.Cell(aa, 37).SetValue(Math.Round(num_of_panels * qty_of_pcb_per_panel * (materialMissRate + 1) + 0.5, 0));
                            }

                            // row aa, column 38(AL) = Storage Bin
                            ws.Cell(aa, 38).SetValue(wsStorageBin);

                            // row aa, column 40(AN) = Std. REEL Qty
                            ws.Cell(aa, 40).SetValue(0);

                            // Get PO order (which will run first):
                            string strPostAssyPONumber = ws.Cell(aa, 27).GetString();
                            int intOrderOfThisPO = 0;
                            // Loop through dgvPulledListPO and find the order to set intOrderOfThisPO:
                            for (tt = 0; tt < intNumOfAddedPO; tt++)
                            {
                                if (strPostAssyPONumber == GetGridCellAsString(dgvPulledListPO, tt, 1))
                                {
                                    intOrderOfThisPO = tt + 1;
                                    break;
                                }
                            }

                            // Find duplicate Item (Material) in list of POs (dgvMaterialsConversionMatrix):
                            // aa is indexer of current row in Excel file being processed
                            // xx is indexer running from row (rowNormalMaterials + 6) (first row of programming materials table) to rowNormalMaterials + 1
                            // meaning, if the newly added row is a material that exists already in the list, we mark the existing row index as dupRow
                            bool dupItem = false;
                            int dupRow = 0;
                            for (xx = rowNormalMaterials + 7; xx <= rowProgrammingMaterials + 1; xx++)
                            {
                                //column 46(AT - PulledList form - SIPLACE + (TRACK-Div)) = column 35 (AI - left side draft - SIPLACE + (TRACK-Div))
                                //column 44(AR - PulledList form - Materials) = column 31 (AE - left side draft - Materials)
                                if (ws.Cell(xx, 46) != null && !ws.Cell(xx, 46).IsEmpty() && 
                                ws.Cell(xx, 44) != null && !ws.Cell(xx, 44).IsEmpty() && 
                                ws.Cell(aa, 35) != null && !ws.Cell(aa, 35).IsEmpty() &&
                                ws.Cell(aa, 31) != null && !ws.Cell(aa, 31).IsEmpty())
                                {
                                    if (ws.Cell(xx, 46).GetString() == ws.Cell(aa, 35).GetString()
                                        && ws.Cell(xx, 44).GetString() == ws.Cell(aa, 31).GetString())
                                    {
                                        dupItem = true;
                                        dupRow = xx;
                                        break;
                                    }
                                }
                            }

                            //  If no duplicate >> add new row to the official pulled list ("DANH SACH LINH KIEN NAP ROM):
                            if (!dupItem)
                            {
                                rowProgrammingMaterials++;
                                ws.Cell(rowProgrammingMaterials, 43).Value = rowProgrammingMaterials - rowNormalMaterials - 6;
                                ws.Cell(rowProgrammingMaterials, 44).Value = ws.Cell(aa, 31).Value;
                                ws.Cell(rowProgrammingMaterials, 45).Value = ws.Cell(aa, 38).Value.ToString() + GetMSL(ws.Cell(aa, 31).Value.GetText());
                                ws.Cell(rowProgrammingMaterials, 46).Value = ws.Cell(aa, 35).Value;
                                ws.Cell(rowProgrammingMaterials, 47 + intNumOfAddedPO).Value = ws.Cell(aa, 37).Value;
                                ws.Cell(rowProgrammingMaterials, 46 + intOrderOfThisPO).Value = ToDoubleSafe(ws.Cell(aa, 37).Value);
                            }
                            else //If there's a duplicate row (material), add the quantity to that row, new column (AU, AV, ...) instead of adding a new row:
                            {
                                ws.Cell(dupRow, 47 + intNumOfAddedPO).Value = ToDoubleSafe(ws.Cell(aa, 37).Value) + ToDoubleSafe(ws.Cell(dupRow, 47 + intNumOfAddedPO).Value);
                                ws.Cell(dupRow, 46 + intOrderOfThisPO).Value = ToDoubleSafe(ws.Cell(dupRow, 46 + intOrderOfThisPO).Value) + ToDoubleSafe(ws.Cell(aa, 37).Value);
                            }
                        } // End of Fill in Excel file with information from dgvMaterialsConversionMatrix - starts from row after last physical materials row + 6 spared rows, column Z

                        // Looping through the excel table just created to log the information (PO - PulledListID - Materials) that we have processed onto server vnmsrv601, table PulledListShiftvsSectorLog (2):
                        for (ii = 1; ii <= intNumOfAddedPO; ii++) //must set ii starts from 1 so that we can calculate columns correctly
                        {
                            for (int jj = rowNormalMaterials + 7; jj <= rowProgrammingMaterials; jj++)
                            {
                                string siplace_and_trackdiv = ws.Cell(jj, 46).GetString();
                                int index_of_position_in_col_AT = InStrSafe(siplace_and_trackdiv, "(");
                                string poNumber = ws.Cell(3, 46 + ii).GetString(); //POs Number added to row 3 (headers), columns AU, AV, AW, ...; ex. SMT102160895B_B654048010(B)
                                string materialNumber = ws.Cell(jj, 44).GetString();
                                double qtyByPO = ToDoubleSafe(ws.Cell(jj, 46 + ii).Value); //Qty of POs added to row jj, columns AU, AV, AW, ...
                                siplace_value_from_grid = MidSafe(siplace_and_trackdiv, 2, index_of_position_in_col_AT - 3);
                                track_div_value_from_grid = MidSafe(siplace_and_trackdiv, index_of_position_in_col_AT + 2, siplace_and_trackdiv.Length - index_of_position_in_col_AT - 3);
                                if (qtyByPO > 0)
                                {
                                    LogPulledListvsSector(cbbPulledListLine.Text, cbbPulledListShift.Text, poNumber, materialNumber, qtyByPO, ii, pulledListID, _sumPOValue, siplace_value_from_grid, track_div_value_from_grid);
                                }
                            }
                        }

                        //----------------------------------------------------------------------------------------------------------------------------------------------------
                        //----------------------------MATERIALS NON PROGRAMMING ON TRAY SECTION - FILL IN FROM dgvMaterialsOnTrayNonProgramMatrix ----------------------------
                        //----------------------------------------------------------------------------------------------------------------------------------------------------

                        // Set the rowOnTrayMaterials equals last row of materials required programming table + 6 rows spared for the "DANH SACH LINH KIEN CHUAN BI THEO TRAY" table's header:
                        rowOnTrayMaterials = rowProgrammingMaterials + 6;
                        bb = aa + 6;

                        // Fill in Excel file with information from dgvMaterialsOnTrayNonProgramMatrix, starting from row below last physical materials row) (column Z to AN first):
                        for (yy = 0; yy < dgvMaterialsOnTrayNonProgramMatrix.Rows.Count; yy++)
                        {
                            bb++;
                            // Get values in advance:
                            string wsPONumber = GetGridCellAsString(dgvMaterialsOnTrayNonProgramMatrix, yy, 2); //PONumber
                            string wsModel = GetGridCellAsString(dgvMaterialsOnTrayNonProgramMatrix, yy, 3); //POModel
                            string wsMaterial = GetGridCellAsString(dgvMaterialsOnTrayNonProgramMatrix, yy, 1); //Materials
                            string wsTopBot = GetGridCellAsString(dgvMaterialsOnTrayNonProgramMatrix, yy, 10); //TopBot
                            string wsQtyByPanel = GetGridCellAsString(dgvMaterialsOnTrayNonProgramMatrix, yy, 7); //QtyPerLoc
                            siplace_value_from_grid = GetGridCellAsString(dgvMaterialsOnTrayNonProgramMatrix, yy, 5); //WSCode
                            track_div_value_from_grid = GetGridCellAsString(dgvMaterialsOnTrayNonProgramMatrix, yy, 6); //WSLocCode
                            double materialMissRate = GetMaterialsMissingRateControl(wsMaterial, "PRA");
                            var wsNumOfPanelInPO = GetNumOfPanelByPO(wsPONumber, wsTopBot); //NumOfPanelInPO
                            var wsMaterialQtyByPOWithoutMissRate = GetGridCellAsString(dgvMaterialsOnTrayNonProgramMatrix, yy, 5);

                            // Set worksheet cells:
                            ws.Cell(bb, 26).Value = bb - aa - 6;//row bb, column 26(Z) = No.
                            ws.Cell(bb, 27).Value = wsPONumber;//row bb, column 27(AA) = PONumber - add "'" to avoid Excel auto-formatting
                            ws.Cell(bb, 28).Value = "'" + wsModel;//row bb, column 28(AB) =  POModel
                            ws.Cell(bb, 29).Value = wsNumOfPanelInPO;// row bb, column 29(AC) = NumOfPanelInPO
                            ws.Cell(bb, 30).Value = wsTopBot;// row bb, column 30(AD) = TOP ? BOT
                            ws.Cell(bb, 31).Value = "'" + wsMaterial;// row bb, column 31(AE) =  Materials
                            ws.Cell(bb, 32).Value = wsMaterialQtyByPOWithoutMissRate;// row bb, column 32(AF) = Materials Qty By PO Without MissingRate
                            ws.Cell(bb, 33).Value = "*" + siplace_value_from_grid + "*";// row bb, column 33(AG) = SIPLACE
                            ws.Cell(bb, 34).Value = "*" + track_div_value_from_grid + "*";// row bb, column 34(AH) = TRACK-Div
                            ws.Cell(bb, 35).Value = ws.Cell(bb, 33).Value.ToString() + "(" + ws.Cell(bb, 34).Value.ToString() + ")";// row bb, column 35(AI) = SIPLACE + (TRACK-Div)
                            ws.Cell(bb, 36).Value = wsQtyByPanel;// row bb, column 36(AJ) = Qty By Panel

                            // row bb, column 37(AK) = Qty In Total (accounted for missing rate) = Panels per PO * Boards per Panel * Missing Rate:
                            // row bb, column 29(AC) = NumOfPanelInPO = Panels per PO
                            // row bb, column 36(AJ) = Qty By Panel = Boards per Panel
                            var num_of_panels = ToDoubleSafe(ws.Cell(bb, 29).Value);
                            var qty_of_pcb_per_panel = ToDoubleSafe(ws.Cell(bb, 36).Value);

                            if (materialMissRate == 0)
                            {
                                ws.Cell(bb, 37).Value = num_of_panels * qty_of_pcb_per_panel; //col AK: Qty By Total Panel
                            }
                            else
                            {
                                ws.Cell(bb, 37).Value = Math.Round(num_of_panels * qty_of_pcb_per_panel * (materialMissRate + 1) + 0.5, 0); //old program - 0.5 >> new pull list will be of qty + 1
                            }

                            // row bb, column 38(AL) = Storage Bin
                            ws.Cell(bb, 38).Value = GetGridCellAsString(dgvMaterialsOnTrayNonProgramMatrix, yy, 10);

                            // row bb, column 40(AN) = Std. REEL Qty
                            ws.Cell(bb, 40).Value = 0;

                            // Get PO order (which will run first):
                            string strPostAssyPONumber = ws.Cell(bb, 27).GetString();
                            int intOrderOfThisPO = 0;
                            // Loop through dgvPulledListPO and find the order to set intOrderOfThisPO:
                            for (tt = 0; tt < intNumOfAddedPO; tt++)
                            {
                                if (strPostAssyPONumber == GetGridCellAsString(dgvPulledListPO, tt, 1))
                                {
                                    intOrderOfThisPO = tt + 1;
                                    break;
                                }
                            }

                            // Find duplicate Item (Material) in list of POs (dgvMaterialsOnTrayNonProgramMatrix):
                            bool dupItem = false;
                            int dupRow = 0;
                            for (xx = rowProgrammingMaterials + 7; xx <= rowOnTrayMaterials + 1; xx++)
                            {
                                //column 46(AT - PulledList form - SIPLACE + (TRACK-Div)) = column 35 (AI - left side draft - SIPLACE + (TRACK-Div))
                                //column 44(AR - PulledList form - Materials) = column 31 (AE - left side draft - Materials)
                                if (ws.Cell(xx, 46) != null && !ws.Cell(xx, 46).IsEmpty() && 
                                ws.Cell(xx, 44) != null && !ws.Cell(xx, 44).IsEmpty() &&
                                ws.Cell(bb, 35) != null && !ws.Cell(bb, 35).IsEmpty() &&
                                ws.Cell(bb, 31) != null && !ws.Cell(bb, 31).IsEmpty())
                                {
                                    if (ws.Cell(xx, 46).GetString() == ws.Cell(bb, 35).GetString()
                                        && ws.Cell(xx, 44).GetString() == ws.Cell(bb, 31).GetString())
                                    {
                                        dupItem = true;
                                        dupRow = xx;
                                        break;
                                    }
                                }
                            }

                            // If no duplicate >> add new row (move to new materials); otherwise, keep adding...
                            if (!dupItem)
                            {
                                rowOnTrayMaterials++;
                                ws.Cell(rowOnTrayMaterials, 43).Value = rowOnTrayMaterials - rowProgrammingMaterials - 6;
                                ws.Cell(rowOnTrayMaterials, 44).Value = ws.Cell(bb, 31).Value;
                                ws.Cell(rowOnTrayMaterials, 45).Value = ws.Cell(bb, 38).Value.ToString() + GetMSL(ws.Cell(bb, 31).Value.GetText());
                                ws.Cell(rowOnTrayMaterials, 46).Value = ws.Cell(bb, 35).Value;
                                ws.Cell(rowOnTrayMaterials, 47 + intNumOfAddedPO).Value = ws.Cell(bb, 37).Value;
                                ws.Cell(rowOnTrayMaterials, 46 + intOrderOfThisPO).Value = ToDoubleSafe(ws.Cell(bb, 37).Value);
                            }
                            else // there's a duplicate row (material)
                            {
                                ws.Cell(dupRow, 47 + intNumOfAddedPO).Value = ToDoubleSafe(ws.Cell(bb, 37).Value) + ToDoubleSafe(ws.Cell(dupRow, 47 + intNumOfAddedPO).Value);
                                ws.Cell(dupRow, 46 + intOrderOfThisPO).Value = ToDoubleSafe(ws.Cell(dupRow, 46 + intOrderOfThisPO).Value) + ToDoubleSafe(ws.Cell(bb, 37).Value);
                            }
                        } // End of Fill in Excel file with information from dgvMaterialsOnTrayNonProgramMatrix - starts from row after last programming materials row + 6 spared rows, column Z

                        // Looping through the excel table just created to log the information (PO - PulledListID - Materials) that we have processed onto server vnmsrv601, table PulledListShiftvsSectorLog (3):
                        for (ii = 1; ii <= intNumOfAddedPO; ii++)
                        {
                            for (int jj = rowNormalMaterials + 7; jj <= rowProgrammingMaterials; jj++)
                            {
                                string siplace_and_trackdiv = ws.Cell(jj, 46).GetString();
                                int index_of_position_in_col_AT = InStrSafe(siplace_and_trackdiv, "(");
                                string poNumber = ws.Cell(3, 46 + ii).GetString();
                                string materialNumber = ws.Cell(jj, 44).GetString();
                                double qtyByPO = ToDoubleSafe(ws.Cell(jj, 46 + ii).Value);
                                siplace_value_from_grid = MidSafe(siplace_and_trackdiv, 2, index_of_position_in_col_AT - 3);
                                track_div_value_from_grid = MidSafe(siplace_and_trackdiv, index_of_position_in_col_AT + 2, siplace_and_trackdiv.Length - index_of_position_in_col_AT - 3);
                                if (qtyByPO > 0)
                                {
                                    LogPulledListvsSector(cbbPulledListLine.Text, cbbPulledListShift.Text, poNumber, materialNumber, qtyByPO, ii, pulledListID, _sumPOValue, siplace_value_from_grid, track_div_value_from_grid);
                                }
                            }
                        }

                        //----------------------------------------------------------------------------------------------------------------------------------------------------
                        //-----------------------------------------PREPARE DATA FOR BARTENDER LABELS PRINTING SECTION --------------------------------------------------------    
                        //----------------------------------------------------------------------------------------------------------------------------------------------------

                        // Clear old information on [DLVNDB].[dbo].[PulledListShiftControl]:
                        ClearPulledListIDLabel();

                        // Update information to [DLVNDB].[dbo].[PulledListShiftControl]:
                        for (int jj = rowProgrammingMaterials + 7; jj <= rowOnTrayMaterials; jj++)
                        {
                            string materialNumber = ws.Cell(jj, 44).GetString();
                            double pcbQty = ToDoubleSafe(ws.Cell(jj, 47 + intNumOfAddedPO).Value);
                            string poNumber = GetPONumberFrom_dgvMultiUniPhysicalModelPulled(materialNumber);
                            string series = ws.Cell(2, 43).GetString();
                            if (pcbQty > 0)
                            {
                                if (materialNumber.Substring(0, 3) == "100")
                                {
                                    InsertDataForBartenderLabels(cbbPulledListLine.Text, pulledListID, _strUserName, series, poNumber, materialNumber, pcbQty);
                                }
                            }
                        }

                        // Create txt files on server to trigger Bartender to print labels:
                        //DialogResult printLabels = MessageBox.Show("Do you want to print labels for the pulled list?", "Print Labels Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        //if (printLabels == DialogResult.Yes)
                        //{
                        //    CreateTxtFilesToTriggerBartender();
                        //    MessageBox.Show("Labels print job has been triggered. Please check the printer.", "Print Labels", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //}
                        if (chkbLabelsPrint.Checked == true)
                        {
                            CreateTxtFilesToTriggerBartender();
                            //MessageBox.Show("Labels print job has been triggered. Please check the printer.", "Print Labels", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        //----------------------------------------------------------------------------------------------------------------------------------------------------
                        //----------------------------------------------------FORMAT EXCEL FILE & EXPORT SECTION -------------------------------------------------------------    
                        //----------------------------------------------------------------------------------------------------------------------------------------------------


                        // ----------------Format table headers for linh kien thuong----------------
                        //--------------------------------------------------------------------------

                        // Add headers (hidden columns A:AP)
                        ws.Cell(3, 26).SetValue("No."); //Excel file Row 3, Column Z
                        ws.Cell(3, 27).SetValue("PONumber");//Excel file Row 3, Column AA
                        ws.Cell(3, 28).Value = "Model";//Excel file Row 3, Column AB
                        ws.Cell(3, 29).Value = "Num Of Panel";
                        ws.Cell(3, 30).Value = "TOP ? BOT";
                        ws.Cell(3, 31).Value = "Materials";
                        ws.Cell(3, 32).Value = "Qty On PO";
                        ws.Cell(3, 33).Value = "SIPLACE";
                        ws.Cell(3, 34).Value = "TRACK-Div";
                        ws.Cell(3, 35).Value = "SIPLACE + (TRACK-Div)";
                        ws.Cell(3, 36).Value = "Qty By Panel";
                        ws.Cell(3, 37).Value = "Qty By Total Panel";
                        ws.Cell(3, 38).Value = "Storage Bin"; //Excel file Row 3, Column AL
                        ws.Cell(3, 39).Value = "Actual PickUp";
                        ws.Cell(3, 40).Value = "Std. REEL Qty";
                        ws.Cell(3, 41).Value = "Num Of Std. REEL";
                        ws.Cell(3, 43).Value = "No.";
                        ws.Cell(3, 44).Value = "Materials";
                        ws.Cell(3, 45).Value = "Storage Bin + (MSL)"; //Excel file Row 3, Column AS
                        ws.Cell(3, 46).Value = "SIPLACE + (TRACK-Div)"; //Excel file Row 3, Column AT

                        // Depends on how many PO are there, these columns will be right after:
                        ws.Cell(3, 47 + intNumOfAddedPO).Value = "Qty Of Materials";
                        ws.Cell(3, 48 + intNumOfAddedPO).Value = "Actual PickUp";
                        ws.Cell(3, 49 + intNumOfAddedPO).Value = "...";

                        // Color & style for header block from colZ to AO:
                        ws.Range(3, 26, 3, 41).Style.Fill.BackgroundColor = XLColor.FromHtml("#FF00FFFF"); //cyan
                        ws.Range(3, 26, 3, 41).Style.Font.Bold = true;

                        // Call GetSeries and paste the value to row 2, column AQ:
                        ws.Cell(2, 43).Value = GetSeries(GetGridCellAsString(dgvPulledListPO, 1, 2)) + " Series";

                        // Add Pulled List Info to row 1, column AQ:
                        ws.Cell(1, 43).Value = "'- Pulled List Line: " + cbbPulledListLine.Text
                            + "\n- Pulled List ID: " + pulledListID + "(" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + ")"
                            + "\n- Pulled List User: " + _strUserName;

                        // Pulled list info row styling row 1 column 43..(49 + intNumOfAddedPO):
                        var pullInfoRange = ws.Range(1, 43, 1, 49 + intNumOfAddedPO);
                        pullInfoRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        pullInfoRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        pullInfoRange.Style.Fill.PatternType = XLFillPatternValues.Solid;
                        pullInfoRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#FF006400");//dark-green
                        pullInfoRange.Style.Font.FontColor = XLColor.FromHtml("#FFFFFFFF");//white
                        pullInfoRange.Style.Font.Bold = true;
                        pullInfoRange.Merge();
                        pullInfoRange.Style.Alignment.WrapText = true;

                        // Merge row - Series - Danh sach linh kien thuong table:
                        var seriesRange = ws.Range(2, 43, 2, 49 + intNumOfAddedPO);
                        seriesRange.Style.Font.FontSize = 15;
                        seriesRange.Style.Fill.PatternType = XLFillPatternValues.Solid;
                        seriesRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFFFF00");//yellow
                        seriesRange.Style.Font.FontColor = XLColor.FromHtml("#FF000000");//black
                        seriesRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        seriesRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        seriesRange.Style.Font.Bold = true;
                        seriesRange.Merge();

                        // Color & style for header block from colAQ to ...(49 + intNumOfAddedPO):
                        ws.Range(3, 43, 3, 46).Style.Fill.BackgroundColor = XLColor.FromHtml("#FF00FFFF");//cyan
                        ws.Range(3, 43, 3, 49 + intNumOfAddedPO).Style.Font.Bold = true;

                        // Style Qty of Materials, Actual PickUp, ... columns headers:
                        var sumUpColumnsRange = ws.Range(3, 47 + intNumOfAddedPO, 3, 49 + intNumOfAddedPO);
                        sumUpColumnsRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#FF90EE90");//light-green
                        sumUpColumnsRange.Style.Font.Bold = true;

                        //--------------------------------------------------------------------------

                        // Copy value of row 2 - Series - to DANH SACH LINH KIEN NAP ROM table:
                        ws.Cell(rowNormalMaterials + 5, 43).Value = ws.Cell(2, 43).Value;

                        // Copy row 2 - Series - to DANH SACH LINH KIEN CHUAN BI THEO TRAY table:
                        ws.Cell(rowProgrammingMaterials + 5, 43).Value = ws.Cell(2, 43).Value;



                        //----------Format table headers for DANH SACH LINH KIEN NAP ROM-----------
                        //--------------------------------------------------------------------------

                        // Set title text "DANH SACH LINH KIEN NAP ROM":
                        ws.Cell(rowNormalMaterials + 3, 43).Value = "DANH SACH LINH KIEN NAP ROM";

                        // Merge row "DANH SACH LINH KIEN NAP ROM" (row rowNormalMaterials + 3):
                        var titleRange1 = ws.Range(rowNormalMaterials + 3, 43, rowNormalMaterials + 3, 49 + intNumOfAddedPO);
                        titleRange1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        titleRange1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        titleRange1.Style.Fill.PatternType = XLFillPatternValues.Solid;
                        titleRange1.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFF0000");//red
                        titleRange1.Style.Font.FontColor = XLColor.FromHtml("#FFFFFFFF");//white
                        titleRange1.Style.Font.Bold = true;
                        titleRange1.Merge();

                        // Set the multiline pulled list info (use \n for new lines and enable wrap)
                        ws.Cell(rowNormalMaterials + 4, 43).Value = $"- Pulled List Line: {cbbPulledListLine.Text}" +
                            $"\n- Pulled List ID: {pulledListID} ({System.DateTime.Now:dd/MM/yyyy HH:mm:ss})" +
                            $"\n- Pulled List User: {_strUserName}";
                        //ws.Cell(rowNormalMaterials + 4, 43).Style.Alignment.WrapText = true;

                        // Pulled list info row styling (rowNormalMaterials + 4) "DANH SACH LINH KIEN NAP ROM":
                        var pullInfoRange1 = ws.Range(rowNormalMaterials + 4, 43, rowNormalMaterials + 4, 49 + intNumOfAddedPO);
                        pullInfoRange1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        pullInfoRange1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        pullInfoRange1.Style.Fill.PatternType = XLFillPatternValues.Solid;
                        pullInfoRange1.Style.Fill.BackgroundColor = XLColor.FromHtml("#FF006400");//dark-green
                        pullInfoRange1.Style.Font.FontColor = XLColor.FromHtml("#FFFFFFFF");//white
                        pullInfoRange1.Style.Font.Bold = true;
                        pullInfoRange1.Merge();
                        pullInfoRange1.Style.Alignment.WrapText = true;

                        // Merge row - Series - DANH SACH LINH KIEN NAP ROM table:
                        var seriesRange1 = ws.Range(rowNormalMaterials + 5, 43, rowNormalMaterials + 5, 49 + intNumOfAddedPO);
                        seriesRange1.Style.Font.FontSize = 15;
                        seriesRange1.Style.Fill.PatternType = XLFillPatternValues.Solid;
                        seriesRange1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        seriesRange1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        seriesRange1.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFFFF00");//yellow
                        seriesRange1.Style.Font.FontColor = XLColor.FromHtml("#FF000000");//black
                        seriesRange1.Style.Font.Bold = true;
                        seriesRange1.Merge();

                        // Headers of "DANH SACH LINH KIEN NAP ROM" - row rowNormalMaterials + 6:
                        ws.Cell(rowNormalMaterials + 6, 43).Value = "No.";
                        ws.Cell(rowNormalMaterials + 6, 44).Value = "Materials";
                        ws.Cell(rowNormalMaterials + 6, 45).Value = "Storage Bin + (MSL)";
                        ws.Cell(rowNormalMaterials + 6, 46).Value = "SIPLACE + (TRACK-Div)";

                        // Add columns corresponding to each PO to headers of "DANH SACH LINH KIEN NAP ROM":
                        for (xx = 0; xx < intNumOfAddedPO; xx++)
                        {
                            string postAssyPONumber = GetGridCellAsString(dgvPulledListPO, xx, 1);
                            string poModel = GetGridCellAsString(dgvPulledListPO, xx, 2);
                            string poSide = GetGridCellAsString(dgvPulledListPO, xx, 3);
                            // Place into cell (rowNormalMaterials+6, 46 + xx)
                            var cell = ws.Cell(rowNormalMaterials + 6, 46 + xx + 1);
                            cell.Value = $"{postAssyPONumber}\n{poModel}\n({poSide})";
                            cell.Style.Alignment.WrapText = true;
                            cell.Style.Font.Bold = true;
                        }

                        // Add sum up columns:
                        ws.Cell(rowNormalMaterials + 6, 47 + intNumOfAddedPO).Value = "Qty Of Materials";
                        ws.Cell(rowNormalMaterials + 6, 48 + intNumOfAddedPO).Value = "Actual PickUp";
                        ws.Cell(rowNormalMaterials + 6, 49 + intNumOfAddedPO).Value = "...";

                        // Style the header range: set text to bold, background to dark-green, enable wrap text:
                        var headerRange1 = ws.Range(rowNormalMaterials + 6, 43, rowNormalMaterials + 6, 46);
                        headerRange1.Style.Fill.BackgroundColor = XLColor.FromHtml("#FF00FFFF"); //cyan
                        headerRange1.Style.Alignment.WrapText = true;
                        headerRange1.Style.Font.Bold = true;

                        // Style the Qty of Materials, Actual PickUp, ... columns:
                        var sumUpColumnsRange1 = ws.Range(rowNormalMaterials + 6, 47 + intNumOfAddedPO, rowNormalMaterials + 6, 49 + intNumOfAddedPO);
                        sumUpColumnsRange1.Style.Fill.BackgroundColor = XLColor.FromHtml("#FF90EE90");//light-green
                        sumUpColumnsRange1.Style.Font.Bold = true;

                        //------Format table headers for DANH SACH LINH KIEN CHUAN BI THEO TRAY-----
                        //--------------------------------------------------------------------------

                        // Merge row - Series - DANH SACH LINH KIEN CHUAN BI THEO TRAY table:
                        var seriesRange2 = ws.Range(rowProgrammingMaterials + 5, 43, rowProgrammingMaterials + 5, 49 + intNumOfAddedPO);
                        //seriesRange2.Style.Fill.BackgroundColor = XLColor.Yellow;
                        //seriesRange2.Style.Font.SetBold();
                        //seriesRange2.Merge();
                        seriesRange2.Style.Font.FontSize = 15;
                        seriesRange2.Style.Fill.PatternType = XLFillPatternValues.Solid;
                        seriesRange2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        seriesRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        seriesRange2.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFFFF00");//yellow
                        seriesRange2.Style.Font.FontColor = XLColor.FromHtml("#FF000000");//black
                        seriesRange2.Style.Font.Bold = true;
                        seriesRange2.Merge();

                        // Set title text "DANH SACH LINH KIEN CHUAN BI THEO TRAY":
                        var titleRange2 = ws.Range(rowProgrammingMaterials + 3, 43, rowProgrammingMaterials + 3, 49 + intNumOfAddedPO);
                        titleRange2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        titleRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        titleRange2.Style.Fill.BackgroundColor = XLColor.Red;
                        titleRange2.Style.Font.FontColor = XLColor.Black; // theme dark1 not available; use black
                        titleRange2.Style.Font.SetBold();
                        titleRange2.Merge();
                        ws.Cell(rowProgrammingMaterials + 3, 43).Value = "DANH SACH LINH KIEN CHUAN BI THEO TRAY";

                        // Pulled list info row styling (rowProgrammingMaterials + 4) "DANH SACH LINH KIEN CHUAN BI THEO TRAY":
                        var pullInfoRange2 = ws.Range(rowProgrammingMaterials + 4, 43, rowProgrammingMaterials + 4, 49 + intNumOfAddedPO);
                        pullInfoRange2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        pullInfoRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        pullInfoRange2.Style.Fill.PatternType = XLFillPatternValues.Solid;
                        pullInfoRange2.Style.Fill.BackgroundColor = XLColor.FromHtml("#FF006400");//dark-green
                        pullInfoRange2.Style.Font.FontColor = XLColor.FromHtml("#FFFFFFFF");//white
                        pullInfoRange2.Style.Font.SetBold();
                        pullInfoRange2.Merge();

                        // Set the multiline pulled list info (use \n for new lines and enable wrap):
                        ws.Cell(rowProgrammingMaterials + 4, 43).Value = $"- Pulled List Line: {cbbPulledListLine.Text}" +
                            $"\n- Pulled List ID: {pulledListID} ({System.DateTime.Now:dd/MM/yyyy HH:mm:ss})" +
                            $"\n- Pulled List User: {_strUserName}";
                        ws.Cell(rowProgrammingMaterials + 4, 43).Style.Alignment.WrapText = true;

                        // Header names row (rowProgrammingMaterials + 6):
                        ws.Cell(rowProgrammingMaterials + 6, 43).Value = "No.";
                        ws.Cell(rowProgrammingMaterials + 6, 44).Value = "Materials";
                        ws.Cell(rowProgrammingMaterials + 6, 45).Value = "Storage Bin + (MSL)";
                        ws.Cell(rowProgrammingMaterials + 6, 46).Value = "SIPLACE + (TRACK-Div)";

                        // Add columns corresponding to each PO:
                        for (xx = 0; xx < intNumOfAddedPO; xx++)
                        {
                            string postAssyPONumber = GetGridCellAsString(dgvPulledListPO, xx, 1);
                            string poModel = GetGridCellAsString(dgvPulledListPO, xx, 2);
                            string poSide = GetGridCellAsString(dgvPulledListPO, xx, 3);
                            // Place into cell (rowProgrammingMaterials + 6, 46 + xx)
                            var cell = ws.Cell(rowProgrammingMaterials + 6, 46 + xx + 1);
                            cell.Value = $"{postAssyPONumber}\n{poModel}\n({poSide})";
                            cell.Style.Alignment.WrapText = true;
                            cell.Style.Font.Bold = true;
                        }

                        // Add sum up columns:
                        ws.Cell(rowProgrammingMaterials + 6, 47 + intNumOfAddedPO).Value = "Qty Of Materials"; // col47 = AU
                        ws.Cell(rowProgrammingMaterials + 6, 48 + intNumOfAddedPO).Value = "Actual PickUp";
                        ws.Cell(rowProgrammingMaterials + 6, 49 + intNumOfAddedPO).Value = "...";

                        // Color & style for header block from colAQ to ...(49 + intNumOfAddedPO):
                        // Style the header range: set text to bold, background to dark-green, enable wrap text:
                        var headerRange2 = ws.Range(rowProgrammingMaterials + 6, 43, rowProgrammingMaterials + 6, 46);
                        headerRange2.Style.Fill.BackgroundColor = XLColor.FromHtml("#FF00FFFF"); //cyan
                        headerRange2.Style.Alignment.WrapText = true;
                        headerRange2.Style.Font.Bold = true;

                        // Style Qty of Materials, Actual PickUp, ... columns headers:
                        var sumUpColumnsRange2 = ws.Range(rowProgrammingMaterials + 6, 47 + intNumOfAddedPO, rowProgrammingMaterials + 6, 49 + intNumOfAddedPO);
                        sumUpColumnsRange2.Style.Fill.BackgroundColor = XLColor.FromHtml("#FF90EE90");//light-green
                        sumUpColumnsRange2.Style.Font.Bold = true;


                        //-----------------------------Format borders-------------------------------
                        //--------------------------------------------------------------------------

                        // Set border for first table: Normal materials:
                        var borderRange1 = ws.Range(1, 43, rowNormalMaterials, 49 + intNumOfAddedPO);
                        borderRange1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        borderRange1.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        borderRange1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Set border for second table: Programming required materials - "DANH SACH LINH KIEN NAP ROM":
                        var borderRange2 = ws.Range(rowNormalMaterials + 4, 43, rowProgrammingMaterials, 49 + intNumOfAddedPO);
                        borderRange2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        borderRange2.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        borderRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Set border for third table: On tray materials - "DANH SACH LINH KIEN CHUAN BI THEO TRAY":
                        var borderRange3 = ws.Range(rowProgrammingMaterials + 4, 43, rowOnTrayMaterials, 49 + intNumOfAddedPO);
                        borderRange3.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        borderRange3.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        borderRange3.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;



                        // --------------------------Set color for cells----------------------------
                        //--------------------------------------------------------------------------

                        // "AM3:AM{rowMaterials}"
                        var range1 = ws.Range($"AM3:AM{rowProgrammingMaterials}");
                        range1.Style.Fill.SetBackgroundColor(XLColor.Yellow);
                        range1.Style.Font.SetBold();
                        range1.Style.Font.FontColor = XLColor.Blue;
                        range1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        range1.Style.NumberFormat.Format = "#,##0_);(#,##0)";

                        // AE3:AE{rowMaterials}
                        var range2 = ws.Range($"AE2:AE{rowProgrammingMaterials}");
                        range2.Style.Fill.SetBackgroundColor(XLColor.Yellow);
                        range2.Style.Font.SetBold();
                        range2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                        // AI3:AI{rowMaterials}
                        var range3 = ws.Range($"AI2:AI{rowProgrammingMaterials}");
                        range3.Style.Fill.SetBackgroundColor(XLColor.Yellow);
                        range3.Style.Font.SetBold();
                        range3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                        // AK3:AK{rowMaterials}
                        var range4 = ws.Range($"AK2:AK{rowProgrammingMaterials}");
                        range4.Style.Fill.SetBackgroundColor(XLColor.Yellow);
                        range4.Style.Font.SetBold();
                        range4.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        range4.Style.NumberFormat.Format = "#,##0_);(#,##0)";



                        // Column(47+intNumOfAddedPO) - Qty Of Materials
                        if (rowNormalMaterials > 4)
                        {
                            var range5a = ws.Range(ws.Cell(4, 47 + intNumOfAddedPO), ws.Cell(rowNormalMaterials, 47 + intNumOfAddedPO));
                            range5a.Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            range5a.Style.NumberFormat.Format = "#,##0_);(#,##0)";
                            range5a.Style.Font.SetBold();
                            range5a.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        }

                        if (rowProgrammingMaterials > rowNormalMaterials + 7)
                        {
                            var range5b = ws.Range(ws.Cell(rowNormalMaterials + 7, 47 + intNumOfAddedPO), ws.Cell(rowProgrammingMaterials, 47 + intNumOfAddedPO));
                            range5b.Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            range5b.Style.NumberFormat.Format = "#,##0_);(#,##0)";
                            range5b.Style.Font.SetBold();
                            range5b.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        }

                        if (rowOnTrayMaterials > rowProgrammingMaterials + 7)
                        {
                            var range5c = ws.Range(ws.Cell(rowProgrammingMaterials + 7, 47 + intNumOfAddedPO), ws.Cell(rowOnTrayMaterials, 47 + intNumOfAddedPO));
                            range5c.Style.Fill.SetBackgroundColor(XLColor.Yellow);
                            range5c.Style.NumberFormat.Format = "#,##0_);(#,##0)";
                            range5c.Style.Font.SetBold();
                            range5c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        }

                        //// Column (48+intNumOfAddedPO) - Actual PickUp
                        //var range6 = ws.Range(ws.Cell(4, 48 + intNumOfAddedPO), ws.Cell(rowProgrammingMaterials - 6, 48 + intNumOfAddedPO));
                        //range6.Style.Fill.SetBackgroundColor(XLColor.Yellow);
                        //range6.Style.Font.SetBold();
                        //range6.Style.Font.FontColor = XLColor.Blue;
                        //range6.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        //range6.Style.NumberFormat.Format = "#,##0_);(#,##0)";

                        // Range (4,47) to (rowNormalMaterials, 47+intNumOfAddedPO)
                        var range7a = ws.Range(ws.Cell(4, 47), ws.Cell(rowNormalMaterials, 47 + intNumOfAddedPO - 1)); //col 47 = AU; ex 3 PO added >> col 47, 48, 49 >> 47 + 3 - 1 = 49
                        range7a.Style.Fill.BackgroundColor = XLColor.White;
                        range7a.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        range7a.Style.NumberFormat.Format = "#,##0_);(#,##0)";

                        var range7b = ws.Range(ws.Cell(rowNormalMaterials + 7, 47), ws.Cell(rowProgrammingMaterials, 47 + intNumOfAddedPO - 1));
                        range7b.Style.Fill.BackgroundColor = XLColor.White;
                        range7b.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        range7b.Style.NumberFormat.Format = "#,##0_);(#,##0)";

                        var range7c = ws.Range(ws.Cell(rowProgrammingMaterials + 7, 47), ws.Cell(rowOnTrayMaterials, 47 + intNumOfAddedPO - 1));
                        range7c.Style.Fill.BackgroundColor = XLColor.White;
                        range7c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        range7c.Style.NumberFormat.Format = "#,##0_);(#,##0)";

                    } // end if SMT_Group or POSTASSY_Group

                    // Set dgvMaterialsRarDivideKANBANBox to excel file:
                    for (tt = 0; tt < dgvMaterialsRarDivideKANBANBox.Rows.Count; tt++)
                    {
                        if (strIsKittingActive == "YES")
                        {
                            intNumOfKBDivided = Convert.ToInt32(GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 9));
                            dblRoundToQty = Math.Floor(Convert.ToDouble(GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 5)) / intNumOfKBDivided);
                            for (hh = 1; hh <= intNumOfKBDivided - 1; hh++)
                            {
                                m++;
                                ws.Cell(m, 16).Value = m;
                                ws.Cell(m, 17).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 1);
                                ws.Cell(m, 18).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 2);
                                ws.Cell(m, 19).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 3);
                                ws.Cell(m, 20).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 4);
                                ws.Cell(m, 21).Value = dblRoundToQty;
                                ws.Cell(m, 22).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 6);
                                ws.Cell(m, 23).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 7);
                                ws.Cell(m, 24).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 8);
                            }

                            m++;
                            ws.Cell(m, 16).Value = m;
                            ws.Cell(m, 17).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 1);
                            ws.Cell(m, 18).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 2);
                            ws.Cell(m, 19).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 3);
                            ws.Cell(m, 20).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 4);
                            ws.Cell(m, 21).Value = Convert.ToDouble(GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 5)) - (intNumOfKBDivided - 1) * dblRoundToQty;
                            ws.Cell(m, 22).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 6);
                            ws.Cell(m, 23).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 7);
                            ws.Cell(m, 24).Value = GetGridCellAsString(dgvMaterialsRarDivideKANBANBox, tt, 8);
                        }
                    }
                            
                    // Set global font for all cells:
                    ws.Style.Font.FontName = "Times New Roman";

                    // Row heights for specific rows:
                    ws.Row(1).Height = 59;
                    ws.Row(2).Height = 29;
                    ws.Row(3).Height = 50;

                    ws.Row(rowNormalMaterials + 4).Height = 59;
                    ws.Row(rowNormalMaterials + 5).Height = 29;
                    ws.Row(rowNormalMaterials + 6).Height = 50;

                    ws.Row(rowProgrammingMaterials + 4).Height = 59;
                    ws.Row(rowProgrammingMaterials + 5).Height = 29;
                    ws.Row(rowProgrammingMaterials + 6).Height = 50;

                    // Set various column widths (using column letters)
                    ws.Column("AC").Width = 5;
                    ws.Column("AD").Width = 5;
                    ws.Column("AF").Width = 6;
                    ws.Column("AJ").Width = 5;
                    ws.Column("AK").Width = 6;
                    ws.Column("AN").Width = 6;
                    ws.Column("AO").Width = 5;
                    ws.Column("AR").Width = 16.5;
                    ws.Column("AT").Width = 15.5;
                    ws.Column("AU").Width = 9;

                    // Set widths for PO columns (48..47+intNumOfAddedPO) to 12
                    for (int c = 48; c <= 47 + intNumOfAddedPO; c++)
                    {
                        ws.Column(c).Width = 12;
                    }
                        
                    // Next single column (48 + intNumOfAddedPO) set width = 9
                    ws.Column(48 + intNumOfAddedPO).Width = 9;

                    // Auto fit columns Z:BZ
                    ws.Columns("Z:BZ").AdjustToContents();

                    // Align some columns center
                    ws.Column("Z").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Column("AG").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Column("AH").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Column("AD").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    
                    ws.Columns("A:AO").Hide();
                    //ws.Column("G").Hide();
                    //ws.Columns("I:M").Hide();

                    // Add IMPORTANT NOTE lines starting at getLastPCBARowInExcel + 2
                    int getNote1Row = 2;
                    ws.Cell(getNote1Row, 1).Value = "IMPORTANT NOTE: TRUOC KHI DUNG T-CODE ZLP14 DE GOI THE PULL LIST PO";
                    ws.Cell(getNote1Row + 1, 2).Value = "'+ PHAI XOA CAC DONG CO VI TRI PCBA BIN";
                    ws.Cell(getNote1Row + 2, 2).Value = "'+ PHAI KIEM TRA SO TO BANG VOI SO LUONG PART CAN GOI VI TRONG TRUONG HOP";
                    ws.Cell(getNote1Row + 3, 3).Value = "MOT PART MOI XUAT HIEN VA CHUA KHAI BAO TRONG HE THONG THI LENH T-CODE ZLP14 SE BO QUA TO!";

                    // Print area setup
                    var printAreaRange = ws.Range(1, 43, rowOnTrayMaterials, 49 + intNumOfAddedPO);
                    ws.PageSetup.PrintAreas.Clear();
                    ws.PageSetup.PrintAreas.Add(printAreaRange.RangeAddress.ToString(XLReferenceStyle.A1));
                    ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    ws.PageSetup.SetRowsToRepeatAtTop(1, 3);

                    // After building workbook, save to file and open it for user
                    outputFile = SaveWorkbookAndOpen(workbook, pulledListLineText);

                    dgvMultiUniPhysicalModelPulled.Sort(dgvMultiUniPhysicalModelPulled.Columns[3], ListSortDirection.Ascending);
                }
            }
            catch (Exception ex)
            {
                //UpdateProgramRunningBug("ExportPulledListFromServer", ex.HResult.ToString(), ex.Message);
                MessageBox.Show(ex.Message, "Exception caught:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static double ToDoubleSafe(object o)
        {
            if (o == null) return 0;
            if (o is double) return (double)o;
            if (o is int) return Convert.ToDouble(o);
            double d;
            if (double.TryParse(Convert.ToString(o, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return d;
            return 0;
        }

        private static int InStrSafe(string source, string find)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(find)) return 0;
            int idx = source.IndexOf(find, StringComparison.Ordinal);
            return (idx >= 0) ? idx + 1 : 0;
        }

        private static string MidSafe(string s, int start, int length)
        {
            if (s == null) return string.Empty;
            if (start < 1) start = 1;
            if (start > s.Length) return string.Empty;
            int zeroBasedStart = start - 1;
            if (zeroBasedStart + length > s.Length) length = s.Length - zeroBasedStart;
            return s.Substring(zeroBasedStart, Math.Max(0, length));
        }

        private static int FlexSortGenericAscending() => 1;  // placeholder
        private static int FlexSortStringDescending() => 2;  // placeholder

        private static string SaveWorkbookAndOpen(XLWorkbook workbook, string pulledListLineText)
        {
            string dir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MPH_PullLists");
            Directory.CreateDirectory(dir);
            string fileName = $"{NormalizeFileName(pulledListLineText)}_PullList_{System.DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string fullPath = System.IO.Path.Combine(dir, fileName);
            workbook.SaveAs(fullPath);
            // Try to open the file in default application
            try
            {
                Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
            }
            catch { /* ignore if cannot open */ }
            return fullPath;
        }

        private static string NormalizeFileName(string s)
        {
            foreach (var c in System.IO.Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');
            return s;
        }

        private void RC_Clear_Click(object sender, EventArgs e)
        {
            dgvCommonPartPO.Rows.Clear();
            dgvCrossPlanningDone.Rows.Clear();
            dgvKANBANPulledList.Rows.Clear();
            dgvLocalPart.Rows.Clear();
            dgvMaterialsConversionMatrix.Rows.Clear();
            dgvMaterialsOnTrayNonProgramMatrix.Rows.Clear();
            dgvMaterialsPONumber.Rows.Clear();
            dgvMaterialsRarDivideKANBANBox.Rows.Clear();
            dgvMultiUniPhysicalModelPulled.Rows.Clear();
            dgvPartFirstPO.Rows.Clear();
            dgvPartPCBAOfPO.Rows.Clear();
            dgvPartRestOfPO.Rows.Clear();
            dgvPartvsQty.Rows.Clear();
            dgvPhysicalModelAfterCOPulled.Rows.Clear();
            dgvPhysicalModelRunningPulled.Rows.Clear();
            dgvPhysicalSAPModelAfterCOPulled.Rows.Clear();
            dgvPLOverallModelPulled.Rows.Clear();
            dgvPLPhysicalModelPulled.Rows.Clear();
            dgvPotentialIssues.Rows.Clear();
            dgvPulledListPO.Rows.Clear();
            dgvUniPOnQtyMaterials.Rows.Clear();
            dgvUniPhysicalModelPulled.Rows.Clear();
            dgvPullListvsPO2.Rows.Clear();
            dgvPullListvsPO.Rows.Clear();
            dgvQtyvsCountDuplicated.Rows.Clear();
            UnlockSelection();
        }

        private void chkbLabelsPrint_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.blnLabelsPrint = chkbLabelsPrint.Checked;
            Properties.Settings.Default.Save();
        }
    }
}
