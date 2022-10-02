//#define DEBUG_SHOW_MESSAGEBOX

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ChuongTrinhDiemDanh_DHThaiNguyen.Classes;
using ChuongTrinhDiemDanh_DHThaiNguyen.RJControls;
using Google.Apis.Sheets.v4.Data;

namespace ChuongTrinhDiemDanh_DHThaiNguyen
{
    public partial class MainForm : Form
    {

        #region global variables
        // Google Sheets API
        GoogleSheet spreadSheet = new GoogleSheet();
        string sheetName;
        string sheetRange_data;
        string sheetRange_total;
        string sheetRange_attendance;
        int sheetVar_updateRangeOffset;
        //string gSheet_
        IList<IList<object>> sheetData_data;
        IList<IList<object>> sheetData_total;

        string path_images;
        string[] allImagePaths;

        #endregion

        #region debug functions
        private void DebugConsolePrint(object message,
    [System.Runtime.CompilerServices.CallerMemberName] string memberName = "")
        {
#if DEBUG
            System.Diagnostics.Trace.WriteLine($"'{memberName}'" + "\t->\t" + message.ToString());
#endif
        }

        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Decription, int Reserved);
        #endregion

        #region user-defined functions
        private void SetAppSettingValue(string key, string value)
        {
            // Open App.Config of executable
            System.Configuration.Configuration config =
             ConfigurationManager.OpenExeConfiguration
                        (ConfigurationUserLevel.None);

            // Add an Application Setting.
            config.AppSettings.Settings[key].Value = value;

            // Save the changes in App.config file.
            config.Save(ConfigurationSaveMode.Modified);

            // Force a reload of a changed section.
            ConfigurationManager.RefreshSection("appSettings");
        }
        // Init data
        private void InitData()
        {
            sheetName = ConfigurationManager.AppSettings["google_sheet_sheet_name"];
            sheetRange_data = ConfigurationManager.AppSettings["google_sheet_data_range"];
            sheetRange_total = ConfigurationManager.AppSettings["google_sheet_total_range"];
            sheetRange_attendance = ConfigurationManager.AppSettings["google_sheet_attendance_range"];

            sheetVar_updateRangeOffset = Int32.Parse(ConfigurationManager.AppSettings["google_sheet_update_range_offset"]);

            path_images = ConfigurationManager.AppSettings["path_images"];

            string rjToggle_checking_status = ConfigurationManager.AppSettings["VAR_rjToggle_checking_status"];
            rjToggleButton_CheckIn.CheckState = (rjToggle_checking_status == "1" ? CheckState.Checked : CheckState.Unchecked);

            allImagePaths = LoadImagePaths(path_images);
        }
        // Check Internet connection        
        public bool CheckInternetConnection()
        {
            int Desc;
            return InternetGetConnectedState(out Desc, 0);
        }
        // Google Sheet API
        private bool GoogleSheet_Auth()
        {
            string credentialPath = ConfigurationManager.AppSettings["google_sheet_credential_path"];
            string spreadSheetID = ConfigurationManager.AppSettings["google_sheet_id"];
            return spreadSheet.Begin(credentialPath, spreadSheetID, "DIEM DANH DAI HOI - DH THAI NGUYEN");
        }
        private bool UpdateAttendanceStatus(int index)
        {
            try
            {
                // Create ValueRange...
                ValueRange valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { new List<object>()
                            { rjToggleButton_CheckIn.Checked ? "false" : "true" /*Out:In*/} };
                string range = $"{sheetName}!E{index + sheetVar_updateRangeOffset}";
                DebugConsolePrint("Update range: " + range);
                return spreadSheet.WriteData(range, valueRange);
            }
            catch (Exception ex)
            {
#if DEBUG_SHOW_MESSAGEBOX
                MessageBox.Show(ex.Message, ex.GetType().Name);
#else
                DebugConsolePrint(ex.Message);
#endif
            }
            return false;
        }

        // Get data from gSheet

        // Load image paths
        public string[] LoadImagePaths(string path)
        {
            string[] allImagePaths = Directory.GetFiles(path_images, "*.*", SearchOption.AllDirectories);
            foreach (var item in allImagePaths)
            {
                DebugConsolePrint(item);
            }
            DebugConsolePrint($"Amount of images: {allImagePaths.Length}");
            return allImagePaths;
        }
        // Update Profile Image
        public void LoadErrorImage()
        {
            string error_img_path = ConfigurationManager.AppSettings["path_image_error"];
            circularPictureBox_Profile.Image = Image.FromFile(error_img_path);
        }
        public void LoadProfileImage(string code)
        {
            bool isFound = false;
            try
            {
                foreach (var path in allImagePaths)
                {
                    if (path.Contains(code))
                    {
                        circularPictureBox_Profile.Image = Image.FromFile(path);
                        isFound = true;
                        break;
                    }
                }
                if (!isFound)
                {
                    LoadErrorImage();
                }
            }
            catch (Exception ex)
            {
                LoadErrorImage();
#if DEBUG_SHOW_MESSAGEBOX
                MessageBox.Show(ex.Message, ex.GetType().Name);
#else
                DebugConsolePrint(ex.Message);
#endif
            }
        }

        // Threading
        delegate void SetTextCallback(string text);
        private void SetThreadTextbox(string text)
        {
            if (this.label_TotalAttendance.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetThreadTextbox);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.label_TotalAttendance.Text = text;
            }
        }
        private void ThreadUpdateTotalAttendance()
        {
            int updateDelay = Int32.Parse(ConfigurationManager.AppSettings["VAR_total_attendance_delay_time"]);
            DebugConsolePrint($"Update delay: {updateDelay}");

            while (true)
            {
                if (CheckInternetConnection())
                {
                    Thread.Sleep(updateDelay);

                    try
                    {
                        // Get batch values (totalAttendees + sheetData)
                        ValueRange valueRanges = spreadSheet.ReadData(sheetRange_total);

                        // Check: ignore if it hasn't change
                        if (valueRanges != null)
                        {
                            string total_checkin = (valueRanges.Values[0][1] == null ? "NULL" : valueRanges.Values[0][1].ToString());
                            string total = (valueRanges.Values[0][0] == null ? "NULL" : valueRanges.Values[0][0].ToString());
                            string percent = (valueRanges.Values[0][2] == null ? "NULL" : valueRanges.Values[0][2].ToString());

                            SetThreadTextbox($"Số Đại biểu: {total_checkin}/{total}" +
                                            $"\n({percent})");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugConsolePrint(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng kiểm tra lại đường truyền mạng. Chương trình sẽ tiếp tục khi có internet", "Cảnh báo");
                    Thread.Sleep(1000);
                }
            }
        }

        // Layout
        public void InitLayoutLocation()
        {
            string path_background_image = ConfigurationManager.AppSettings["path_image_background"];
            this.BackgroundImage = Image.FromFile(path_background_image);

            DebugConsolePrint($"form size: (w,h) = ({this.Width},{this.Height})");
            int[] panelInputLocation = { (int)(this.Width * 0.175), (int)(this.Height * 0.68) };
            DebugConsolePrint($"new panel input location: (x,y) = ({panelInputLocation[0]},{panelInputLocation[1]})");
            panel_Input.Location = new Point(panelInputLocation[0], panelInputLocation[1]);

            int[] panelProfileInfoLocation = { (int)(this.Width * 0.06), (int)(this.Height * 0.37) };
            DebugConsolePrint($"new panel input location: (x,y) = ({panelProfileInfoLocation[0]},{panelProfileInfoLocation[1]})");
            panel_ProfileInfo.Location = new Point(panelProfileInfoLocation[0], panelProfileInfoLocation[1]);

            int[] panelProfilePictureLocation = { (int)(this.Width * 0.57), (int)(this.Height * 0.3) };
            DebugConsolePrint($"new panel input location: (x,y) = ({panelProfilePictureLocation[0]},{panelProfilePictureLocation[1]})");
            panel_ProfilePicture.Location = new Point(panelProfilePictureLocation[0], panelProfilePictureLocation[1]);
        }

        // Update all range
        private void UpdateAllAttendance(string flag)
        {
            List<object> resetValues = new List<object>() { flag.ToUpper() };

            ValueRange resetValuesRange = new ValueRange();

            List<IList<object>> resetValuesList = new List<IList<object>>();

            for (int i = 0; i < 204; i++)
            {
                resetValuesList.Add(resetValues);
            }

            resetValuesRange.Values = resetValuesList;

            if (spreadSheet.WriteData(sheetRange_attendance, resetValuesRange))
                MessageBox.Show("Đã xóa dữ liệu thành công", "Thông báo");
            else
                MessageBox.Show("Xóa dữ liệu thất bại, vui lòng thử lại", "Thông báo");

        }
        #endregion



        #region automatic code generation by VS
        public MainForm()
        {
            InitializeComponent();
            InitLayoutLocation();
            InitData();
        }

        private void textbox_Input_KeyDown(object sender, KeyEventArgs e)
        {
            string inputCode = textbox_Input.Text;
            if (e.KeyCode == Keys.Enter && inputCode.Length > 0)
            {
                if (CheckInternetConnection())
                {
                    bool isFound = false;
                    foreach (var rowData in sheetData_data)                     // ID index in spreadsheet
                    {
                        if (rowData[1].ToString().Equals(inputCode))            // Check ID
                        {
                            label_Welcome.Text = "Chào mừng Đại biểu";
                            label_Name.Text = rowData[2].ToString();            // Họ và tên
                            label_Organization.Text = rowData[3].ToString();    // Đoàn

                            LoadProfileImage(inputCode);

                            UpdateAttendanceStatus(Int32.Parse(rowData[0].ToString()) /*Index of update range*/);

                            isFound = true;
                            break;
                        }
                    }

                    if (!isFound)
                    {
                        label_Welcome.Text = "";
                        label_Name.Text = "Dữ liệu không tồn tại";      // Họ và tên
                        label_Organization.Text = "";                   // Đoàn
                        LoadErrorImage();
                    }

                    textbox_Input.Clear();
                }
            }
        }

        private void rjToggleButton_CheckIn_CheckedChanged(object sender, EventArgs e)
        {
            if (rjToggleButton_CheckIn.Checked) // Check-out
            {
                label_CheckIn.Text = "Check-out";
                SetAppSettingValue("VAR_rjToggle_checking_status", "1");
            }
            else // Check-in
            {
                label_CheckIn.Text = "Check-in";
                SetAppSettingValue("VAR_rjToggle_checking_status", "0");
            }
            textbox_Input.Focus();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(Environment.ExitCode);
        }

        private void MainForm_SizeChanged(object sender, EventArgs e)
        {
            InitLayoutLocation();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Task.Delay(500);

            if (CheckInternetConnection())
            {
                if (GoogleSheet_Auth())
                {
                    string[] batchRange = { sheetRange_data, sheetRange_total };
                    IList<ValueRange> valueRanges = spreadSheet.ReadBatchData(batchRange);
                    int attempt = 5;
                    while (valueRanges == null && attempt > 0)
                    {
                        valueRanges = spreadSheet.ReadBatchData(batchRange);
                        --attempt;
                    }

                    sheetData_data = valueRanges[0].Values;
                    sheetData_total = valueRanges[1].Values;

                    label_TotalAttendance.Text = $"Số Đại biểu: {sheetData_total[0][1]}/{sheetData_total[0][0]}" /*total*/+
                        $"\n({sheetData_total[0][2]})" /*percent*/;

                    DebugConsolePrint($"Number: {sheetData_total[0][1]} - Total: {sheetData_total[0][0]} - Percent: {sheetData_total[0][2]}");

                    Task.Delay(500);

                    Thread thread = new Thread(new ThreadStart(ThreadUpdateTotalAttendance));
                    thread.Start();
                }
            }

            textbox_Input.Focus();
        }
        #endregion

        private void rjButton_Reset_Click(object sender, EventArgs e)
        {
            switch (MessageBox.Show("Bạn có chắc chắn muốn xóa liệu điểm danh?",
                            "Cảnh báo",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question))
            {
                case DialogResult.Yes:
                    // "Yes" processing
                    if (CheckInternetConnection())
                    {
                        UpdateAllAttendance("false");
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng kiểm tra lại đường truyền mạng. Chương trình sẽ tiếp tục khi có internet", "Cảnh báo");
                    }
                    break;

                case DialogResult.No:
                    // "No" processing
                    break;
            }
        }

        private void rjButton_Setting_Click(object sender, EventArgs e)
        {

        }
    }
}
