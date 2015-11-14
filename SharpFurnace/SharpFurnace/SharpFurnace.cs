// These were included by default after creating the windows form display.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

// Included MccDaq to use the Universal Library functions.
using MccDaq;

// Included this since this program uses threads.
using System.Threading;

// Included this since this program writes to a file.
using System.IO;

// Included this to send emails
using System.Net.Mail;
using System.Net;


namespace USB_TC_Control
{
    public partial class USB_TC_Control_Form : Form
    {
        #region Global Variables


        #region USB ThermoCouple Reader DAQ Settings

        // The number of thermocouple input channels, set to 8 by default.
        public static int CHANCOUNT = 8;

        // The max. number of thermocouple input channels supported.
        public const int MAX_CHANCOUNT = 8;

        public static int NUM_WALLCOUNT = 2;

        // Give the device a name. Valid names are USB-TC and TC.
        public const string DEVICE = "TC";
                
        // This stores the number given to the USB-TC board. The number
        // given is usually 0 by default.
        public static int BoardNum;

        public static MccBoard thermocouple_board;
        public static MccBoard oven_board;

        #endregion

        #region Temperature Data & Data Settings

        // Set default temperature scale to degree Celsius.
        public static TempScale my_TempScale = TempScale.Celsius;
        
        // This is a copy of the TempData[] array in the degree celsius
        // scale. This is the array used for logging the temperature only.
        public float[] CurrentTemperatures = new float[MAX_CHANCOUNT] { 0, 0, 0, 0, 0, 0, 0, 0 };

        public List<float[]> PlateauTemperatureWindow = new List<float[]>();
        public int PlateauWindowLength = 120;
        public double PlateauTemperatureDeviationThreshold = 1;
        public double PercentageOffSetTemperature = 5;

        public static double set_point_temperature = 0;

        public PIDSettings oven_pid_settings = new PIDSettings();

        public Boolean IsLockedCurrentTemp_Array = false;

        #endregion

        #region Oven Process Thread Declarations

        // This thread handles the reading and displaying of thermocouples.
        private Thread temperature_read_thread;

        // This thread handles the oven heating coils and gas cooling valves.
        private Thread oven_control_thread;

        // This thread handles the temperature and error logging during the
        // heating process.
        private Thread log_thread;

        #endregion

        #region Data and Log File Save Settings

        // String variable used for complete file directory and name.
        public string directory = "";

        // String variable used for file folder.
        public string filepath = "";

        // String variable used for file name.
        public string filename = "";

        // DateTime object used to log the date and time logging started.
        public DateTime Start_Log_File_Time = new DateTime();

        // String variable used to to format date_time to a string.
        public string format = " M_d_yy @ h.mm tt";

        #endregion

        #region Global Oven State Variables

        //Heating Element On/Off Colors
        public static Color ElementOnColor = Color.Red;
        public static Color ElementOffColor = Color.Black;
        public static Color ElementDisabledColor = Color.LightGray;

        // This is a global variable used to check if the thermocouples' 
        // readings are being displayed.
        // 0 = temperature is not being read and displayed
        // 1 = temperature is being read and displayed
        // -1 = exiting program / error reading
        public static ComponentCheckStatusEnum read_check = 0;

        // This is a global variable used to set oven state by button clicks.
        public static ComponentCheckStatusEnum oven_check = 0;

        public static AirChanStatusEnum air_status = AirChanStatusEnum.AirOff;
        public static OvenChanEnum oven_status = OvenChanEnum.AllOff;

        // This is a global variable used in graphing to check the oven state.
        public static int graph_check = 2;

        // This is a global variable used to set oven state by button clicks.
        public static ComponentCheckStatusEnum log_check = ComponentCheckStatusEnum.None;

        // This is a global variable used in updating the temperature log.
        public static LogStatusEnum log_state = LogStatusEnum.NotLogged;

        // This is a global variable to check if temp_read_thread is started. The
        // log_thread starts with the temp_read_thread.
        public static ComponentCheckStatusEnum temperature_start_check = ComponentCheckStatusEnum.None;

        // This is a global variable to check if oven_control_thread is started.
        public static ComponentCheckStatusEnum oven_start_check = ComponentCheckStatusEnum.None;

        //Used to record whether or not the oven_cooling period has been activated in an oven run
        public static ComponentCheckStatusEnum oven_cooling_check = ComponentCheckStatusEnum.None;

        // String variable that stores the choice of cooling (air / nitrogen).
        public string CoolingMethod = "";

        // This variable is used as a flag for aborting the oven code
        // If true, the code will call a method to stop all of the threads and 
        // end the program; if false, the code will continue it's execution
        public static bool abort_oven_program = false;

        // Global boolean flag to let the code know that the method used to
        // close the program and stop all of the threads has been called already.
        public static bool close_program_inprogress = false;

        // This variable is used to make sure that all checkboxes are not
        // unchecked during an oven run (flying blind)
        public bool not_showing_temp = false;

        // Variables that are used for logging and keeping track of oven state.
        public bool HeatingStarted = false;
        public bool HeatingStopped = false;
        public bool MaxTempPlateauReached = false;
        public bool CoolingStarted = false;
        public bool IsCriticalApplicationError = false;
        public bool OvenRunDone = false;

        public bool outer_status = false;
        public bool sample_zone_status = false;
        public bool inner_status = false;

        public bool outer_zone_heater_disabled = false;
        public bool sample_zone_heater_disabled = false;
        public bool inner_zone_heater_disabled = false;

        public const int inner_zone_heater_index = 2;
        public const int sample_zone_heater_index = 1;
        public const int outer_zone_heater_index = 0;


        public bool HeatingStarted_logged = false;
        public bool HeatingStopped_logged = false;
        public bool MaxTempReached_logged = false;
        public bool CoolingStarted_logged = false;
        public bool IsCriticalApplicationError_logged = false;
        public bool OvenRunDone_logged = false;

        public DateTime MaxTemp_StartTime;
        public Boolean MaxTemp_StartTime_Set = false;

        #region Error Notification / Error Warning Variables

        // This variable is used in updating the Warning box
        public static int warning_time = 20;

        // This variable is used as a flag to let error processor know if an 
        // error window is already open (not the PID error!)
        public static Boolean application_error_window_open = false;
        
        public string CriticalErrorMessage = "";

        #endregion

        
        // Variables that are used by the is_parameters_changed() function
        // These variables are used to detect a change in state of the oven
        // system and its various control and user settings parameters
        public string User_copy = "";
        public string UserEmail_copy = "";
        public string UserPhone_copy = "";
        public string BatchCode_copy = "";
        public float MaxTemp_copy = 20;
        public TimeSpan HoldTime_copy = new TimeSpan(0);
        public string CoolMethod_copy = "";
        public string directory_copy = "";
        public string OvenName_copy = "";
        public float SwitchTemp_copy = 20;
        public float StopTemp_copy = 30;
        public string OvenStartTime_copy = "";
        public string TimeElapsed_copy = "";
        public string HoldStartTime_copy = "";
        public string HoldTimeRemaining_copy = "";
        public string CoolStartTime_copy = "";
        public string CoolTimeElapsed_copy = "";

        // Variable to keep track of oven run time
        DateTime Start_Oven_Time = new DateTime();

        // Variable to keep track of heating time
        DateTime Heat_Time = new DateTime();

        // Variable to keep track of the temperature read time
        DateTime Start_Temp_Time = new DateTime();

        #endregion

        #region Email / SMS Variables and Settings

        // These variables are for the email function
        public SmtpClient smtp_client = new SmtpClient();
        public MailMessage email_message = new MailMessage();
        public MailAddress from_address = new MailAddress("rapid.beta.tester@gmail.com","Oven Notifications");
        public MailAddress to_address = new MailAddress("test@test.com");
        public String email_from_password = "igfoup666";

        // Regular expression to check email addresses
        public const string permissive_email_reg_ex =
                        @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+" +
                        @"(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)" +
                        @"*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+" +
                        @"[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z";

        #endregion

        #endregion

        #region Enumerators

        // An enumerator 
        public enum OvenChanEnum
        {
            AllOff = 0x00,
            OvenOn = 0x01,
            OvenOff = 0x02,
            Inner = 0x04,
            SampleZone = 0x08,
            Outer = 0x10,
        }

        public enum OvenHeatOnOffEnum
        {
            None,
            On,
            Off
        }

        public enum AirChanStatusEnum
        {
            AirOff = 0x00,
            LowN2 = 0x01,
            HighN2 = 0x02,
            Air = 0x04
        }

        public enum ParameterTypeEnum
        { 
            None = 0,
            Oven = 1,
            ThermoCouples = 2
        }

        public enum TemperatureVsTimeDataSeriesEnum
        { 
            None = -1,
            Inner = 7,
            Outer = 6,
            TrayInner_Edge = 2,
            TraySampleZone_Edge = 1,
            TrayOuter_Edge = 0,
            TrayInner_Center = 3,
            TraySampleZone_Center = 4,
            TrayOuter_Center = 5
        }

        public enum ErrorCodeEnum
        { 
            NoError = 0,
            Error = 1
        }

        public enum ComponentCheckStatusEnum
        {
            Aborted = -1,
            None = 0,
            Set = 1,
            Running = 2
        }

        public enum LogStatusEnum
        { 
            Aborted = -1,
            NotLogged = 0,
            Logged = 1
        }

        #endregion

        #region Delegates


        // A delegate used to the update the temperature readouts.
        public delegate void Update_Temperature(float[] TemperatureData, float[] TemperatureData_C);

        // A delegate used to the update the WARNING box.
        public delegate void Update_Warning(string Warning, ErrorCodeEnum ErrorCode);

        // A delegate used to the reset the checkboxes.
        public delegate void Reset_Checkboxes();

        //A delegate used to trigger the close oven program process;
        public delegate void AbortOvenCode();

        public delegate void EnableStartOvenButton();

        public delegate void SetHeatingElementStatusButtonCallback(OvenChanEnum oven_element);

        #endregion
        
        #region Methods

        #region Form Constructors

        // Necessary constructor for the program to load properly.
        public USB_TC_Control_Form()
        {
            Program.splash_screen_form.BeginInvoke(
                (Action)(() => Program.splash_screen_form.WaitForHideSplashScreen()));

            InitializeComponent();

            StartCoolingButton.Enabled = false;

            System.Windows.Forms.DialogResult user_resp = System.Windows.Forms.DialogResult.None;
            
            do
            {

                // Locate the USB-TC and give it a number.
                // BNum is a TextBox
                BoardNum = GetBoardNum(DEVICE, UsbThermoCoupleDaqBoardNumberTextBox);

                if (BoardNum == -1)
                {
                    user_resp =
                        MessageBox.Show(String.Format("No USB-{0} Thermocouple Input DAQ Device detected! "
                                        + "Click 'Cancel' to exit out of the oven program.",
                                        DEVICE),
                                        "Comm Error",
                                        MessageBoxButtons.RetryCancel);

                    // Throw an error message if no USB-TC device is found. It is
                    // recommended to first locate and caliberate the board in
                    // InstaCal before use.

                    // Give user the option to exit from the code in the modal error message window.
                    if (user_resp == System.Windows.Forms.DialogResult.Cancel)
                    {
                        StopThreads_AndCloseOvenProgram();
                        return;
                    }
                }
                else
                {
                    // Set default value for the button press detection variables.
                    read_check = ComponentCheckStatusEnum.None;
                    oven_check = ComponentCheckStatusEnum.None;
                    log_check = ComponentCheckStatusEnum.None;

                    //Instantiate boards for ThermoCouples and Oven
                    SetupThermoCoupleDAQBoard();
                    SetupOvenDAQBoard();
                }

            } while (user_resp == System.Windows.Forms.DialogResult.Retry);


            // Detect number of connected thermocouples
            DetectThermoCoupleInputs();
            
            Outer_HeatingElementStatusButton.BackColor = ElementOffColor;
            Inner_HeatingElementStatusButton.BackColor = ElementOffColor;
            SampleZone_HeatingElementStatusButton.BackColor = ElementOffColor;

            // Default batch code
            BatchIDCodeTextBox.Text = "_";

            // Default user
            UserNameTextBox.Text = "User";

            // Default email
            UserEmailTextBox.Text = "user@emailaddress.com";

            // Default phone
            UserPhoneTextBox.Text = "111-222-3333";

            // Default max. temp
            MaxTemperatureNumericUpDown.Value = 50;
            

            // Default hold time
            HoldTimeAtPeakNumericUpDown.Value = 30;

            // Default Stop Cooling Temperature
            StopCoolingTemperatureNumericUpDown.Value = 25;
            
            // Default cooling method is Nitrogen
            CoolMethodComboBox.SelectedIndex = 1;

            // Setup graph
            TempertureVsTimeGraph.ChartAreas[0].BackColor = Color.Black;
            TempertureVsTimeGraph.ChartAreas[0].AxisX.MajorGrid.LineColor =
                Color.DarkSlateGray;
            TempertureVsTimeGraph.ChartAreas[0].AxisY.MajorGrid.LineColor =
                Color.DarkSlateGray;

            TempertureVsTimeGraph.ChartAreas[0].AxisX.Minimum = 0;
            TempertureVsTimeGraph.ChartAreas[0].AxisY.Minimum = 0;

            TempertureVsTimeGraph.Series[
                Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                             TemperatureVsTimeDataSeriesEnum.TrayOuter_Edge)].Points.AddXY(0, 0);

            TempertureVsTimeGraph.Series[
                Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                             TemperatureVsTimeDataSeriesEnum.TraySampleZone_Edge)].Points.AddXY(0, 0);

            TempertureVsTimeGraph.Series[
                Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                             TemperatureVsTimeDataSeriesEnum.TrayInner_Edge)].Points.AddXY(0, 0);

            TempertureVsTimeGraph.Series[
                Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                             TemperatureVsTimeDataSeriesEnum.TrayOuter_Center)].Points.AddXY(0, 0);

            TempertureVsTimeGraph.Series[
                Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                             TemperatureVsTimeDataSeriesEnum.TraySampleZone_Center)].Points.AddXY(0, 0);

            TempertureVsTimeGraph.Series[
                Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                             TemperatureVsTimeDataSeriesEnum.TrayInner_Center)].Points.AddXY(0, 0);

            TempertureVsTimeGraph.Series[
                Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                             TemperatureVsTimeDataSeriesEnum.Inner)].Points.AddXY(0, 0);

            TempertureVsTimeGraph.Series[
                Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                             TemperatureVsTimeDataSeriesEnum.Outer)].Points.AddXY(0, 0);
        }

        #endregion

        #region Form Control Event Methods

        private void USB_TC_Control_Form_Load(object sender, EventArgs e)
        {
            Program.hide_splash_screen = true;

            SetAirRelayChannelOutput(AirChanStatusEnum.AirOff);
            SetOvenChannelOutput(OvenChanEnum.AllOff);
        }

        // This function is called when the Read TC button in the GUI is
        // clicked. It sets the value of the global variable read_check to 1,
        // changes the Read TC button color to Green, and resets the color of
        // the Stop TC button to default, so that the user knows which button
        // was last clicked.
        private void ReadThermoCouplesbutton_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(WarningTextBox.Text))
            {
                if (not_showing_temp == true)
                    ResetThemocoupleInputStatusCheckboxes();
                UpdateWarningDisplay("", ErrorCodeEnum.NoError);
            }

            if (temperature_start_check == ComponentCheckStatusEnum.None && 
                oven_start_check == ComponentCheckStatusEnum.None)
                ResetInputControls();

            if (CheckParameters(ParameterTypeEnum.ThermoCouples) != ErrorCodeEnum.NoError)
            {
                MessageBox.Show(
                    "Check and complete highlighted thermocouple" +
                    " read parameters!", 
                    "   ERROR!");

                return;
            }

            UpdateDirectory();
            if (String.IsNullOrWhiteSpace(directory))
            {
                SaveLogFileDirectoryTextBox.BackColor = Color.MistyRose;

                MessageBox.Show(
                    String.Format(
                        "Please choose a directory to save log file in!{0}" +
                        "Also add other (optional) parameters, like the batch " +
                        "code and email address...",
                        Environment.NewLine),
                    "   " + "ERROR!");

                return;
            }
            else
            {
                SaveLogFileDirectoryTextBox.BackColor = Color.White;
            }

            if (IsParametersChanged())
            {
                UpdateDirectory();
                MessageBox.Show(
                    "Please double check all values, just to be " +
                    "sure. Click the Read TC button when ready.", 
                    "   Check Values");

                CopyParameters();
                return;
            }

            UserNameTextBox.ReadOnly = true;
            UserNameTextBox.BorderStyle = BorderStyle.FixedSingle;
            BatchIDCodeTextBox.ReadOnly = true;
            BatchIDCodeTextBox.BorderStyle = BorderStyle.FixedSingle;
            OvenNameTextBox.ReadOnly = true;
            OvenNameTextBox.BorderStyle = BorderStyle.FixedSingle;
            UserEmailTextBox.Enabled = false;
            UserPhoneTextBox.Enabled = false;
            SwitchToAirTemperatureNumericUpDown.Enabled = false;
            MaxTemperatureNumericUpDown.Enabled = false;            
            HoldTimeAtPeakNumericUpDown.Enabled = false;
            CoolMethodComboBox.Enabled = false;

            UserPhoneTextBox.BackColor = Color.White;

            // Create and start the threads when the Read TC button is clicked
            // for the first time. Log thread and the temperature read thread
            // are the two threads started here
            if (temperature_start_check == ComponentCheckStatusEnum.None)
            {
                CopyParameters();

                Start_Temp_Time = DateTime.Now;
                Start_Log_File_Time = Start_Temp_Time;
                temperature_read_thread = new Thread(new ThreadStart(ReadTemperature));
                temperature_read_thread.Start();
                temperature_start_check = ComponentCheckStatusEnum.Running;

                if (log_check != ComponentCheckStatusEnum.Running)
                {
                    log_thread = new Thread(new ThreadStart(UpdateLog));
                    log_check = ComponentCheckStatusEnum.Running;
                    log_state = LogStatusEnum.NotLogged;
                    log_thread.Start();
                }
            }

            if (IsParametersChanged())
                UpdateDirectory();

            if (read_check != ComponentCheckStatusEnum.Running)
                Start_Temp_Time = DateTime.Now;

            read_check = ComponentCheckStatusEnum.Running;

            warning_time = 20;

            ReadTC_Button.BackColor = Color.GreenYellow;            
            
            StopTC_Button.BackColor = default(Color);
        }

        // This function is called when the Stop TC button in the GUI is
        // clicked. It sets the value of the global variable read_check to 0,
        // changes the Stop TC button color to Green, and resets the color of
        // the Read TC button to default, so that the user knows which button 
        // was last clicked.
        private void StopTCbutton_Click(object sender, EventArgs e)
        {
            if (oven_check == ComponentCheckStatusEnum.None)
            {
                UserEmailTextBox.Enabled = true;
                UserPhoneTextBox.Enabled = true;
                MaxTemperatureNumericUpDown.Enabled = true;
                HoldTimeAtPeakNumericUpDown.Enabled = true;
                CoolMethodComboBox.Enabled = true;
                SwitchToAirTemperatureNumericUpDown.Enabled = true;
            }

            read_check = ComponentCheckStatusEnum.None;

            ReadTC_Button.BackColor = default(Color);

            for (int i = 0; i < MAX_CHANCOUNT; i++)
            {
                this.Controls["Channel" + i.ToString("#0") + "TextBox"].ForeColor = Color.Gray;
            }

            if (temperature_start_check != ComponentCheckStatusEnum.None)
                StopTC_Button.BackColor = Color.GreenYellow;
        }

        // This function is called when the Start Oven button in the GUI is
        // clicked. It sets the value of the global variable check to 3, and
        // changes the Start Oven button color to Green.
        private void StartOvenButton_Click(object sender, EventArgs e)
        {
            UserEmailTextBox.Enabled = true;
            UserPhoneTextBox.Enabled = true;
            MaxTemperatureNumericUpDown.Enabled = true;
            HoldTimeAtPeakNumericUpDown.Enabled = true;
            CoolMethodComboBox.Enabled = true;
            SwitchToAirTemperatureNumericUpDown.Enabled = true;

            if (CheckParameters(ParameterTypeEnum.Oven) != 
                    ErrorCodeEnum.NoError)
            {
                if (String.IsNullOrWhiteSpace(directory))
                    SaveLogFileDirectoryTextBox.BackColor = Color.MistyRose;
                else
                    SaveLogFileDirectoryTextBox.BackColor = Color.White;

                MessageBox.Show(
                    "Check and complete highlighted oven run" +
                    " parameters!", 
                    "   ERROR!");

                return;
            }

            if (String.IsNullOrWhiteSpace(directory))
            {
                SaveLogFileDirectoryTextBox.BackColor = Color.MistyRose;
                MessageBox.Show(
                    "Please choose a directory to save log " +
                    "file in!", 
                    "   " + "ERROR!");

                return;
            }
            else
                SaveLogFileDirectoryTextBox.BackColor = Color.White;

            if (IsParametersChanged() ||
                (temperature_start_check == ComponentCheckStatusEnum.Set && 
                 oven_start_check == ComponentCheckStatusEnum.None))
            {
                UpdateDirectory();
                MessageBox.Show("Please double check all values, just to be " +
            "sure. Click the Start Oven button when ready.", "   Check Values");

                CopyParameters();
                oven_start_check = ComponentCheckStatusEnum.Set;
                return;
            }

            UserNameTextBox.ReadOnly = true;
            UserNameTextBox.BorderStyle = BorderStyle.FixedSingle;
            UserEmailTextBox.ReadOnly = true;
            UserEmailTextBox.BorderStyle = BorderStyle.FixedSingle;
            UserPhoneTextBox.ReadOnly = true;
            UserPhoneTextBox.BorderStyle = BorderStyle.FixedSingle;
            UserPhoneTextBox.BackColor = Color.White;
            BatchIDCodeTextBox.ReadOnly = true;
            BatchIDCodeTextBox.BorderStyle = BorderStyle.FixedSingle;
            //MaxTemperatureNumericUpDown.ReadOnly = true;
            //MaxTemperatureNumericUpDown.BorderStyle = BorderStyle.FixedSingle;
            //HoldTimeAtPeakNumericUpDown.ReadOnly = true;
            //HoldTimeAtPeakNumericUpDown.BorderStyle = BorderStyle.FixedSingle;
            OvenNameTextBox.ReadOnly = true;
            OvenNameTextBox.BorderStyle = BorderStyle.FixedSingle;
            CoolMethodComboBox.Enabled = false;
            //SwitchToAirTemperatureNumericUpDown.ReadOnly = true;
            //SwitchToAirTemperatureNumericUpDown.BorderStyle = BorderStyle.FixedSingle;

            if (IsParametersChanged())
                UpdateDirectory();

            if (oven_check != ComponentCheckStatusEnum.Running)
            {
                Start_Oven_Time = DateTime.Now;
                OvenStartTimeTextBox.Text = Start_Oven_Time.ToString("T");
                graph_check++;
            }

            CopyParameters();

            int count = 0;

            for (int i = 0; i < MAX_CHANCOUNT; i++)
            {
                if (((CheckBox)this.Controls["Channel" +
                    i.ToString("#0") + "CheckBox"]).Enabled == true &&
                    ((CheckBox)this.Controls["Channel" +
                    i.ToString("#0") + "CheckBox"]).Checked == true)
                {
                    count++;
                }
            }

            if (count == 0)
            {
                not_showing_temp = true;
            }

            if (not_showing_temp == true)
            {
                ResetThemocoupleInputStatusCheckboxes();
                not_showing_temp = false;
            }

            // Start thread once the Start Oven button is clicked for the
            // first time.
            if (oven_start_check == ComponentCheckStatusEnum.Set)
            {
                oven_control_thread = new Thread(new ThreadStart(PID_control));
                oven_control_thread.Start();
                oven_start_check = ComponentCheckStatusEnum.Running;

                CopyParameters();

                if (temperature_start_check != ComponentCheckStatusEnum.Running)
                {
                    temperature_read_thread = new Thread(new ThreadStart(ReadTemperature));
                    temperature_read_thread.Start();
                    temperature_start_check = ComponentCheckStatusEnum.Running;
                    
                    if (log_check != ComponentCheckStatusEnum.Running)
                    {
                        log_thread = new Thread(new ThreadStart(UpdateLog));
                        log_check = ComponentCheckStatusEnum.Running;
                        log_state = LogStatusEnum.NotLogged;
                        log_thread.Start();
                    }

                    read_check = ComponentCheckStatusEnum.Running;
                    ReadTC_Button.BackColor = Color.GreenYellow;
                    StopTC_Button.BackColor = default(Color);
                }
            }

            oven_check = ComponentCheckStatusEnum.Running;

            StopOvenButton.BackColor = default(Color);
            StartOvenButton.BackColor = Color.GreenYellow;
        }

        // This function is called when the Start Oven button in the GUI is
        // clicked. It sets the value of the global variable check to 3, and
        // changes the Start Oven button color to Green.
        private void StopOvenButton_Click(object sender, EventArgs e)
        {
            UserNameTextBox.ReadOnly = false;
            UserNameTextBox.BorderStyle = BorderStyle.Fixed3D;
            OvenNameTextBox.ReadOnly = false;
            OvenNameTextBox.BorderStyle = BorderStyle.Fixed3D;
            UserEmailTextBox.ReadOnly = false;
            UserEmailTextBox.BorderStyle = BorderStyle.Fixed3D;
            UserPhoneTextBox.ReadOnly = false;
            UserPhoneTextBox.BorderStyle = BorderStyle.Fixed3D;
            UserPhoneTextBox.BackColor = Color.White;
            BatchIDCodeTextBox.ReadOnly = false;
            BatchIDCodeTextBox.BorderStyle = BorderStyle.Fixed3D;
            MaxTemperatureNumericUpDown.ReadOnly = false;
            MaxTemperatureNumericUpDown.BorderStyle = BorderStyle.Fixed3D;
            HoldTimeAtPeakNumericUpDown.ReadOnly = false;
            HoldTimeAtPeakNumericUpDown.BorderStyle = BorderStyle.Fixed3D;
            CoolMethodComboBox.Enabled = true;

            if (oven_check != ComponentCheckStatusEnum.None)
            {
                DialogResult dialogResult = 
                    MessageBox.Show(
                        "WARNING! This will stop the oven run process and shut down all coils!" +
                        " Are you sure you want to stop the oven run?", 
                        "   " +
                        "STOP OVEN WARNING",
                        MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {
                    oven_check = ComponentCheckStatusEnum.None;
                    StartOvenButton.BackColor = default(Color);

                    if (oven_start_check != ComponentCheckStatusEnum.Running)
                        StopOvenButton.BackColor = Color.GreenYellow;

                    HeatingStopped = true;
                    HeatingStarted = false;
                    HeatingStarted_logged = false;
                }
                else
                    return;
            }
            return;
        }

        // This function is called when the degree Celsius button in the GUI
        // is clicked. It changes the button color to Green and resets the color
        // of the other temperature scale buttons, so that the user knows which 
        // temperature scale is currently being displayed.
        private void Celsius_Click(object sender, EventArgs e)
        {
            my_TempScale = TempScale.Celsius;
            CelsiusButton.BackColor = Color.GreenYellow;
            FahrenheitButton.BackColor = default(Color);
            KelvinButton.BackColor = default(Color);
        }

        // This function is called when the degree Fahrenheit button in the GUI
        // is clicked. It changes the button color to Green and resets the color
        // of the other temperature scale buttons, so that the user knows which 
        // temperature scale is currently being displayed.
        private void Fahrenheit_Click(object sender, EventArgs e)
        {
            my_TempScale = TempScale.Fahrenheit;
            CelsiusButton.BackColor = default(Color);
            FahrenheitButton.BackColor = Color.GreenYellow;
            KelvinButton.BackColor = default(Color);
        }

        // This function is called when the Kelvin scale button in the GUI is
        // clicked. It changes the button color to Green and resets the color
        // of the other temperature scale buttons, so that the user knows which 
        // temperature scale is currently being displayed.
        private void Kelvin_Click(object sender, EventArgs e)
        {
            my_TempScale = TempScale.Kelvin;
            CelsiusButton.BackColor = default(Color);
            FahrenheitButton.BackColor = default(Color);
            KelvinButton.BackColor = Color.GreenYellow;
        }

        // This function is used to open a directory to which the user wants to
        // save the log file.
        private void OpenDirectory_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            fbd.Description = "Select a location in which the temperature log" +
                " files will be saved.";

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SaveLogFileDirectoryTextBox.BackColor = Color.White;

                filepath = fbd.SelectedPath;

                if (BatchIDCodeTextBox.Text == "_" ||
                    String.IsNullOrWhiteSpace(BatchIDCodeTextBox.Text))
                    BatchIDCodeTextBox.Text = UserNameTextBox.Text;

                Start_Log_File_Time = DateTime.Now;

                filename = String.Join("", UserNameTextBox.Text, "_", BatchIDCodeTextBox.Text, "",
                    Start_Log_File_Time.ToString(format), ".txt");

                directory = Path.Combine(filepath, filename);

                SaveLogFileDirectoryTextBox.Text = directory;
            }
        }

        // This function fill in values for the oven run parameters. It is
        // useful for testing purposes.
        private void TestMode_Click(object sender, EventArgs e)
        {
            ResetInputControls();
            UserNameTextBox.Text = "isaac";
            UserEmailTextBox.Text = "ihilburn@caltech.edu";
            UserPhoneTextBox.Text = "805-258-3141";
            OvenNameTextBox.Text = "Lowenstam";
            BatchIDCodeTextBox.Text = "EmptyAir";
            MaxTemperatureNumericUpDown.Value = 50;
            HoldTimeAtPeakNumericUpDown.Value = 30;
            UserPhoneTextBox.BackColor = Color.White;
            CoolMethodComboBox.SelectedIndex = 0;
            SwitchToAirTemperatureNumericUpDown.Value = 150;
            StopCoolingTemperatureNumericUpDown.Value = 27;
            SaveLogFileDirectoryTextBox.Text = @"C:\Users\lab\Dropbox\Ge SURF\Stuff\Oven Data";
            filepath = SaveLogFileDirectoryTextBox.Text;

        }

        // This function is called when the cooling method is changed.
        private void CoolMethod_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CoolMethodComboBox.SelectedIndex == 1)
            {
                SwitchToAirTemperatureNumericUpDown.Enabled = true;
                SwitchToAirTemperatureNumericUpDown.ReadOnly = false;
                SwitchToAirTemperatureNumericUpDown.BorderStyle = BorderStyle.Fixed3D;
            }
            else
            {
                SwitchToAirTemperatureNumericUpDown.ReadOnly = true;
                SwitchToAirTemperatureNumericUpDown.BackColor = Color.White;
                SwitchToAirTemperatureNumericUpDown.BorderStyle = BorderStyle.FixedSingle;
            }
        }

        //Event handler for the 'Exit' button
        private void ButtonExit_Click(object sender, EventArgs e)
        {
            StopThreads_AndCloseOvenProgram();
        }

        //Turn on Air Cooling (overrides current air relay valve settings)
        private void ForceAir_Click(object sender, EventArgs e)
        {
            if (air_status != AirChanStatusEnum.Air)
            {
                SetAirRelayChannelOutput(AirChanStatusEnum.Air);
                ForceAirButton.Text = "Stop Air";
            }
            else
            {
                SetAirRelayChannelOutput(AirChanStatusEnum.AirOff);
                ForceAirButton.Text = "Force Air";
            }
        }

        //Turn on Nitrogen Cooling (overrides current nitrogen relay valve settings)
        private void ForceN2_Click(object sender, EventArgs e)
        {
            if (air_status != AirChanStatusEnum.HighN2)
            {
                SetAirRelayChannelOutput(AirChanStatusEnum.HighN2);
                ForceN2Button.Text = "Stop N2";
            }
            else
            {
                SetAirRelayChannelOutput(AirChanStatusEnum.AirOff);
                ForceN2Button.Text = "Force N2";
            }

        }

        private void Outer_HeatingElementStatusButton_Click(object sender, EventArgs e)
        {
            if (outer_zone_heater_disabled)
            {
                outer_zone_heater_disabled = false;
                Outer_HeatingElementStatusButton.BackColor = ElementOffColor;
            }
            else
            {
                outer_zone_heater_disabled = true;
                Outer_HeatingElementStatusButton.BackColor = ElementDisabledColor;
            }
        }

        private void SampleZone_HeatingElementStatusButton_Click(object sender, EventArgs e)
        {
            if (sample_zone_heater_disabled)
            {
                sample_zone_heater_disabled = false;
                SampleZone_HeatingElementStatusButton.BackColor = ElementOffColor;
            }
            else
            {
                sample_zone_heater_disabled = true;
                SampleZone_HeatingElementStatusButton.BackColor = ElementDisabledColor;

            }
        }

        private void Inner_HeatingElementStatusButton_Click(object sender, EventArgs e)
        {
            if (inner_zone_heater_disabled)
            {
                inner_zone_heater_disabled = false;
                Inner_HeatingElementStatusButton.BackColor = ElementOffColor;
            }
            else
            {
                inner_zone_heater_disabled = true;
                Inner_HeatingElementStatusButton.BackColor = ElementDisabledColor;

            }
        }
        
        private void StartCoolingButton_Click(object sender, EventArgs e)
        {
            //End Oven Heating
            oven_check = ComponentCheckStatusEnum.None;
            CoolingStarted = true;
        }

        private void OpenPIDSettingsWindowButton_Click(object sender, EventArgs e)
        {
            Program.pid_settings_form.Show();
        }

        private void TestEmailButton_Click(object sender, EventArgs e)
        {
            SendEmail(
                "Test Email Settings",
                String.Format(
                    "Dear {0},{1}" +
                    "{2}This is a test of the C# Oven Code Email Settings.",
                    UserNameTextBox.Text,
                    Environment.NewLine,
                    "\t"));
        }

        #endregion
        
        #region DAQ USB ThermoCouple Input related Methods
        // This function changes CHANCOUNT, by detecting the actual number of
        // working thermocouple inputs connected to the USB-TC board.
        private void DetectThermoCoupleInputs()
        {
            MccBoard daq = new MccDaq.MccBoard(BoardNum);
            MccDaq.ErrorInfo RetVal;

            CHANCOUNT = MAX_CHANCOUNT;
            NUM_WALLCOUNT = 2;

            float[] temp = new float[MAX_CHANCOUNT];

            try
            {
                for (int i = 0; i < MAX_CHANCOUNT; i++)
                {
                    RetVal = daq.TIn(i, TempScale.Celsius,
                        out temp[i], ThermocoupleOptions.Filter);
                    
                    if (RetVal.Value != 0)
                    {
                        CHANCOUNT--;

                        ((CheckBox)this.Controls["Channel" +
                            i.ToString("#0") + "CheckBox"]).Checked = false;

                        ((CheckBox)this.Controls["Channel" +
                            i.ToString("#0") + "CheckBox"]).Enabled = false;

                        if (i == (int)TemperatureVsTimeDataSeriesEnum.Inner ||
                            i == (int)TemperatureVsTimeDataSeriesEnum.Outer)
                        {
                            NUM_WALLCOUNT--;
                        }
                    }
                    else
                    {
                        ((CheckBox)this.Controls["Channel" +
                            i.ToString("#0") + "CheckBox"]).Checked = true;

                        ((CheckBox)this.Controls["Channel" +
                            i.ToString("#0") + "CheckBox"]).Enabled = true;
                    }


                }

                NumThermocouplesDetectedTextBox.Text = " " + CHANCOUNT.ToString();

                return;
            }

            catch(Exception e)
            {
                MessageBox.Show(
                    String.Format(
                        "ThermoCouple_Finder has encountered an unexpected error!{0}" +
                        "Error Message: {1}{0}" +
                        "Error Source: {2}{0}" +
                        "Stack Trace: {3}{0}",
                        Environment.NewLine,
                        e.Message,
                        e.Source,
                        e.StackTrace)
                    , "   " + "ERROR!");
                return;
            }
        }
        
        // This function detects the board number. The two parameters of this
        // function are the device name and a TextBox object that holds the
        // value of the board number (which is set when the board is detected by
        // Instacal). It creates a board object with a board number from 0 to 99
        // (since that is the range of MCC board numbers) and checks if that  
        // board number corresponds to the USB-TC device. It is important that
        // the variable DEVICE is set to "TC" or "USB-TC", since that is what
        // InstaCal names the USB-TC device when it is detected. If no board is
        // found, a value of -1 is returned.
        public static int GetBoardNum(string dev, TextBox BNum)
        {
            for (int BoardNum = 0; BoardNum < 99; BoardNum++)
            {
                MccDaq.MccBoard daq = new MccDaq.MccBoard(BoardNum);

                if (daq.BoardName.Contains(dev))
                {
                    BNum.Text = " " + BoardNum.ToString();
                    
                    return BoardNum;
                }
            }
            return -1;
        }

        #endregion

        #region Error Checking Methods

        // The 3 functions below check for error by checking the value of an
        // ErrorInfo object. If there is an error, it displays the error 
        // message in a message box and returns 1. If there is no error, the
        // function just returns 0. (Overloaded IsError function)
        public static ErrorCodeEnum IsError(ErrorInfo error)
        {
            if (abort_oven_program) return ErrorCodeEnum.Error;

            while (application_error_window_open)
            {
                Thread.Sleep(20);
            }

            if (abort_oven_program) return ErrorCodeEnum.Error;
            
            if (error.Value != 0)
            {
                application_error_window_open = true;
                MessageBox.Show(error.Message);
                application_error_window_open = false;
                return ErrorCodeEnum.Error;
            }
            return ErrorCodeEnum.NoError;
        }
        
        // Helper function for IsError that gives user the option to respond to the error
        public static ErrorCodeEnum IsError_WithExitOption(object sender, ErrorInfo error)
        {
            if (abort_oven_program) return ErrorCodeEnum.Error;

            while (application_error_window_open)
            {
                Thread.Sleep(20);
            }

            if (abort_oven_program) return ErrorCodeEnum.Error;

            try
            {
                USB_TC_Control_Form oven_form = (USB_TC_Control_Form)sender;

                AbortOvenCode abortcode_del = new AbortOvenCode(oven_form.StopThreads_AndCloseOvenProgram);

                if (error.Value != 0)
                {
                    System.Windows.Forms.DialogResult user_resp = DialogResult.None;

                    application_error_window_open = true;
                    user_resp = MessageBox.Show(error.Message + Environment.NewLine +
                                                "Click 'Cancel' to exit out of the oven program",
                                                "Error",
                                                MessageBoxButtons.OKCancel);

                    if (user_resp == DialogResult.Cancel)
                    {
                        abort_oven_program = true;

                        // Set check flags for the three threads (Oven PID 
                        // Control, Read Temp, Log Temp) to -1 to end the
                        // threads
                        oven_check = ComponentCheckStatusEnum.Aborted;
                        read_check = ComponentCheckStatusEnum.Aborted;
                        log_check = ComponentCheckStatusEnum.Aborted;

                        // Pause 100 milliseconds
                        Thread.Sleep(100);

                        // Invoke the close/abort program delegate
                        oven_form.BeginInvoke(abortcode_del);
                    }

                    application_error_window_open = false;

                    return ErrorCodeEnum.Error;
                }
            }
            catch(Exception e)
            {

                
                    MessageBox.Show(
                        String.Format(
                            "Unexpected Error occurred while processing Oven Application Error." +
                            "{0}Error Message: {1}{0}" +
                            "Error Source: {2}{0}" +
                            "Stack Trace: {3}{0}",
                            Environment.NewLine,
                            e.Message,
                            e.Source,
                            e.StackTrace),
                        "    Unexpected Error!");                   

                //Call the single argument error handler
                IsError(error);
            }

            return ErrorCodeEnum.NoError;
        }
        
        // See above two functions
        public static ErrorCodeEnum IsError(Object sender, ErrorInfo e, Boolean allow_user_to_exit_program)
        {
            if (allow_user_to_exit_program)
            {
                return IsError_WithExitOption(sender, e);
            }
            else
            {
                return IsError(e);
            }
        }

        #endregion

        #region GUI Update Methods

        // This is a helper function used by update_temp() to reset the 
        // temperature display checkboxes in the event that they are all
        // unchecked during an oven run.
        private void ResetThemocoupleInputStatusCheckboxes()
        {
            for (int i = 0; i < MAX_CHANCOUNT; i++)
            {
                if (((CheckBox)this.Controls["Channel" +
                    i.ToString("#0") + "CheckBox"]).Enabled == true &&
                    ((CheckBox)this.Controls["Channel" +
                    i.ToString("#0") + "CheckBox"]).Checked == false)
                {
                    ((CheckBox)this.Controls["Channel" +
                    i.ToString("#0") + "CheckBox"]).Checked = true;
                }
            }
        }

        // This function resets the user input boxes' colors to white
        private void ResetInputControls()
        {
            UserNameTextBox.BackColor = Color.White;
            UserEmailTextBox.BackColor = Color.White;
            BatchIDCodeTextBox.BackColor = Color.White;
            MaxTemperatureNumericUpDown.BackColor = Color.White;
            HoldTimeAtPeakNumericUpDown.BackColor = Color.White;
            CoolMethodComboBox.BackColor = Color.White;
            SaveLogFileDirectoryTextBox.BackColor = Color.White;
            OvenNameTextBox.BackColor = Color.White;
            SwitchToAirTemperatureNumericUpDown.BackColor = Color.White;
        }

        // This function updates the filename in case the user sets the
        // directory before setting the batch name
        private bool UpdateDirectory()
        {
            string temp_filename = "";
            string temp_directory = "";

            DateTime date_time_copy = DateTime.Now;

            temp_filename = String.Join("", UserNameTextBox.Text, "_", BatchIDCodeTextBox.Text, "",
                    date_time_copy.ToString(format), ".tsv");

            temp_directory = Path.Combine(filepath, temp_filename);

            if (SaveLogFileDirectoryTextBox.Text != temp_directory)
            {
                directory = temp_directory;
                SaveLogFileDirectoryTextBox.Text = directory;
                return true;
            }
            return false;
        }
        
        // This function updates the WARNING box.
        public void UpdateWarningDisplay(string Warning, 
                                           ErrorCodeEnum ErrorCode)
        {
            if (abort_oven_program)
            {
                StopThreads_AndCloseOvenProgram();
                return;
            }

            try
            {

                WarningTextBox.Text = Warning;
                switch (ErrorCode)
                {
                    case ErrorCodeEnum.NoError:
                        
                        ReadTC_Button.BackColor = Color.GreenYellow;
                        StopTC_Button.BackColor = default(Color);
                        for (int i = 0; i < MAX_CHANCOUNT; i++)
                        {
                            if (this.Controls["Channel" + i.ToString("#0") + "TextBox"].ForeColor
                                == Color.Gray)
                            {
                                this.Controls["Channel" + i.ToString("#0") + "TextBox"].ForeColor
                                    = Color.LawnGreen;
                            }
                        }
                        break;

                    case ErrorCodeEnum.Error:

                        ReadTC_Button.BackColor = default(Color);
                        StopTC_Button.BackColor = Color.Red;
                        for (int i = 0; i < MAX_CHANCOUNT; i++)
                        {
                            this.Controls["Channel" + i.ToString("#0") + "TextBox"].ForeColor =
                                Color.Gray;
                        }
                        break;

                    default:
                        break;
                }
                return;
            }

            catch
            {
                MessageBox.Show(
                    "Update Error Warning Display has encountered an unexpected error!", 
                    "   ERROR!");
                return;
            }
        }

        // This function updates the  values and colors for the temperature 
        // readout and the times elapsed for the oven run. It uses a for loop to
        // set the values of the textboxes that shows the temperature values. It
        // shows an error using a message box.
        public void UpdateTemperatureDisplays(float[] TempData, float[] TempData_C)
        {
            if (abort_oven_program)
            {
                StopThreads_AndCloseOvenProgram();
                return;
            }

            float Outer = 0;
            float Inner = 0;
            float Tray_Outer_Edge = 0;
            float Tray_SZ_Edge = 0;
            float Tray_Inner_Edge = 0;
            float Tray_Outer_Center = 0;
            float Tray_SZ_Center = 0;
            float Tray_Inner_Center = 0;
            double Time;

            System.TimeSpan timespan = new System.TimeSpan();

            try
            {
                if (TempData[0] == 1234)
                {
                    for (int i = 0; i < MAX_CHANCOUNT; i++)
                    {
                        this.Controls["Channel" + i.ToString("#0") + "TextBox"].ForeColor =
                                Color.Gray;
                    }
                    return;
                }

                for (int i = 0; i < MAX_CHANCOUNT; i++)
                {
                    if (TempData[i] != -9999)
                    {
                        this.Controls["Channel" + i.ToString("#0") + "TextBox"].Text =
                            (" " + (int)(TempData[i])).ToString();
                        this.Controls["Channel" + i.ToString("#0") + "TextBox"].ForeColor =
                                Color.LawnGreen;
                    }
                    else
                    {
                        this.Controls["Channel" + i.ToString("#0") + "TextBox"].Text =
                                " ***";
                        this.Controls["Channel" + i.ToString("#0") + "TextBox"].ForeColor =
                                Color.Gray;
                    }

                    //Change the place holder temperature value (-9999) to 0
                    if (TempData_C[i] == -9999)
                        TempData_C[i] = 0;
                    
                    if (i == 0)
                    {
                        Tray_Outer_Edge = TempData_C[i];
                    }
                    else if (i == 1)
                    {
                        Tray_SZ_Edge = TempData_C[i];
                    }
                    else if (i == 2)
                    {
                        Tray_Inner_Edge = TempData_C[i];
                    }
                    else if (i == 3)
                    {
                        Tray_Inner_Center = TempData_C[i];
                    }
                    else if (i == 4)
                    {
                        Tray_SZ_Center = TempData_C[i];
                    }
                    else if (i == 5)
                    {
                        Tray_Outer_Center = TempData_C[i];
                    }
                    else if (i == 6)
                        Outer = TempData_C[i];
                    else if (i == 7)
                        Inner = TempData_C[i];
                }

                // Graphing section:
                if (temperature_start_check == ComponentCheckStatusEnum.Running)
                {
                    if (graph_check == 3)
                    {
                        foreach (var series in TempertureVsTimeGraph.Series)
                        {
                            series.Points.Clear();
                        }
                        graph_check = 1;
                    }

                    if (oven_check == ComponentCheckStatusEnum.Running)
                        timespan = DateTime.Now.Subtract(Start_Oven_Time);
                    else
                        timespan = DateTime.Now.Subtract(Start_Temp_Time);

                    if (oven_check == ComponentCheckStatusEnum.Running)
                        TimeElapsedTextBox.Text = timespan.ToString(@"h\:mm\:ss");

                    Time = timespan.TotalMinutes;
                                        
                    TempertureVsTimeGraph.Series[
                        Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                                    TemperatureVsTimeDataSeriesEnum.TrayOuter_Edge)].Points.AddXY(Time,
                        Tray_Outer_Edge);
                    
                    TempertureVsTimeGraph.Series[
                        Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                                    TemperatureVsTimeDataSeriesEnum.TraySampleZone_Edge)].Points.AddXY(Time,
                        Tray_SZ_Edge);
                    
                    TempertureVsTimeGraph.Series[
                        Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                                    TemperatureVsTimeDataSeriesEnum.TrayInner_Edge)].Points.AddXY(Time,
                        Tray_Inner_Edge);
                    
                    TempertureVsTimeGraph.Series[
                        Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                                    TemperatureVsTimeDataSeriesEnum.TrayOuter_Center)].Points.AddXY(Time,
                        Tray_Outer_Center);
                    
                    TempertureVsTimeGraph.Series[
                        Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                                    TemperatureVsTimeDataSeriesEnum.TraySampleZone_Center)].Points.AddXY(Time,
                        Tray_SZ_Center);
                    
                    TempertureVsTimeGraph.Series[
                        Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                                    TemperatureVsTimeDataSeriesEnum.TrayInner_Center)].Points.AddXY(Time,
                        Tray_Inner_Center);
                    
                    TempertureVsTimeGraph.Series[
                        Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                                    TemperatureVsTimeDataSeriesEnum.Outer)].Points.AddXY(Time, Outer);
                    
                    TempertureVsTimeGraph.Series[
                        Enum.GetName(typeof(TemperatureVsTimeDataSeriesEnum),
                                    TemperatureVsTimeDataSeriesEnum.Inner)].Points.AddXY(Time, Inner);
                    
                    this.Refresh();
                }

                return;
            }

            catch(Exception e)
            {
                MessageBox.Show(
                    String.Format(
                        "Update Temperature Display has encountered an unexpected error!{0}" +
                        "Error Message: {1}{0}" +
                        "Error Source: {2}{0}" +
                        "Stack Trace: {3}{0}",
                        Environment.NewLine,
                        e.Message,
                        e.Source,
                        e.StackTrace),
                    "   " + "ERROR!");

                return;
            }
        }

        public void UpdateAirStatusIndicators(AirChanStatusEnum air_state)
        {
            bool air_on = false;
            bool high_n2_on = false;
            bool low_n2_on = false;

            switch (air_state)
            { 
                case AirChanStatusEnum.Air:
                    air_on = true;
                    break;
                                    
                case AirChanStatusEnum.HighN2:
                    high_n2_on = true;
                    break;

                case AirChanStatusEnum.LowN2:
                    low_n2_on = true;
                    break;

                default:
                    //do nothing - air is off
                    break;
            }

            if (air_on)
                SetAirIndicator_ToOn();
            else
                SetAirIndicator_ToOff();

            if (high_n2_on)
                SetHighN2Indicator_ToOn();
            else
                SetHighN2Indicator_ToOff();

            if (low_n2_on)
                SetLowN2Indicator_ToOn();
            else
                SetLowN2Indicator_ToOff();

            this.BeginInvoke((Action)(() => this.Refresh()));

        }

        public void SetAirIndicator_ToOn()
        {
            AirOnOffIndicatorButton.BeginInvoke((Action)(() => AirOnOffIndicatorButton.BackColor = Color.GreenYellow));
        }

        public void SetAirIndicator_ToOff()
        {
            AirOnOffIndicatorButton.BeginInvoke((Action)(() => AirOnOffIndicatorButton.BackColor = Color.Black));
        }

        public void SetHighN2Indicator_ToOn()
        {
            HighN2OnOffIndicatorButton.BeginInvoke((Action)(() => HighN2OnOffIndicatorButton.BackColor = Color.GreenYellow));
        }

        public void SetHighN2Indicator_ToOff()
        {
            HighN2OnOffIndicatorButton.BeginInvoke((Action)(() => HighN2OnOffIndicatorButton.BackColor = Color.Black));
        }

        public void SetLowN2Indicator_ToOn()
        {
            LowN2OnOffIndicatorButton.BeginInvoke((Action)(() => LowN2OnOffIndicatorButton.BackColor = Color.GreenYellow));
        }

        public void SetLowN2Indicator_ToOff()
        {
            LowN2OnOffIndicatorButton.BeginInvoke((Action)(() => LowN2OnOffIndicatorButton.BackColor = Color.Black));
        }

        //Updates Form Controls to show current status for each oven heating coil element
        public void SetHeatingElementStatus(OvenChanEnum oven_element)
        {
            switch (oven_element)
            {
                case OvenChanEnum.AllOff:
                case OvenChanEnum.OvenOff:
                    outer_status = false;
                    sample_zone_status = false;
                    inner_status = false;
                    break;

                case OvenChanEnum.Inner:
                    inner_status = true;
                    break;

                case OvenChanEnum.Outer:
                    outer_status = true;
                    break;

                case OvenChanEnum.SampleZone:
                    sample_zone_status = true;
                    break;

                case OvenChanEnum.OvenOn:
                    outer_status = true;
                    sample_zone_status = true;
                    inner_status = true;
                    break;

                default:
                    //exit function
                    return;
            }

            if (outer_status)
                SetOuterIndicatorToOn();
            else
                SetOuterIndicatorToOff();

            if (inner_status)
                SetInnerIndicatorToOn();
            else
                SetInnerIndicatorToOff();

            if (sample_zone_status)
                SetSampleZoneIndicatorToOn();
            else
                SetSampleZoneIndicatorToOff();


        }

        public void SetHeatingElementStatus(OvenChanEnum oven_element,
                                            OvenHeatOnOffEnum element_on_or_off)
        {
            switch (oven_element)
            {
                case OvenChanEnum.AllOff:
                case OvenChanEnum.OvenOff:
                    outer_status = false;
                    sample_zone_status = false;
                    inner_status = false;
                    break;

                case OvenChanEnum.Inner:
                    if (element_on_or_off == OvenHeatOnOffEnum.On)
                    {
                        inner_status = true;
                    }
                    else
                    {
                        inner_status = false;                    
                    }
                    break;

                case OvenChanEnum.Outer:
                    if (element_on_or_off == OvenHeatOnOffEnum.On)
                    {
                        outer_status = true;
                    }
                    else
                    {
                        outer_status = false;
                    }
                    break;

                case OvenChanEnum.SampleZone:
                    if (element_on_or_off == OvenHeatOnOffEnum.On)
                    {
                        sample_zone_status = true;
                    }
                    else
                    {
                        sample_zone_status = false;
                    }
                    break;

                case OvenChanEnum.OvenOn:
                    outer_status = true;
                    sample_zone_status = true;
                    inner_status = true;
                    break;

                default:
                    //exit function
                    return;
            }

            if (outer_status)
                SetOuterIndicatorToOn();
            else
                SetOuterIndicatorToOff();

            if (inner_status)
                SetInnerIndicatorToOn();
            else
                SetInnerIndicatorToOff();

            if (sample_zone_status)
                SetSampleZoneIndicatorToOn();
            else
                SetSampleZoneIndicatorToOff();


        }

        private void SetOuterIndicatorToOn()
        {
            if (outer_zone_heater_disabled) return;
            
            if (Outer_HeatingElementStatusButton.InvokeRequired)
            {
                Outer_HeatingElementStatusButton.BeginInvoke(
                    ((Action)(() => Outer_HeatingElementStatusButton.BackColor = ElementOnColor)));
            }
            else
            {
                Outer_HeatingElementStatusButton.BackColor = ElementOnColor;
            }
        }

        private void SetOuterIndicatorToOff()
        {
            if (outer_zone_heater_disabled) return;

            if (Outer_HeatingElementStatusButton.InvokeRequired)
            {
                Outer_HeatingElementStatusButton.BeginInvoke(
                    ((Action)(() => Outer_HeatingElementStatusButton.BackColor = ElementOffColor)));
            }
            else
            {
                Outer_HeatingElementStatusButton.BackColor = ElementOffColor;
            }
        }

        private void SetInnerIndicatorToOn()
        {
            if (inner_zone_heater_disabled) return;

            if (Inner_HeatingElementStatusButton.InvokeRequired)
            {
                Inner_HeatingElementStatusButton.BeginInvoke(
                    ((Action)(() => Inner_HeatingElementStatusButton.BackColor = ElementOnColor)));
            }
            else
            {
                Inner_HeatingElementStatusButton.BackColor = ElementOnColor;
            }
        }

        private void SetInnerIndicatorToOff()
        {
            if (inner_zone_heater_disabled) return;

            if (Inner_HeatingElementStatusButton.InvokeRequired)
            {
                Inner_HeatingElementStatusButton.BeginInvoke(
                    ((Action)(() => Inner_HeatingElementStatusButton.BackColor = ElementOffColor)));
            }
            else
            {
                Inner_HeatingElementStatusButton.BackColor = ElementOffColor;
            }
        }

        private void SetSampleZoneIndicatorToOn()
        {
            if (sample_zone_heater_disabled) return;

            if (SampleZone_HeatingElementStatusButton.InvokeRequired)
            {
                SampleZone_HeatingElementStatusButton.BeginInvoke(
                    ((Action)(() => SampleZone_HeatingElementStatusButton.BackColor = ElementOnColor)));
            }
            else
            {
                SampleZone_HeatingElementStatusButton.BackColor = ElementOnColor;
            }
        }

        private void SetSampleZoneIndicatorToOff()
        {
            if (sample_zone_heater_disabled) return;

            if (SampleZone_HeatingElementStatusButton.InvokeRequired)
            {
                SampleZone_HeatingElementStatusButton.BeginInvoke(
                    ((Action)(() => SampleZone_HeatingElementStatusButton.BackColor = ElementOffColor)));
            }
            else
            {
                SampleZone_HeatingElementStatusButton.BackColor = ElementOffColor;
            }
        }

        private void EnableStartCoolingButton()
        {
            StartCoolingButton.Enabled = true;
        }

        #endregion  

        #region Form Control Values Validation Methods

        // This function checks to see if the user input parameters
        // are valid and useable (with the exception of directory)
        private ErrorCodeEnum CheckParameters(ParameterTypeEnum par_type_id)
        {
            if (par_type_id != ParameterTypeEnum.ThermoCouples &&
                par_type_id != ParameterTypeEnum.Oven)
            {
                MessageBox.Show(
                    String.Format(
                        "Invalid Parameter Type ({0}) submitted to Check Parameters Function.",
                        Enum.GetName(typeof(ParameterTypeEnum),par_type_id)),
                    "   ERROR!");

                return ErrorCodeEnum.Error;
            }

            int check = 0;

            if (String.IsNullOrWhiteSpace(UserNameTextBox.Text) || UserNameTextBox.Text == "User")
            {
                UserNameTextBox.BackColor = Color.MistyRose;
                check++;
            }
            else
                UserNameTextBox.BackColor = Color.White;

            if (String.IsNullOrWhiteSpace(BatchIDCodeTextBox.Text) ||
                    BatchIDCodeTextBox.Text == "_" || BatchIDCodeTextBox.Text ==
                    UserNameTextBox.Text)
            {
                BatchIDCodeTextBox.BackColor = Color.MistyRose;
                check++;
            }
            else
                BatchIDCodeTextBox.BackColor = Color.White;

            if (String.IsNullOrWhiteSpace(OvenNameTextBox.Text))
            {
                OvenNameTextBox.BackColor = Color.MistyRose;
                check++;
            }
            else
                OvenNameTextBox.BackColor = Color.White;

            if (par_type_id == ParameterTypeEnum.Oven)
            {
                try
                {
                    String[] delimeters = new String[] { ",", ";" };
                    String[] email_addresses = UserEmailTextBox.Text.Split(delimeters,
                        StringSplitOptions.RemoveEmptyEntries);

                    if (email_addresses.Length <= 0)
                    {
                        throw new Exception("Email Address cannot be blank.");
                    }

                    foreach (string addr in email_addresses)
                    {
                        if (!Regex.IsMatch(addr.Trim(), permissive_email_reg_ex))
                        {
                            throw new Exception(
                                String.Format(
                                    "Invalid Email Address: {0}{1}",
                                    Environment.NewLine,
                                    addr));
                        }
                    }

                    if (UserEmailTextBox.Text == "user@emailaddress.com")
                        throw new Exception("Please change the email address to a " +
                                            "different email from the default email address.");

                    UserEmailTextBox.BackColor = Color.White;
                }
                catch
                {
                    UserEmailTextBox.BackColor = Color.MistyRose;
                    check++;
                }

                try
                {
                    if (MaxTemperatureNumericUpDown.Value > 800 ||
                        MaxTemperatureNumericUpDown.Value < 20)
                    {
                        MaxTemperatureNumericUpDown.BackColor = Color.MistyRose;
                        check++;
                    }
                    else
                    {
                        if (MaxTemperatureNumericUpDown.Value >= 150 &&
                            CoolMethodComboBox.SelectedIndex == 0)
                            MessageBox.Show(
                                "Recommend using Nitrogen (N2) " +
                                    "cooling when heating samples to 150 °C and above.",
                                "   SUGGESTION");
                        MaxTemperatureNumericUpDown.BackColor = Color.White;
                    }
                }
                catch
                {
                    MaxTemperatureNumericUpDown.BackColor = Color.MistyRose;
                    check++;
                }

                try
                {
                    if (HoldTimeAtPeakNumericUpDown.Value > 100 ||
                        HoldTimeAtPeakNumericUpDown.Value < 5)
                    {
                        HoldTimeAtPeakNumericUpDown.BackColor = Color.MistyRose;
                        check++;
                    }
                    else
                        HoldTimeAtPeakNumericUpDown.BackColor = Color.White;
                }
                catch
                {
                    HoldTimeAtPeakNumericUpDown.BackColor = Color.MistyRose;
                    check++;
                }

                if (CoolMethodComboBox.SelectedItem == null)
                {
                    CoolMethodComboBox.BackColor = Color.MistyRose;
                    check++;
                }
                else if (CoolMethodComboBox.SelectedIndex == 1)
                {
                    SwitchToAirTemperatureNumericUpDown.Enabled = true;
                    try
                    {
                        if (SwitchToAirTemperatureNumericUpDown.Value > 500 ||
                            SwitchToAirTemperatureNumericUpDown.Value < 20)
                        {
                            SwitchToAirTemperatureNumericUpDown.BackColor = Color.MistyRose;
                            check++;
                        }
                        else
                            SwitchToAirTemperatureNumericUpDown.BackColor = Color.White;
                    }
                    catch
                    {
                        SwitchToAirTemperatureNumericUpDown.BackColor = Color.MistyRose;
                        check++;
                    }
                    CoolMethodComboBox.BackColor = Color.White;
                }
                else if (CoolMethodComboBox.SelectedIndex == 0)
                {
                    SwitchToAirTemperatureNumericUpDown.Value = 20;
                    SwitchToAirTemperatureNumericUpDown.BackColor = Color.White;
                    SwitchToAirTemperatureNumericUpDown.Enabled = false;
                    CoolMethodComboBox.BackColor = Color.White;
                }
                else
                    CoolMethodComboBox.BackColor = Color.White;

                try
                {
                    if (StopCoolingTemperatureNumericUpDown.Value > 50 ||
                        StopCoolingTemperatureNumericUpDown.Value < 25)
                    {
                        StopCoolingTemperatureNumericUpDown.BackColor = Color.MistyRose;
                        check++;
                    }
                    else
                        StopCoolingTemperatureNumericUpDown.BackColor = Color.White;
                }
                catch
                {
                    StopCoolingTemperatureNumericUpDown.BackColor = Color.MistyRose;
                    check++;
                }

                try
                {
                    if (!Program.pid_settings_form.ProgramPidSettings.IsChanged)
                    {
                        OpenPIDSettingsWindowButton.BackColor = Color.OrangeRed;
                        check++;
                    }
                    else
                    {
                        OpenPIDSettingsWindowButton.BackColor = DefaultBackColor;
                    }
                }
                catch
                {
                    OpenPIDSettingsWindowButton.BackColor = Color.MistyRose;
                    check++;
                }
            }

            if (check > 0)
            {
                return ErrorCodeEnum.Error;
            }
            else
            {
                return ErrorCodeEnum.NoError;
            }
        }

        #endregion

        #region Oven State Update & Maintenance Methods

        // This function copies the values of the oven run parameters that are 
        // set when starting the oven or thermocouple reader for the first time.
        private void CopyParameters()
        {
            User_copy = UserNameTextBox.Text;
            BatchCode_copy = BatchIDCodeTextBox.Text;
            UserEmail_copy = UserEmailTextBox.Text;
            UserPhone_copy = UserPhoneTextBox.Text;
            MaxTemp_copy = (float)MaxTemperatureNumericUpDown.Value;
            HoldTime_copy = new TimeSpan(0,(int)HoldTimeAtPeakNumericUpDown.Value,0);
            CoolMethod_copy = CoolMethodComboBox.SelectedIndex.ToString();
            directory_copy = directory;
            OvenName_copy = OvenNameTextBox.Text;
            SwitchTemp_copy = (float)SwitchToAirTemperatureNumericUpDown.Value;
            StopTemp_copy = (float)StopCoolingTemperatureNumericUpDown.Value;
            OvenStartTime_copy = OvenStartTimeTextBox.Text;
            TimeElapsed_copy = TimeElapsedTextBox.Text;
            HoldStartTime_copy = HoldStartTimeTextBox.Text;
            HoldTimeRemaining_copy = HoldTimeRemainingTextBox.Text;
            CoolStartTime_copy = CoolStartTimeTextBox.Text;
            CoolTimeElapsed_copy = CoolTimeElapsedTextBox.Text;
        }

        // This function checks to see if the user changed any of the 
        // parameters after stopping the oven or the thermocouple reader.
        private bool IsParametersChanged()
        {
            if (User_copy != UserNameTextBox.Text ||
            UserEmail_copy != UserEmailTextBox.Text ||
            BatchCode_copy != BatchIDCodeTextBox.Text ||
            UserPhone_copy != UserPhoneTextBox.Text ||
            MaxTemp_copy != (float)MaxTemperatureNumericUpDown.Value ||
            HoldTime_copy != new TimeSpan(0,(int)HoldTimeAtPeakNumericUpDown.Value,0) ||
            CoolMethod_copy != CoolMethodComboBox.SelectedIndex.ToString() ||
            directory_copy != directory ||
            SwitchTemp_copy != (float)SwitchToAirTemperatureNumericUpDown.Value ||
            StopTemp_copy != (float)StopCoolingTemperatureNumericUpDown.Value ||
            OvenName_copy != OvenNameTextBox.Text)
            {
                if (BatchCode_copy != BatchIDCodeTextBox.Text)
                    log_state = 0;
                return true;
            }

            return false;
        }

        #endregion

        #region Oven Process Abort Methods

        // This function closes the window and exits the program. It checks
        // if the threads are in the stopped state, and if not it stops them by
        // the value of check to -1. It then changes the color of the Stop TC
        // button to red to notify the user, and again checks to see if the
        // threads are stopped and aborted. If not, it counts to 10 and aborts
        // the thread using Thread.Abort() before closing.
        public void StopThreads_AndCloseOvenProgram()
        {
            if (close_program_inprogress) return;

            close_program_inprogress = true;

            if (oven_start_check != ComponentCheckStatusEnum.None || 
                temperature_start_check != ComponentCheckStatusEnum.None)
            {
                DialogResult dialogResult = 
                    MessageBox.Show(
                        "WARNING! This will exit the program and shut down all ongoing processes" +
                        " Are you sure you want to exit?", 
                        "   Exit Application?", 
                        MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.No)
                {
                    abort_oven_program = false;
                    close_program_inprogress = false;
                    return;
                }
            }

            if (temperature_start_check == ComponentCheckStatusEnum.Running && 
                read_check != ComponentCheckStatusEnum.None)
            {
                this.StopTC_Button.Text = "Stopping";
                this.StopTC_Button.BackColor = Color.Firebrick;
            }

            if (oven_start_check == ComponentCheckStatusEnum.Running && 
                oven_check != ComponentCheckStatusEnum.None)
            {
                this.StopOvenButton.Text = "Stopping";
                this.StopOvenButton.BackColor = Color.Firebrick;
            }

            this.Refresh();

            Thread.Sleep(1000);

            if ((temperature_read_thread != null && 
                 temperature_read_thread.ThreadState != ThreadState.Stopped) || 
                (oven_control_thread != null &&
                 oven_control_thread.ThreadState != ThreadState.Stopped) ||
                (log_thread != null && 
                 log_thread.ThreadState != ThreadState.Stopped))
            {
                read_check = ComponentCheckStatusEnum.Aborted;
                oven_check = ComponentCheckStatusEnum.Aborted;
                log_check = ComponentCheckStatusEnum.Aborted;
                oven_cooling_check = ComponentCheckStatusEnum.Aborted;
            }


            int counter = 0;

            while ((temperature_read_thread != null && 
                    temperature_read_thread.ThreadState != ThreadState.Stopped && 
                    temperature_read_thread.ThreadState != ThreadState.Aborted) || 
                   (oven_control_thread != null &&
                   oven_control_thread.ThreadState != ThreadState.Stopped &&
                   oven_control_thread.ThreadState != ThreadState.Aborted) ||
                   (log_thread != null && 
                    log_thread.ThreadState != ThreadState.Stopped && 
                    log_thread.ThreadState != ThreadState.Aborted))
            {
                counter++;

                if (counter > 10)
                {
                    if (temperature_read_thread != null)
                    {
                        temperature_read_thread.Abort();
                    }
                    if (oven_control_thread != null)
                    {
                        oven_control_thread.Abort();
                    }
                    if (log_thread != null)
                    {
                        log_thread.Abort();
                    }
                }

                oven_cooling_check = ComponentCheckStatusEnum.Aborted;

                Thread.Sleep(100);
            }

            log_thread = null;

            this.Close();
        }

        #endregion
        
        #region Email / SMS Update Methods

        // This function is used to send out the oven notification emails.
        public void SendEmail(string subject, string body)
        {
            try
            {
                //Set To Email address
                NetworkCredential smtp_credentials =
                    new NetworkCredential(from_address.Address, email_from_password);
                

                // SMPT host setup
                smtp_client.Host = "smtp.gmail.com";
                smtp_client.Port = 465;
                smtp_client.UseDefaultCredentials = false;
                smtp_client.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp_client.Credentials = smtp_credentials;
                smtp_client.EnableSsl = true;

                string email_addr = UserEmailTextBox.Text.Replace(" ", string.Empty);
                email_addr = email_addr.Replace(";", ",");
                string email_name = UserNameTextBox.Text.Replace(" ", String.Empty);
                email_name = email_name.Replace(";",",");

                String[] addr_array = email_addr.Split(new String[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                String[] name_array = email_name.Split(new String[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                int num_addr = addr_array.GetLength(0);
                int num_names = name_array.GetLength(0);

                to_address = new MailAddress(addr_array[0], name_array[0]);

                for(int i = 0;i < num_addr; i++)
                {
                    int name_index = i;
                    if (name_index > num_names) name_index = num_names - 1;

                    MailAddress temp_addr = new MailAddress(addr_array[i],
                                                            name_array[name_index]);

                    email_message.To.Add(temp_addr);    
                }                
                
                email_message.Subject = subject;
                email_message.Body = body;
                email_message.From = from_address;
                
                smtp_client.SendAsync(email_message, true);  
            }
            catch (Exception ex)
            {
                if (log_check != ComponentCheckStatusEnum.Aborted)
                    MessageBox.Show(ex.ToString(), "   ERROR!");
            }
        }

        #endregion

        #region Data Logger

        // This function updates the temperature log text file. It returns if
        // the number of channels is invalid. Otherwise, it uses a while loop to
        // output temperature readings to a .txt file in the same directory as
        // the program folder. It shows an error if there is an open channel.
        public void UpdateLog()
        {
            string subject = "";
            string body = "";

            string separator = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" +
                            "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" +
                            "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" +
                            "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";

            while (log_check == ComponentCheckStatusEnum.Running &&
                   log_thread.ThreadState != ThreadState.AbortRequested)
            {
                if (log_state == LogStatusEnum.NotLogged)
                {
                    log_state = LogStatusEnum.Logged;

                    string[] text1 = 
                        {"Temperature and events log file.", 
                            " ",
                            "Created on " + 
                            Start_Log_File_Time.ToString(format),
                            " ",
                            "Temperature read approx. every 5 " +
                            "seconds in degrees Celsius.", 
                            " ",
                            " "};

                    System.IO.File.WriteAllLines(directory, text1);

                    string text2 = "Time\t";

                    for (int i = 0; i < MAX_CHANCOUNT; i++)
                    {
                        text2 += String.Format("Channel{0}", i) + "\t";
                    }

                    using (StreamWriter sw = File.AppendText(directory))
                    {
                        sw.WriteLine(text2);
                        sw.WriteLine(" ");
                    }
                }

                if (HeatingStarted && !HeatingStarted_logged)
                {
                    string temp = DateTime.Now.ToString("T") + "\t\t" +
                            "HEATING STARTED!";

                    using (StreamWriter sw = File.AppendText(directory))
                    {
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                        sw.WriteLine(temp);
                        sw.WriteLine("\t\t\tUser: " + User_copy);
                        sw.WriteLine("\t\t\tBatch code: " + BatchCode_copy);
                        sw.WriteLine("\t\t\tMax. Temperature: " + MaxTemp_copy);
                        sw.WriteLine("\t\t\tHold Time: " + HoldTime_copy);
                        sw.WriteLine("\t\t\tCooling method: " + CoolMethod_copy);
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                    }

                    subject = "Heating has started on " + OvenName_copy +
                        " oven!";

                    body = User_copy + ",\n\n" + subject +
                        "\n\nParameters are: \nBatch Code: " + BatchCode_copy +
                        "\nMax.Temp: " + MaxTemp_copy +
                        "\nHold Time: " + HoldTime_copy +
                        "\nCooling Method: " + CoolMethod_copy +
                        "\nTime: " + OvenStartTime_copy + "\n";
                    SendEmail(subject, body);
                    HeatingStarted_logged = true;
                }

                else if (HeatingStopped && !HeatingStopped_logged)
                {
                    string temp = DateTime.Now.ToString("T") + "\t\t" +
                            "HEATING STOPPED!";

                    using (StreamWriter sw = File.AppendText(directory))
                    {
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                        sw.WriteLine(temp);
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                    }

                    subject = "Heating has stopped on " + OvenName_copy +
                        " oven!";
                    body = User_copy + ",\n\n" + subject;
                    SendEmail(subject, body);
                    HeatingStopped_logged = true;
                }

                else if (MaxTempPlateauReached && !MaxTempReached_logged)
                {
                    string temp = DateTime.Now.ToString("T") + "\t\t" +
                            "MAX. TEMP REACHED!";

                    using (StreamWriter sw = File.AppendText(directory))
                    {
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                        sw.WriteLine(temp);
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                    }

                    subject = "Max. temp reached on " + OvenName_copy +
                        " oven!";
                    body = User_copy + ",\n\n" + subject;
                    SendEmail(subject, body);
                    MaxTempReached_logged = true;
                }

                else if (CoolingStarted && !CoolingStarted_logged)
                {
                    string temp = DateTime.Now.ToString("T") + "\t\t" +
                            "COOLING STARTED!";

                    using (StreamWriter sw = File.AppendText(directory))
                    {
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                        sw.WriteLine(temp);
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                    }

                    subject = "Cooling has started on " + OvenName_copy +
                        " oven!";
                    body = User_copy + ",\n\n" + subject;
                    SendEmail(subject, body);
                    CoolingStarted_logged = true;
                }

                else if (IsCriticalApplicationError && !IsCriticalApplicationError_logged)
                {
                    string temp = DateTime.Now.ToString("T") + "\t\t" +
                            "CRITICAL ERROR!";

                    using (StreamWriter sw = File.AppendText(directory))
                    {
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                        sw.WriteLine(temp);
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                    }

                    subject = "Critical error on " + OvenName_copy + " oven!";
                    body = User_copy + ",\n\n" + subject + "\n" +
                        CriticalErrorMessage;
                    SendEmail(subject, body);
                    IsCriticalApplicationError_logged = true;
                }

                else if (OvenRunDone && !OvenRunDone_logged)
                {
                    string temp = DateTime.Now.ToString("T") + "\t\t" +
                            "OVEN RUN COMPLETE!";

                    using (StreamWriter sw = File.AppendText(directory))
                    {
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                        sw.WriteLine(temp);
                        sw.WriteLine("\t\t\tUser: " + User_copy);
                        sw.WriteLine("\t\t\tBatch code: " + BatchCode_copy);
                        sw.WriteLine("\t\t\tMax. Temperature: " + MaxTemp_copy);
                        sw.WriteLine("\t\t\tHold Time: " + HoldTime_copy);
                        sw.WriteLine("\t\t\tCooling method: " + CoolMethod_copy);
                        sw.WriteLine(" ");
                        sw.WriteLine(separator);
                        sw.WriteLine(" ");
                    }

                    subject = "Oven run completed on " + OvenName_copy + " oven!";
                    body = User_copy + ",\n\n" + subject + "\n";
                    SendEmail(subject, body);
                    
                    //Stop Logging -- Oven run is done
                    log_check = ComponentCheckStatusEnum.None;
                }

                string text3 = DateTime.Now.ToString("T") + "\t";

                for (int i = 0; i < MAX_CHANCOUNT; i++)
                {
                    text3 += CurrentTemperatures[i].ToString() + "\t";
                }


                using (StreamWriter sw = File.AppendText(directory))
                {
                    sw.WriteLine(text3);
                }

                Thread.Sleep(4800);
            }

            OvenRunDone = false;
            OvenRunDone_logged = false;
        }

        #endregion

        #region Set Air & Oven Relays

        // This function turns on the heater coils according to the string
        // parameter passed to it. Usable string parameters are: "all", "0",
        // "1", "2", and "none"; "all" and "none" turn on/off all the coils,
        // while "0", "1", "2" turns on the coil corresponding to that channel
        // output from the USB-TC device.
        private void SetChannelOutput(Object state)
        {
            try
            {
                //Try to cast the input as an Oven Channel Status Enumerator
                //If there is no error, then the input IS an Oven Channel Status Enumerator instance
                OvenChanEnum oven_state = (OvenChanEnum)state;
                SetOvenChannelOutput(oven_state);
            }
            catch
            {
                try
                {
                    //Otherwise,
                    //Try to cast the input as an Air Channel Status Enumerator
                    //If there is no error, then the input IS an Air Channel Status Enumerator instance
                    AirChanStatusEnum air_relay_state = (AirChanStatusEnum)state;
                    SetAirRelayChannelOutput(air_relay_state);
                }
                catch
                { 
                    //Input is neither an air-relay state or an oven channel state
                    //Pop-up an error
                    MessageBox.Show(
                        "Invalid Set Channel Output State submitted to SetChannelOutput.",
                        "    Error!");
                }
            
            }
            
        }

        //This function turns the oven heating coils on and off
        //using Digital Output TTL logic
        //Uses values from the OvenChanStatusEnum enumerator to
        //choose its actions
        private void SetOvenChannelOutput(OvenChanEnum state)
        {
            MccDaq.ErrorInfo RetVal;

            switch (state)
            {
                case OvenChanEnum.OvenOn:
                    for (int i = 0; i < 3; i++)
                    {
                        if ((i == inner_zone_heater_index &&
                             inner_zone_heater_disabled) ||
                            (i == sample_zone_heater_index &&
                             sample_zone_heater_disabled) ||
                            (i == outer_zone_heater_index &&
                             outer_zone_heater_disabled))
                        {
                            //Toggle heating element off
                            //Heating element has been programmatically disabled
                            RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                                 i,
                                                 DigitalLogicState.Low);
                            IsError(this, RetVal, true);
                        }
                        else
                        {
                            //Toggle heating element on
                            RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                                 i,
                                                 DigitalLogicState.High);
                            IsError(this, RetVal, true);
                        }
                    }
                    break;

                case OvenChanEnum.OvenOff:
                    for (int i = 0; i < 3; i++)
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             i,
                                             DigitalLogicState.Low);
                        IsError(this, RetVal, true);
                    }
                    break;

                case OvenChanEnum.AllOff:
                    for (int i = 0; i < 3; i++)
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             i,
                                             DigitalLogicState.Low);
                        IsError(this, RetVal, true);
                    }
                    break;

                case OvenChanEnum.Inner:

                    if (inner_zone_heater_disabled)
                    {
                        //Turn Heating Element off
                        //Has been programmatically disabled
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             inner_zone_heater_index,
                                             DigitalLogicState.Low);
                        IsError(this, RetVal, true);
                    }
                    else
                    {
                        //Turn heating element on
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort,
                                             inner_zone_heater_index,
                                             DigitalLogicState.High);
                        IsError(this, RetVal, true);
                    }
                    break;

                case OvenChanEnum.SampleZone:

                    if (sample_zone_heater_disabled)
                    {
                        //Turn Heating Element off
                        //Has been programmatically disabled
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort,
                                             sample_zone_heater_index,
                                             DigitalLogicState.Low);
                        IsError(this, RetVal, true);
                    }
                    else
                    {
                        //Turn heating element on
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             sample_zone_heater_index,
                                             DigitalLogicState.High);
                        IsError(this, RetVal, true);
                    }
                    break;

                case OvenChanEnum.Outer:

                    if (outer_zone_heater_disabled)
                    {
                        //Turn Heating Element off
                        //Has been programmatically disabled
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             outer_zone_heater_index,
                                             DigitalLogicState.Low);
                        IsError(this, RetVal, true);
                    }
                    else
                    {
                        //Turn heating element on
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort,
                                             outer_zone_heater_index,
                                             DigitalLogicState.High);
                        IsError(this, RetVal, true);
                    }
                    break;

                default:
                    RetVal = new ErrorInfo(1);
                    MessageBox.Show(
                        "Set Oven Channel Output error -- invalid Oven Channel Output State input.",
                        "   " + "ERROR!");
                    break;
            }

            RetVal = oven_board.DConfigBit(DigitalPortType.AuxPort, 0,
                DigitalPortDirection.DigitalOut);

            if (IsError(this, RetVal, true) == ErrorCodeEnum.NoError)
            {
                oven_status = state;
                SetHeatingElementStatus(state);
            }
        }

        private void SetupThermoCoupleDAQBoard()
        {
            thermocouple_board = new MccBoard(BoardNum);
        }

        private void SetupOvenDAQBoard()
        {
            oven_board = new MccBoard(BoardNum);

            ConfigureOvenBoardDioChannelsForOutput();
        }

        private void ConfigureOvenBoardDioChannelsForOutput()
        {
            MccDaq.ErrorInfo RetVal = new MccDaq.ErrorInfo();

            if (oven_board == null) return;

            for (int i = 0; i < 7; i++)
            {
                RetVal = oven_board.DConfigBit(DigitalPortType.AuxPort, i,
                    DigitalPortDirection.DigitalOut);
                IsError(this, RetVal, true);
            }
        }

        private void SetOvenChannelOutput(OvenChanEnum channel, OvenHeatOnOffEnum on_or_off)
        {
            MccDaq.ErrorInfo RetVal = new ErrorInfo();

            switch (channel)
            {
                case OvenChanEnum.OvenOn:
                    for (int i = 0; i < 3; i++)
                    {
                        if ((i == inner_zone_heater_index &&
                             inner_zone_heater_disabled) ||
                            (i == sample_zone_heater_index &&
                             sample_zone_heater_disabled) ||
                            (i == outer_zone_heater_index &&
                             outer_zone_heater_disabled))
                        {
                            //Toggle heating element off
                            //Heating element has been programmatically disabled
                            RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                                 i,
                                                 DigitalLogicState.Low);
                            IsError(this, RetVal, true);
                        }
                        else
                        {
                            //Toggle heating element on
                            RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                                 i,
                                                 DigitalLogicState.High);
                            IsError(this, RetVal, true);
                        }
                    }
                    break;

                case OvenChanEnum.OvenOff:
                    for (int i = 0; i < 3; i++)
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             i,
                                             DigitalLogicState.Low);
                        IsError(this, RetVal, true);
                    }
                    break;

                case OvenChanEnum.AllOff:
                    for (int i = 0; i < 3; i++)
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             i,
                                             DigitalLogicState.Low);
                        IsError(this, RetVal, true);
                    }
                    break;

                case OvenChanEnum.Inner:

                    if (on_or_off == OvenHeatOnOffEnum.On &&
                        !inner_zone_heater_disabled)
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort,
                                             inner_zone_heater_index,
                                             DigitalLogicState.High);
                    }
                    else
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             inner_zone_heater_index,
                                             DigitalLogicState.Low);
                    }

                    IsError(this, RetVal, true);
                    break;

                case OvenChanEnum.SampleZone:
                    
                    if (on_or_off == OvenHeatOnOffEnum.On &&
                        !sample_zone_heater_disabled)
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort,
                                             sample_zone_heater_index,
                                             DigitalLogicState.High);
                    }
                    else
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             sample_zone_heater_index,
                                             DigitalLogicState.Low);
                    }

                    IsError(this, RetVal, true);
                    break;

                case OvenChanEnum.Outer:
                    
                    if (on_or_off == OvenHeatOnOffEnum.On &&
                        !outer_zone_heater_disabled)
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 
                                             outer_zone_heater_index,
                                             DigitalLogicState.High);
                    }
                    else
                    {
                        RetVal = oven_board.DBitOut(DigitalPortType.AuxPort,
                                             outer_zone_heater_index,
                                             DigitalLogicState.Low);
                    }

                    IsError(this, RetVal, true);
                    break;

                default:
                    RetVal = new ErrorInfo(1);
                    MessageBox.Show(
                        "Set Oven Channel Output error -- invalid Oven Channel Output State input.",
                        "   " + "ERROR!");
                    break;
            }

            RetVal = oven_board.DConfigBit(DigitalPortType.AuxPort, 0,
                DigitalPortDirection.DigitalOut);

            if (IsError(this, RetVal, true) == ErrorCodeEnum.NoError)
            {
                oven_status = channel;
                SetHeatingElementStatus(channel, on_or_off);
            }
        }

        //This function toggles the Air/Nitrogen Relay Valves open and closed
        //using Digital Output TTL logic
        //Uses values from the AirChanStatusEnum enumerator to
        //choose its actions
        private void SetAirRelayChannelOutput(AirChanStatusEnum state)
        {
            MccDaq.ErrorInfo RetVal;

            //Turn all channels off, first
            for (int i = 3; i < 6; i++)
            {
                RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, i,
                    DigitalLogicState.Low);
                IsError(this, RetVal, true);
            }
            
            switch (state)
            {
                case AirChanStatusEnum.AirOff:
                    //Already done above.
                    break;

                case AirChanStatusEnum.LowN2:
                    RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 5,
                        DigitalLogicState.High);
                    IsError(this, RetVal, true);
                    break;

                case AirChanStatusEnum.HighN2:
                    RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 3,
                        DigitalLogicState.High);
                    IsError(this, RetVal, true);
                    break;

                case AirChanStatusEnum.Air:
                    RetVal = oven_board.DBitOut(DigitalPortType.AuxPort, 4,
                        DigitalLogicState.High);
                    IsError(this, RetVal, true);
                    break;

                default:
                    RetVal = new ErrorInfo(1);
                    MessageBox.Show("Set Air Relay Channel Output error -- invalid Air Relay Channel output state inputed.",
                        "   " + "ERROR!");
                    break;
            }

            air_status = state;
            UpdateAirStatusIndicators(state);

        }

        #endregion

        #region Temperature PID Conroller & Temperature Input Reader

        // This function reads and displays the temperature inputs by detecting
        // the buttons clicked, by detecting the value of the variable "check".
        // It runs a while loop to keep checking for button presses, and using
        // Thread.Sleep(), it checks buttons, reads and displays temperatures in
        // the selected scale at a set frequency (2 Hz). The loop ends when the
        // value of check is set to -1 by clicking the Exit button. Temperatures
        // are read using the TInScan function, and an array is passed to the 
        // update_temp() function to display them using a delegate.
        private void ReadTemperature()
        {
            while (read_check != ComponentCheckStatusEnum.Aborted)
            {
                int count = 0;
                float[] TempData = new float[MAX_CHANCOUNT];
                float[] TempData_C = new float[MAX_CHANCOUNT];

                MccDaq.ErrorInfo RetVal;

                Update_Temperature update_temp_del = new
                    Update_Temperature(UpdateTemperatureDisplays);

                Update_Warning update_warning_del = new
                    Update_Warning(UpdateWarningDisplay);

                Reset_Checkboxes resetcheckboxes_del = new
                    Reset_Checkboxes(ResetThemocoupleInputStatusCheckboxes);

                for (int i = 0; i < MAX_CHANCOUNT; i++)
                {
                    if (((CheckBox)this.Controls["Channel" +
                        i.ToString("#0") + "CheckBox"]).Enabled == true &&
                        ((CheckBox)this.Controls["Channel" +
                        i.ToString("#0") + "CheckBox"]).Checked == true)
                    {
                        count++;
                    }
                }

                if (count == 0)
                {
                    TempData = new float[MAX_CHANCOUNT] {-9999, -9999,
                        -9999, -9999, -9999, -9999, -9999, -9999};

                    //Array.Copy(TempData, TempData_C, 8);

                    not_showing_temp = true;
                    //BeginInvoke(update_temp_del, TempData, TempData_C);
                }
                else
                    not_showing_temp = false;

                if (read_check == ComponentCheckStatusEnum.Running &&
                    not_showing_temp == false)
                {
                    for (int i = 0; i < MAX_CHANCOUNT; i++)
                    {
                        if (((CheckBox)this.Controls["Channel" +
                            i.ToString("#0") + "CheckBox"]).Checked == true) // null object reference error, suggests using "new" keyword
                        {
                            RetVal = thermocouple_board.TIn(i, TempScale.Celsius,
                                out TempData_C[i],
                                ThermocoupleOptions.Filter);
                            IsError(this, RetVal, true);
                        }
                        else
                            TempData_C[i] = -9999;

                        while (IsLockedCurrentTemp_Array)
                        { 
                            //Nothing - pause till lock lifted
                        }

                        CurrentTemperatures[i] = TempData_C[i];

                        if (PlateauTemperatureWindow == null)
                            PlateauTemperatureWindow = new List<float[]>();

                        AddToMaxTemperaturePlateau(ref TempData_C);
                        MaxTempPlateauReached = IsAtSetTemperaturePlateau();
                    }

                    switch (my_TempScale)
                    {
                        case TempScale.Fahrenheit:
                            for (int i = 0; i < MAX_CHANCOUNT; i++)
                            {
                                if (((CheckBox)this.Controls["Channel" +
                                    i.ToString("#0") + "CheckBox"]).Checked == true)
                                {
                                    RetVal = thermocouple_board.TIn(i, TempScale.Fahrenheit,
                                        out TempData[i],
                                        ThermocoupleOptions.Filter);
                                    IsError(this, RetVal, true);
                                }
                                else
                                    TempData[i] = -9999;
                            }
                            break;

                        case TempScale.Kelvin:
                            for (int i = 0; i < MAX_CHANCOUNT; i++)
                            {
                                if (((CheckBox)this.Controls["Channel" +
                                    i.ToString("#0") + "CheckBox"]).Checked == true)
                                {
                                    RetVal = thermocouple_board.TIn(i, TempScale.Kelvin,
                                        out TempData[i],
                                        ThermocoupleOptions.Filter);
                                    IsError(this, RetVal, true);
                                }
                                else
                                    TempData[i] = -9999;
                            }
                            break;

                        default:
                            for (int i = 0; i < MAX_CHANCOUNT; i++)
                            {
                                if (((CheckBox)this.Controls["Channel" +
                                    i.ToString("#0") + "CheckBox"]).Checked == true)
                                {
                                    RetVal = thermocouple_board.TIn(i, TempScale.Celsius,
                                        out TempData[i],
                                        ThermocoupleOptions.Filter);
                                    IsError(this, RetVal, true);
                                }
                                else
                                    TempData[i] = -9999;
                            }
                            break;
                    }

                    BeginInvoke(update_temp_del, TempData, TempData_C);
                    System.Threading.Thread.Sleep(500);
                }

                if ((read_check == ComponentCheckStatusEnum.None &&
                     oven_check == ComponentCheckStatusEnum.Running) ||
                    (not_showing_temp == true &&
                     oven_check == ComponentCheckStatusEnum.Running))
                {
                    string temp =
                        String.Format(
                            "WARNING! Temperature readings are no longer live during an oven run!{0}" +
                            "Automatically fixing this in {1} seconds...",
                            Environment.NewLine,
                            ((int)(warning_time / 2)).ToString("0"));


                    BeginInvoke(update_warning_del, temp, 1);

                    Thread.Sleep(500);

                    warning_time--;

                    if (warning_time <= 0)
                    {
                        read_check = ComponentCheckStatusEnum.Running;
                        BeginInvoke(resetcheckboxes_del);
                        not_showing_temp = false;
                        BeginInvoke(update_warning_del, " ", 0);
                        warning_time = 20;
                    }
                }

                // Make sure that the text is grayed out when the Stop TC button
                // is in its clicked state.
                if (read_check == ComponentCheckStatusEnum.None &&
                    oven_check != ComponentCheckStatusEnum.Running)
                {
                    TempData[0] = 1234;
                    TempData_C[0] = 1234;
                    BeginInvoke(update_temp_del, TempData, TempData_C);
                    Thread.Sleep(500);
                }
            }
        }

        private void AddToMaxTemperaturePlateau(ref float[] temperature_data)
        {
            if (PlateauTemperatureWindow == null)
                PlateauTemperatureWindow = new List<float[]>();

            if (PlateauTemperatureWindow.Count < PlateauWindowLength)
            {
                PlateauTemperatureWindow.Add(temperature_data);
            }
            else
            {
                for (int i = 0; i < PlateauTemperatureWindow.Count - 1; i++)
                {
                    PlateauTemperatureWindow[i] = PlateauTemperatureWindow[i + 1];
                }

                PlateauTemperatureWindow[PlateauTemperatureWindow.Count - 1] = temperature_data;
            }
        }

        private Boolean IsAtSetTemperaturePlateau()
        {
            Boolean RetVal = false;

            float set_temperature = MaxTemp_copy;

            if (PlateauTemperatureWindow == null) return false;
            if (PlateauTemperatureWindow.Count < PlateauWindowLength) return false;

            float[] averages = new float[MAX_CHANCOUNT] { 0, 0, 0, 0, 0, 0, 0, 0 };
            double[] stddev = new double[MAX_CHANCOUNT] { 0, 0, 0, 0, 0, 0, 0, 0 };

            //calculate averages
            foreach (float[] TempData_C in PlateauTemperatureWindow)
            {
                for (int i = 0; i < MAX_CHANCOUNT - NUM_WALLCOUNT; i++)
                {
                    averages[i] += TempData_C[i] / PlateauWindowLength;
                }               
            }

            //calculate standard deviations
            foreach (float[] TempData_C in PlateauTemperatureWindow)
            {
                for (int i = 0; i < MAX_CHANCOUNT - NUM_WALLCOUNT; i++)
                {
                    stddev[i] += Math.Pow(averages[i] - TempData_C[i], 2);
                }
            }

            for (int i = 0; i < MAX_CHANCOUNT - NUM_WALLCOUNT; i++)
            {
                stddev[i] = Math.Sqrt(stddev[i]) / PlateauWindowLength;
                                
                if (stddev[i] < PlateauTemperatureDeviationThreshold &&
                    Math.Abs(averages[i] - set_temperature) / set_temperature < PercentageOffSetTemperature / 100)
                {
                    if (i == 0)
                    {
                        RetVal = true;
                    }
                    else
                    {
                        RetVal = RetVal & true;
                    }
                }
                else
                {
                    RetVal = RetVal & false;
                }
            }
            
            return RetVal;
        }



        // This function implements PID control in the oven, by using three
        // USB-TC output channels. It calls the functions channel_out() to 
        // control the heater coils according to the known setup.
        private void PID_control()
        {
            int count = 0;
            
            oven_pid_settings = Program.pid_settings_form.ProgramPidSettings;

            set_point_temperature = MaxTemp_copy;

            double[] CurrentTemp_EachZone = new double[3] { 0, 0, 0 };

            double[] PreviousTemp_EachZone = new double[3] { 0, 0, 0 };

            int heat_increment_time_in_milliseconds = (int)((oven_pid_settings.cycle_time_in_seconds / 110) * 1000);

            //Heat per each zone
            int[] heat_EachZone = new int[3] { 0, 0, 0 };

            //rate correction per each zone
            double[] rate_correction_EachZone = new double[3] { 0, 0, 0 };

            bool fill_n2_check = false;
            bool n2_running_check = false;

            bool outer_coil_on = false;
            bool inner_coil_on = false;
            bool sample_zone_coil_on = false;
                        
            SetOvenChannelOutput(OvenChanEnum.AllOff);   // start with all oven channels off
            SetAirRelayChannelOutput(AirChanStatusEnum.AirOff);     //Start with air switched off

            if (CoolMethod_copy == "1" &&
                    n2_running_check == false)
            {
                // Flow high pressure nitrogen for 10 seconds to fill
                // the oven cavity with nitrogen
                if (fill_n2_check == false)
                {
                    count = 0;
                    SetAirRelayChannelOutput(AirChanStatusEnum.HighN2);
                    while (count <= 10)
                    {
                        Thread.Sleep(1000);
                        count++;
                    }
                    fill_n2_check = true;
                    SetAirRelayChannelOutput(AirChanStatusEnum.AirOff);
                }

                // Low pressure N2 trickle begins
                SetAirRelayChannelOutput(AirChanStatusEnum.LowN2);
                n2_running_check = true;

                bool low_n2_adjusted = false;

                while (!low_n2_adjusted)
                { 
                    MessageBox.Show(
                        "Please change the Low N2 Gauge Level (by twisting the Gauge now) until " +
                        "the black indicator bead is beween the two black arrows. ",
                        "Set Gauge Level", 
                        MessageBoxButtons.OK);

                    DialogResult result =
                        MessageBox.Show(
                            "Is the Low N2 Gauge set to the correct level?",
                            "Verify Gauge Level",
                            MessageBoxButtons.YesNo);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        low_n2_adjusted = true;
                    }
                }
            }

            EnableStartOvenButton start_oven_button_del = new EnableStartOvenButton(EnableStartCoolingButton);
            StartCoolingButton.BeginInvoke(start_oven_button_del);

            HeatingStarted = true;

            while (oven_check == ComponentCheckStatusEnum.Running)
            {
                Heat_Time = DateTime.Now;

                DateTime start_heat_increment;
                DateTime end_heat_increment;

                for (int i = 1; i <= 100; i++)
                {
                    start_heat_increment = DateTime.Now;

                    if (heat_EachZone[0] > 0)
                    {
                        if (!outer_coil_on)
                        {
                            SetOvenChannelOutput(OvenChanEnum.Outer, OvenHeatOnOffEnum.On);
                            outer_coil_on = true;
                        }
                    }
                    else
                    {
                        SetOvenChannelOutput(OvenChanEnum.Outer, OvenHeatOnOffEnum.Off);
                        outer_coil_on = false;
                    }

                    if (heat_EachZone[1] > 0)
                    {
                        if (!sample_zone_coil_on)
                        {
                            SetOvenChannelOutput(OvenChanEnum.SampleZone, OvenHeatOnOffEnum.On);
                            sample_zone_coil_on = true;
                        }
                    }
                    else
                    {
                        SetOvenChannelOutput(OvenChanEnum.SampleZone, OvenHeatOnOffEnum.Off);
                        sample_zone_coil_on = false;
                    }

                    if (heat_EachZone[2] > 0)
                    {
                        if (!inner_coil_on)
                        {
                            SetOvenChannelOutput(OvenChanEnum.Inner, OvenHeatOnOffEnum.On);
                            inner_coil_on = true;
                        }
                    }
                    else
                    {
                        SetOvenChannelOutput(OvenChanEnum.Inner, OvenHeatOnOffEnum.Off);
                        inner_coil_on = false;
                    }

                    for (int j = 0; j < 3; j++)
                    {
                        heat_EachZone[j]--;
                    }

                    end_heat_increment = DateTime.Now;

                    TimeSpan heat_increment_duration = (end_heat_increment - start_heat_increment);

                    if (heat_increment_duration.TotalMilliseconds < heat_increment_time_in_milliseconds)
                    {
                        System.Threading.Thread.Sleep((int)(heat_increment_time_in_milliseconds - heat_increment_duration.TotalMilliseconds));
                    }

                    //If Set Temp Reached and Start Set Temp Hold Time not set
                    //Set Start Set Temp Hold Time
                    if (MaxTempPlateauReached &&
                        !MaxTemp_StartTime_Set)
                    {
                        MaxTemp_StartTime = DateTime.Now;
                        MaxTemp_StartTime_Set = true;
                    }

                    //If Set Temp Reached, then wait Hold Time till start cool down
                    //Use MaxTemp_StartTime to determine the curent duration of the Hold Peak Time
                    if (MaxTempPlateauReached)
                    {
                        TimeSpan peak_duration = DateTime.Now - MaxTemp_StartTime;

                        double hold_time_minutes = HoldTime_copy.TotalMinutes;

                        HoldTimeRemaining_copy = (hold_time_minutes - peak_duration.TotalMinutes).ToString("#0.0#");
                        HoldTimeRemainingTextBox.BeginInvoke((Action)(() => HoldTimeRemainingTextBox.Text = HoldTimeRemaining_copy));

                        if (peak_duration.TotalMinutes > hold_time_minutes)
                        {
                            oven_check = ComponentCheckStatusEnum.None;
                        }
                    }
                }

                //turn off oven heat outside of heating cycle
                SetOvenChannelOutput(OvenChanEnum.OvenOff);
                inner_coil_on = false;
                outer_coil_on = false;
                sample_zone_coil_on = false;
                
                ///////////////////////////////////////////////////////////////

                //Set interlock for PID cycle reading temperature
                IsLockedCurrentTemp_Array = true;

                //Select Max Temperature on the Outer Zone of the Sample tray
                CurrentTemp_EachZone[0] = Math.Max(CurrentTemperatures[0],
                                                   CurrentTemperatures[5]);

                //Select Max Temperature at Sample Tray in Sample Zone
                CurrentTemp_EachZone[1] = Math.Max(CurrentTemperatures[1],
                                                   CurrentTemperatures[4]);

                //Select Max Temperature at Sample Tray in Inner Zone
                CurrentTemp_EachZone[2] = Math.Max(CurrentTemperatures[2],
                                                   CurrentTemperatures[3]);

                IsLockedCurrentTemp_Array = false;

                for (int i = 0; i < 3; i++)
                {
                    //Make sure a Previous Temp of Zero is never used
                    //this will throw off the rate correction parameter
                    if (PreviousTemp_EachZone[i] == 0)
                        PreviousTemp_EachZone[i] = CurrentTemp_EachZone[i];

                    //Calculate the rate correction - D in PID
                    rate_correction_EachZone[i] =
                        (PreviousTemp_EachZone[i] - CurrentTemp_EachZone[i]) * 100 * oven_pid_settings.rate_EachZone[i];

                    //Calculate the heat duty cycle % parameter from P, I, & D
                    heat_EachZone[i] =
                        (int)((100 / oven_pid_settings.proportional_band_EachZone[i]) *
                                (set_point_temperature - CurrentTemp_EachZone[i]) +
                                oven_pid_settings.offset_EachZone[i] +
                                rate_correction_EachZone[i]);

                    //Anti-windup - prevent I from incrementing unless we're heating
                    //and at set-point
                    if (heat_EachZone[i] < 100 &&
                        heat_EachZone[i] > 0 &&
                        rate_correction_EachZone[i] == 0)
                        oven_pid_settings.offset_EachZone[i] +=
                            (set_point_temperature - CurrentTemp_EachZone[i]) * oven_pid_settings.reset_EachZone[i];

                    //Store current temp to previous temp for next cycle's calculation of the 
                    //Rate Correction (D) component.
                    PreviousTemp_EachZone[i] = CurrentTemp_EachZone[i];
                }
            }

            oven_start_check = ComponentCheckStatusEnum.None;

            //turn off oven heat outside of heating cycle
            SetOvenChannelOutput(OvenChanEnum.OvenOff);
            inner_coil_on = false;
            outer_coil_on = false;
            sample_zone_coil_on = false;

            HeatingStopped = true;
            MaxTempPlateauReached = false;
            MaxTemp_StartTime_Set = false;
            HeatingStarted = false;
            HeatingStarted_logged = false;
            MaxTempReached_logged = false;
            
            //If not thread aborting, start cooling
            if (oven_control_thread.ThreadState == ThreadState.AbortRequested ||
                oven_control_thread.ThreadState == ThreadState.Aborted)
                return;

            if (oven_check == ComponentCheckStatusEnum.Aborted) return;

            oven_cooling_check = ComponentCheckStatusEnum.Running;
            DoOvenCooling();
            
        }

        public void DoOvenCooling()
        {
            bool high_n2_on = false;
           
            //Read Switch N2 to Air and Stop Temperatures Directly from the controls
            //in case the user has changed them
            float switch_temperature = 10000;
            float stop_temperature = 10000;

            switch_temperature = SwitchTemp_copy;
            stop_temperature = StopTemp_copy;

            //Lock CurrentTemperatures array for access by the oven thread.
            IsLockedCurrentTemp_Array = true;

            //Check to see if the oven temperature is below the Stop Cooling Temperature 
            bool already_below_stop_temp = true;

            for (int i = 0; i < MAX_CHANCOUNT - NUM_WALLCOUNT; i++)
            {
                if (CurrentTemperatures[i] > stop_temperature)
                {
                    //Don't need to check any more temperatures
                    //Temperature is too high to switch
                    i = MAX_CHANCOUNT;
                    already_below_stop_temp = false;
                }
                else
                {
                    already_below_stop_temp = true;
                }
            }

            //Unlock the CurrentTemperatures Array
            IsLockedCurrentTemp_Array = false;

            CoolingStarted = true;

            //If the oven temperature is already below the stop cooling temperature
            //Set flags to end cooling
            //Else, start the Air or High Pressure N2
            if (already_below_stop_temp)
            {
                oven_cooling_check = ComponentCheckStatusEnum.None;
            }
            else
            {
                //start cooling
                if (CoolMethod_copy == "1")
                {
                    SetAirRelayChannelOutput(AirChanStatusEnum.HighN2);
                    high_n2_on = true;
                }
                else
                {
                    SetAirRelayChannelOutput(AirChanStatusEnum.Air);
                    high_n2_on = false;
                }
            }

            while (oven_cooling_check == ComponentCheckStatusEnum.Running &&
                   oven_check != ComponentCheckStatusEnum.Aborted)
            {
                //Read Switch N2 to Air and Stop Temperatures Directly from the controls
                //in case the user has changed them
                switch_temperature = SwitchTemp_copy;
                stop_temperature = StopTemp_copy;
                
                IsLockedCurrentTemp_Array = true;

                if (CoolMethod_copy == "1" &&
                    high_n2_on)
                {
                    bool do_switch = true;

                    for (int i = 0; i < MAX_CHANCOUNT - NUM_WALLCOUNT; i++)
                    {
                        if (CurrentTemperatures[i] > switch_temperature)
                        {
                            //Don't need to check any more temperatures
                            //Temperature is too high to switch
                            i = MAX_CHANCOUNT;
                            do_switch = false;
                        }
                        else
                        {
                            do_switch = true;
                        }
                    }

                    if (do_switch)
                    {
                        SetAirRelayChannelOutput(AirChanStatusEnum.Air);
                        high_n2_on = false;
                    }
                }
                else
                {
                    bool is_cooling_done = true;

                    for (int i = 0; i < MAX_CHANCOUNT - NUM_WALLCOUNT; i++)
                    {
                        if (CurrentTemperatures[i] > stop_temperature)
                        {
                            //Don't need to check any more temperatures
                            //Temperature is too high to switch
                            i = MAX_CHANCOUNT;
                            is_cooling_done = false;
                        }
                        else
                        {
                            is_cooling_done = true;
                        }
                    }

                    if (is_cooling_done)
                    {
                        //Set oven cooling check status to none to end the cooling loop
                        oven_cooling_check = ComponentCheckStatusEnum.None;
                    }
                }

                IsLockedCurrentTemp_Array = false;

                //Sleep for five seconds
                System.Threading.Thread.Sleep(5000);
            }

            SetAirRelayChannelOutput(AirChanStatusEnum.AirOff);
            OvenRunDone = true;
            CoolingStarted = false;
            CoolingStarted_logged = false;
            HeatingStopped = false;
            HeatingStopped_logged = false;

            //Stop reading the temperature
            read_check = ComponentCheckStatusEnum.None;
        }

        

        #endregion
        
        #endregion

        private void PulseSetpointTestButton_Click(object sender, EventArgs e)
        {
            String log_string = "Start Set-Point Pulse Experiment";

            //Pulse amplitude - set to half of current Max Temp Value
            double pulse_amplitude = MaxTemp_copy / 2;

            //Pulse half-width = 4 times PID cycles (3 sec cycle time)
            int pulse_length_seconds = (int)this.oven_pid_settings.cycle_time_in_seconds * 4;

            log_string =
                String.Format(
                "********************************{1}" +
                "Start Positive Pulse Time, {0}{1}" +
                "Original Set Point, {2}{1}" +
                "Pulse Amplitude, {3}{1}" +
                "Pulse 1/2 Width (sec), {4}{1}",
                DateTime.Now.ToString("yyyy/MM/dd, HH:mm:ss"),
                Environment.NewLine,
                set_point_temperature.ToString("#0.0#"),
                pulse_amplitude.ToString("#0.0#"),
                pulse_length_seconds.ToString("0"));

            set_point_temperature = set_point_temperature + pulse_amplitude;

            TimeSpan pulse_half_width = new TimeSpan(0, 0, pulse_length_seconds);

            //Wait first half - positive half - of pulse
            System.Threading.Thread.Sleep((int)pulse_half_width.TotalMilliseconds);

            log_string =
                String.Format(
                "Start Negative Pulse Time, {0}{1}" +
                "Original Set Point, {2}{1}" +
                "Pulse Amplitude, -{3}{1}" +
                "Pulse 1/2 Width (sec), {4}{1}",
                DateTime.Now.ToString("yyyy/MM/dd, HH:mm:ss"),
                Environment.NewLine,
                set_point_temperature.ToString("#0.0#"),
                pulse_amplitude.ToString("#0.0#"),
                pulse_length_seconds.ToString("0"));

            //Change setup for negative half
            set_point_temperature = set_point_temperature - (pulse_amplitude * 2);

            //Wait second half - negative half - of pulse
            System.Threading.Thread.Sleep((int)pulse_half_width.TotalMilliseconds);

            //return set_point to original value;
            set_point_temperature = set_point_temperature + pulse_amplitude;

            //Pulse experiment is done
            log_string =
                String.Format(
                "End Pulse Time,{0}{1}" +
                "Original Set Point,{2}{1}" +
                "Peak-To-Peak Pulse Amplitude,{3}{1}" +
                "Pulse Total Width (sec),{4}{1}" +
                "**********************************{1}{1}",
                DateTime.Now.ToString("yyyy/MM/dd, HH:mm:ss"),
                Environment.NewLine,
                set_point_temperature.ToString("#0.0#"),
                (pulse_amplitude * 2).ToString("#0.0#"),
                (pulse_length_seconds * 2).ToString("0"));

            string file_path = SaveLogFileDirectoryTextBox.Text;
            file_path = Path.Combine(file_path, "pulse_setpoint_log_file.csv");

            if (!File.Exists(file_path))
            {
                File.Create(file_path);
            }

            using (FileStream fs = new FileStream(file_path,
                                                  FileMode.Append,
                                                  FileAccess.Write,
                                                  FileShare.ReadWrite))
            {
                byte[] array = Encoding.ASCII.GetBytes(log_string);
                int array_len = array.GetLength(0);

                fs.Write(array, 0, array_len);
                fs.Close();
            }
        }

        private void StopCoolingTemperatureTextBox_TextChanged(object sender, EventArgs e)
        {
            StopTemp_copy = (float)StopCoolingTemperatureNumericUpDown.Value;
        }

        private void SwitchTempTextBox_TextChanged(object sender, EventArgs e)
        {
            SwitchTemp_copy = (float)SwitchToAirTemperatureNumericUpDown.Value;
        }

        private void HoldTimeTextBox_TextChanged(object sender, EventArgs e)
        {
            HoldTime_copy = new TimeSpan(0, (int)HoldTimeAtPeakNumericUpDown.Value, 0);
        }

        private void MaxTempTextBox_TextChanged(object sender, EventArgs e)
        {
            MaxTemp_copy = (float)MaxTemperatureNumericUpDown.Value;
            set_point_temperature = MaxTemp_copy;
        }
    }
}
