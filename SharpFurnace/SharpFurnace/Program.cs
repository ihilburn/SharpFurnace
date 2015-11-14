using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace USB_TC_Control
{
    static class Program
    {
        public static USB_TC_Control_Form temperature_control_form;
        public static PID_Settings_Form pid_settings_form;
        public static SplashScreen splash_screen_form;

        public static Boolean hide_splash_screen;
        public static System.Threading.Thread splash_screen_thread;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);


            splash_screen_thread = new System.Threading.Thread(ShowSplashScreen);
            splash_screen_thread.IsBackground = true;
            splash_screen_thread.Start();
           
            pid_settings_form = new PID_Settings_Form();
            temperature_control_form = new USB_TC_Control_Form();
                        
            temperature_control_form.WindowState = FormWindowState.Maximized;
            temperature_control_form.FormClosed += new FormClosedEventHandler(temp_control_form_OnClosed);

            Application.Run(temperature_control_form);

        }

        private static void ShowSplashScreen()
        {
            hide_splash_screen = false;

            splash_screen_form = new SplashScreen();
            Application.Run(splash_screen_form);
        }

        public static void temp_control_form_OnClosed(Object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
