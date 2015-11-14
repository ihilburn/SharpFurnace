using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

namespace USB_TC_Control
{
    public delegate void OnUpdateLoadTextHandler(string new_load_text);

    public partial class SplashScreen : Form
    {
        public SplashScreen()
        {
            InitializeComponent();
        }

        private void SplashScreen_Load(object sender, EventArgs e)
        {
            VersionLabel.Text =
                String.Format(
                    "v{0}.{1}.{2} r{3}",
                    Assembly.GetExecutingAssembly().GetName().Version.Major.ToString("0"),
                    Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString("0"),
                    Assembly.GetExecutingAssembly().GetName().Version.Build.ToString("0"),
                    Assembly.GetExecutingAssembly().GetName().Version.Revision.ToString("0"));
        }

        public void WaitForHideSplashScreen()
        {
            while (!Program.hide_splash_screen)
            {
                //Do nothing
            }

            this.Hide();
        }
    }
}
