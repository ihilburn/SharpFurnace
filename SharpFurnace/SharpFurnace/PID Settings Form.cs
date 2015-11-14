using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace USB_TC_Control
{
    public partial class PID_Settings_Form : Form
    {
        public PIDSettings ProgramPidSettings { get; set; }
        public String XmlSettingsFilePath { get; set; }

        private PIDSettings initial_defaults;
        private PIDSettings empty_air_defaults;
        private PIDSettings empty_n2_defaults;
        private PIDSettings quarter_full_air_defaults;
        private PIDSettings quarter_full_n2_defaults;
        private PIDSettings half_full_air_defaults;
        private PIDSettings half_full_n2_defaults;
        private PIDSettings threequarters_full_air_defaults;
        private PIDSettings threequarters_full_n2_defaults;
        private PIDSettings full_air_defaults;
        private PIDSettings full_n2_defaults;

        public PID_Settings_Form()
        {
            InitializeComponent();

            if (!File.Exists(XmlSettingsFilePath))
            {
                XmlSettingsFilePath = @"C:\Users\lab\Dropbox\Ge SURF\Stuff\Settings\DefaultPIDSettingsXMLFile.xml";
            }

            InstantiateDefaultPIDSettingsInstances();
            LoadDefaultPIDSettings_FromXMLFile();

            SetPIDSettings_ToControls(initial_defaults);
            this.ProgramPidSettings = initial_defaults;

            
        }

        private void InstantiateDefaultPIDSettingsInstances()
        {
            initial_defaults = new PIDSettings();
            empty_air_defaults = new PIDSettings();
            empty_n2_defaults = new PIDSettings();
            quarter_full_air_defaults = new PIDSettings();
            quarter_full_n2_defaults = new PIDSettings();
            half_full_air_defaults = new PIDSettings();
            half_full_n2_defaults = new PIDSettings();
            threequarters_full_air_defaults = new PIDSettings();
            threequarters_full_n2_defaults = new PIDSettings();
            full_air_defaults = new PIDSettings();
            full_n2_defaults = new PIDSettings();
        }
        
        private PIDSettings GetPIDSettings_FromControls()
        {
            PIDSettings pid_settings = new PIDSettings();

            pid_settings.cycle_time_in_seconds = (double)this.CycleTimeInSecondsNumericUpDown.Value;

            pid_settings.proportional_band_EachZone[0] = (double)this.OuterProportionalBandNumericUpDown.Value;
            pid_settings.proportional_band_EachZone[1] = (double)this.SampleZoneProportionalBandNumericUpDown.Value;
            pid_settings.proportional_band_EachZone[2] = (double)this.InnerProportionalBandNumericUpDown.Value;

            pid_settings.offset_EachZone[0] = (double)this.OuterOffsetNumericUpDown.Value;
            pid_settings.offset_EachZone[1] = (double)this.SampleZoneOffsetNumericUpDown.Value;
            pid_settings.offset_EachZone[2] = (double)this.InnerOffsetNumericUpDown.Value;

            pid_settings.rate_EachZone[0] = (double)this.OuterRateNumericUpDown.Value;
            pid_settings.rate_EachZone[1] = (double)this.SampleZoneRateNumericUpDown.Value;
            pid_settings.rate_EachZone[2] = (double)this.InnerRateNumericUpDown.Value;

            pid_settings.reset_EachZone[0] = (double)this.OuterResetNumericUpDown.Value;
            pid_settings.reset_EachZone[1] = (double)this.SampleZoneResetNumericUpDown.Value;
            pid_settings.reset_EachZone[2] = (double)this.InnerResetNumericUpDown.Value;

            return pid_settings;
        }

        private void SetPIDSettings_ToControls(PIDSettings pid_settings)
        {
            this.CycleTimeInSecondsNumericUpDown.Value = (decimal)pid_settings.cycle_time_in_seconds;

            this.OuterProportionalBandNumericUpDown.Value = (decimal)pid_settings.proportional_band_EachZone[0];
            this.SampleZoneProportionalBandNumericUpDown.Value = (decimal)pid_settings.proportional_band_EachZone[1];
            this.InnerProportionalBandNumericUpDown.Value = (decimal)pid_settings.proportional_band_EachZone[2];

            this.OuterOffsetNumericUpDown.Value = (decimal)pid_settings.offset_EachZone[0];
            this.SampleZoneOffsetNumericUpDown.Value = (decimal)pid_settings.offset_EachZone[1];
            this.InnerOffsetNumericUpDown.Value = (decimal)pid_settings.offset_EachZone[2];

            this.OuterRateNumericUpDown.Value = (decimal)pid_settings.rate_EachZone[0];
            this.SampleZoneRateNumericUpDown.Value = (decimal)pid_settings.rate_EachZone[1];
            this.InnerRateNumericUpDown.Value = (decimal)pid_settings.rate_EachZone[2];

            this.OuterResetNumericUpDown.Value = (decimal)pid_settings.reset_EachZone[0];
            this.SampleZoneResetNumericUpDown.Value = (decimal)pid_settings.reset_EachZone[1];
            this.InnerResetNumericUpDown.Value = (decimal)pid_settings.reset_EachZone[2];
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void DefaultPIDSettingsSelectorComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (DefaultPIDSettingsSelectorComboBox.SelectedIndex)
            { 
                case 0:
                    SetPIDSettings_ToControls(empty_air_defaults);
                    break;

                case 1:
                    SetPIDSettings_ToControls(empty_n2_defaults);
                    break;

                case 2:
                    SetPIDSettings_ToControls(quarter_full_air_defaults);
                    break;

                case 3:
                    SetPIDSettings_ToControls(quarter_full_n2_defaults);
                    break;

                case 4:
                    SetPIDSettings_ToControls(half_full_air_defaults);
                    break;

                case 5:
                    SetPIDSettings_ToControls(half_full_n2_defaults);
                    break;

                case 6:
                    SetPIDSettings_ToControls(threequarters_full_air_defaults);
                    break;

                case 7:
                    SetPIDSettings_ToControls(threequarters_full_n2_defaults);
                    break;

                case 8:
                    SetPIDSettings_ToControls(full_air_defaults);
                    break;

                case 9:
                    SetPIDSettings_ToControls(full_n2_defaults);
                    break;

                default:
                    SetPIDSettings_ToControls(initial_defaults);
                    break;

            }
        }

        private void LoadDefaultPIDSettings_FromXMLFile()
        {
            try
            {
                using (FileStream fs = new FileStream(this.XmlSettingsFilePath,
                                                      FileMode.Open,
                                                      FileAccess.Read,
                                                      FileShare.Read))
                {
                    XmlTextReader xml_reader = new XmlTextReader(fs);

                    xml_reader.ReadStartElement("PIDDefaults");

                    xml_reader.ReadStartElement("InitialDefaults");
                    initial_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("EmptyAir");
                    empty_air_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("EmptyN2");
                    empty_n2_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("QuarterFullAir");
                    quarter_full_air_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("QuarterFullN2");
                    quarter_full_n2_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("HalfFullAir");
                    half_full_air_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("HalfFullN2");
                    half_full_n2_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("ThreeQuartersFullAir");
                    threequarters_full_air_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("ThreeQuartersFullN2");
                    threequarters_full_n2_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("FullAir");
                    full_air_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.ReadStartElement("FullN2");
                    full_n2_defaults.LoadPIDSetting_FromXmlTextReader(ref xml_reader);
                    xml_reader.ReadEndElement();

                    xml_reader.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(
                    String.Format(
                        "Unexpected Error Loading Default PID Settings from XML Settings File.{0}" +
                        "Error Message: {1}{0}" +
                        "Error Source: {2}{0}" +
                        "Stack Trace: {3}{0}",
                        Environment.NewLine,
                        e.Message,
                        e.Source,
                        e.StackTrace), 
                        "      ERROR!!");
            }
        }

        private void SaveDefaultPIDSettings_ToXMLFile()
        {
            try
            {

                XmlTextWriter xml_writer = new XmlTextWriter(this.XmlSettingsFilePath,Encoding.UTF8);

                xml_writer.WriteStartElement("PIDDefaults");
                
                xml_writer.WriteStartElement("InitialDefaults");
                xml_writer.WriteRaw(initial_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("EmptyAir");
                xml_writer.WriteRaw(this.empty_air_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("EmptyN2");
                xml_writer.WriteRaw(this.empty_n2_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("QuarterFullAir");
                xml_writer.WriteRaw(this.quarter_full_air_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("QuarterFullN2");
                xml_writer.WriteRaw(this.quarter_full_n2_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("HalfFullAir");
                xml_writer.WriteRaw(this.half_full_air_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("HalfFullN2");
                xml_writer.WriteRaw(this.half_full_n2_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("ThreeQuartersFullAir");
                xml_writer.WriteRaw(this.threequarters_full_air_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("ThreeQuartersFullN2");
                xml_writer.WriteRaw(this.threequarters_full_n2_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("FullAir");
                xml_writer.WriteRaw(this.full_air_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteStartElement("FullN2");
                xml_writer.WriteRaw(this.full_n2_defaults.PIDSetting_ToXMLString());
                xml_writer.WriteEndElement();

                xml_writer.WriteEndElement();
                xml_writer.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(
                    String.Format(
                        "Unexpected Error Saving Default PID Settings from XML Settings File.{0}" +
                        "Error Message: {1}{0}" +
                        "Error Source: {2}{0}" +
                        "Stack Trace: {3}{0}",
                        Environment.NewLine,
                        e.Message,
                        e.Source,
                        e.StackTrace),
                        "      ERROR!!");
            }
        }

        private void SetPIDDefaultFromDisplayButton_Click(object sender, EventArgs e)
        {
            switch (DefaultPIDSettingsSelectorComboBox.SelectedIndex)
            {
                case 0:
                    empty_air_defaults = GetPIDSettings_FromControls();
                    break;

                case 1:
                    empty_n2_defaults = GetPIDSettings_FromControls();
                    break;

                case 2:
                    quarter_full_air_defaults = GetPIDSettings_FromControls();
                    break;

                case 3:
                    quarter_full_n2_defaults = GetPIDSettings_FromControls();
                    break;

                case 4:
                    half_full_air_defaults = GetPIDSettings_FromControls();
                    break;

                case 5:
                    half_full_n2_defaults = GetPIDSettings_FromControls();
                    break;

                case 6:
                    threequarters_full_air_defaults = GetPIDSettings_FromControls();
                    break;

                case 7:
                    threequarters_full_n2_defaults = GetPIDSettings_FromControls();
                    break;

                case 8:
                    full_air_defaults = GetPIDSettings_FromControls();
                    break;

                case 9:
                    full_n2_defaults = GetPIDSettings_FromControls();
                    break;

                default:
                    initial_defaults = GetPIDSettings_FromControls();
                    break;

            }
        }

        private void SetPIDValuesButton_Click(object sender, EventArgs e)
        {
            this.ProgramPidSettings = GetPIDSettings_FromControls();
            this.ProgramPidSettings.IsChanged = true;
            Program.temperature_control_form.oven_pid_settings = this.ProgramPidSettings;

            this.Hide();
        }

        private void SaveDefaultPIDSettingsToXmlFileButton_Click(object sender, EventArgs e)
        {
            SaveDefaultPIDSettings_ToXMLFile();
        }


    }

    public class PIDSettings
    {
        public Boolean IsChanged { get; set; }

        //Proportional Band - P in PID
        public double[] proportional_band_EachZone { get; set; }

        //Offset per each zone
        public double[] offset_EachZone { get; set; }

        //reset per each zone
        public double[] reset_EachZone { get; set; }

        //rate per each zone
        public double[] rate_EachZone { get; set; }

        //Cycle time - in seconds
        public double cycle_time_in_seconds { get; set; }

        public PIDSettings()
        {
            this.IsChanged = false;
            this.proportional_band_EachZone = new Double[3] { 2, 3, 30 };
            this.offset_EachZone = new Double[3] { 5, 5, 5 };
            this.reset_EachZone = new Double[3] { 0.1, 0.1, 0.1 };
            this.rate_EachZone = new Double[3] { 0.25, 0.25, 0.25 };
            this.cycle_time_in_seconds = 2;
        }

        public String PIDSetting_ToXMLString()
        {
            String xml_string = String.Empty;

            xml_string +=
                String.Format(
                    "<cycle_time_in_seconds>{1}</cycle_time_in_seconds>{0}",
                    Environment.NewLine,
                    this.cycle_time_in_seconds.ToString("#0.0#"));

            xml_string +=
                String.Format(
                    "<outer_zone>{0}" +
                    "{1}<proportional_band>{2}</proportional_band>{0}" +
                    "{1}<offset>{3}</offset>{0}" +
                    "{1}<reset>{4}</reset>{0}" +
                    "{1}<rate>{5}</rate>{0}" +
                    "</outer_zone>{0}",
                    Environment.NewLine,
                    "\t",
                    this.proportional_band_EachZone[0].ToString("#0.0#"),
                    this.offset_EachZone[0].ToString("#0.0#"),
                    this.reset_EachZone[0].ToString("#0.0#"),
                    this.rate_EachZone[0].ToString("#0.0#"));

            xml_string +=
                String.Format(
                    "<sample_zone>{0}" +
                    "{1}<proportional_band>{2}</proportional_band>{0}" +
                    "{1}<offset>{3}</offset>{0}" +
                    "{1}<reset>{4}</reset>{0}" +
                    "{1}<rate>{5}</rate>{0}" +
                    "</sample_zone>{0}",
                    Environment.NewLine,
                    "\t",
                    this.proportional_band_EachZone[1].ToString("#0.0#"),
                    this.offset_EachZone[1].ToString("#0.0#"),
                    this.reset_EachZone[1].ToString("#0.0#"),
                    this.rate_EachZone[1].ToString("#0.0#"));

            xml_string +=
                String.Format(
                    "<inner_zone>{0}" +
                    "{1}<proportional_band>{2}</proportional_band>{0}" +
                    "{1}<offset>{3}</offset>{0}" +
                    "{1}<reset>{4}</reset>{0}" +
                    "{1}<rate>{5}</rate>{0}" +
                    "</inner_zone>{0}",
                    Environment.NewLine,
                    "\t",
                    this.proportional_band_EachZone[2].ToString("#0.0#"),
                    this.offset_EachZone[2].ToString("#0.0#"),
                    this.reset_EachZone[2].ToString("#0.0#"),
                    this.rate_EachZone[2].ToString("#0.0#"));

            return xml_string;
        }

        public void LoadPIDSetting_FromXmlTextReader(ref XmlTextReader xml_reader)
        {
            this.cycle_time_in_seconds = Double.Parse(xml_reader.ReadElementString("cycle_time_in_seconds"));

            xml_reader.ReadStartElement("outer_zone");
            this.proportional_band_EachZone[0] = Double.Parse(xml_reader.ReadElementString("proportional_band"));
            this.offset_EachZone[0] = Double.Parse(xml_reader.ReadElementString("offset"));
            this.reset_EachZone[0] = Double.Parse(xml_reader.ReadElementString("reset"));
            this.rate_EachZone[0] = Double.Parse(xml_reader.ReadElementString("rate"));
            xml_reader.ReadEndElement();

            xml_reader.ReadStartElement("sample_zone");
            this.proportional_band_EachZone[1] = Double.Parse(xml_reader.ReadElementString("proportional_band"));
            this.offset_EachZone[1] = Double.Parse(xml_reader.ReadElementString("offset"));
            this.reset_EachZone[1] = Double.Parse(xml_reader.ReadElementString("reset"));
            this.rate_EachZone[1] = Double.Parse(xml_reader.ReadElementString("rate"));
            xml_reader.ReadEndElement();

            xml_reader.ReadStartElement("inner_zone");
            this.proportional_band_EachZone[2] = Double.Parse(xml_reader.ReadElementString("proportional_band"));
            this.offset_EachZone[2] = Double.Parse(xml_reader.ReadElementString("offset"));
            this.reset_EachZone[2] = Double.Parse(xml_reader.ReadElementString("reset"));
            this.rate_EachZone[2] = Double.Parse(xml_reader.ReadElementString("rate"));
            xml_reader.ReadEndElement();
        }
    }
}
