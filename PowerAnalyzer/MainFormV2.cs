using ClosedXML.Excel;
using Guna.UI2.WinForms;
using PowerAnalyzer.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerAnalyzer
{
    public partial class MainFormV2 : Form
    {
        private SerialPort _serialPort;
        private List<byte> _buffer = new List<byte>();

        private readonly byte[] startMarker = { 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
        private readonly byte[] endMarker = { 0xFE, 0xFE, 0xFE, 0xFE, 0xFE };

        string v1, v2, v3, systemVolt;
        string amp1, amp2, amp3, systemAmp;
        string watt1, watt2, watt3, systemWatt;
        string pf1, pf2, pf3, systemPF;
        string va1, va2, va3, systemVA;
        string frequency, temperature;
        string meanVolt1, meanVolt2, meanVolt3, systemMeanVolt;
        string peakAmp1, peakAmp2, peakAmp3, systemPeakAmp;
        string peakVolt1, peakVolt2, peakVolt3, systemPeakVolt;

        string meanAmp1, meanAmp2, meanAmp3, systemMeanAmp;
        string FfPh1V, FfPh2V, FfPh3V, FfPhSysV;

        

        string CfPh1V, CfPh2V, CfPh3V, CfPhSysV;
        string FfPh1A, FfPh2A, FfPh3A, FfPhSysA;
        string CfPh1A, CfPh2A, CfPh3A, CfPhSysA;

        string pf1Dp="0",pf2Dp = "0", pf3Dp = "0", sysPfDp = "0";
        string frequencyDp = "0";
        string temperatureDp = "0";
        string v1Dp = "0", v2Dp = "0", v3Dp = "0", sysVDp = "0";        
        string i1Dp = "0", i2Dp = "0", i3Dp = "0", sysIDp = "0";
        string p1Dp = "0", p2Dp = "0", p3Dp = "0", sysPDp = "0";
        string ffPh1Vdp = "0", ffPh2Vdp = "0", ffPh3Vdp = "0", ffPhSysVdp = "0";
        string ffPh1Adp = "0", ffPh2Adp = "0", ffPh3Adp = "0", ffPhSysAdp = "0";
        string cfPh1Vdp = "0", cfPh2Vdp = "0", cfPh3Vdp = "0", cfPhSysVdp = "0";
        string cfPh1Adp = "0", cfPh2Adp = "0", cfPh3Adp = "0", cfPhSysAdp = "0";

        #region Lifecycle

        public MainFormV2()
        {
            InitializeComponent();


            _serialPort = new SerialPort
            {
                BaudRate = 9600, // Change according to your device
                DataBits = 8,
                Parity = Parity.None,
                StopBits = StopBits.One
            };
            _serialPort.DataReceived += SerialPort_DataReceived;

            LoadAvailablePorts();
            LoadPreference();

            foreach(DataGridViewColumn column in dataGrid.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            }
        }

        
        private void LoadAvailablePorts()
        {
            cmbPort.Items.Clear();
            string[] ports = SerialPort.GetPortNames();

            if (ports.Length > 0)
            {
                cmbPort.Items.AddRange(ports);
                cmbPort.SelectedIndex = 0; // Select first port
            }
            else
            {
                cmbPort.Items.Add("No Ports Found");
                cmbPort.SelectedIndex = 0;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_serialPort != null && _serialPort.IsOpen)
                _serialPort.Close();

            base.OnFormClosing(e);
        }

        #endregion

        #region Connect Disconnect

        private void toggleSwitchConnectDisconnect_CheckedChanged(object sender, EventArgs e)
        {
            if (toggleSwitchConnectDisconnect.Checked)
            {
                Connect(); 
            }
            else
            {
                Disconnect();
            }
        }

        private void Connect()
        {
            try
            {
                //if (cmbPort.SelectedItem == null || cmbPort.SelectedItem.ToString() != "COM3")
                //{
                //    toggleSwitchConnectDisconnect.Checked = false;
                //    MessageBox.Show("No valid COM port selected!");
                //    return;
                //}

                if (_serialPort.IsOpen) _serialPort.Close();

                _serialPort.PortName = cmbPort.SelectedItem.ToString();
                _serialPort.Open();

                btnLog.Enabled = true;
                btnClearAll.Enabled = true;
                btnExportExcel.Enabled = true;

                btnFast.Enabled = true;
                btnAccurate.Enabled = true;
                btnHold.Enabled = true;

                // Console.Write("Connected to " + _serialPort.PortName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error connecting: " + ex.Message);
            }
        }

        private void Disconnect()
        {
            try
            {
                if (_serialPort.IsOpen)
                {
                    _serialPort.Close();
                    // Console.WriteLine("Disconnected");                    
                }

                btnLog.Enabled = false;
                btnClearAll.Enabled = false;
                btnExportExcel.Enabled = false;

                btnFast.Enabled = false;
                btnAccurate.Enabled = false;
                btnHold.Enabled = false;

                ClearAllData();                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error disconnecting: " + ex.Message);
            }
        }

        private void ClearAllData()
        {
            lblV1.Text = "0.0";
            lblV2.Text = "0.0";
            lblV3.Text = "0.0";
            lblSystemVolt.Text = "0.0";

            lblAmp1.Text = "0.0";
            lblAmp2.Text = "0.0";
            lblAmp3.Text = "0.0";
            lblSystemAmp.Text = "0.0";

            lblWatt1.Text = "0.0";
            lblWatt2.Text = "0.0";
            lblWatt3.Text = "0.0";
            lblSystemWatt.Text = "0.0";

            lblPF1.Text = "0.0";
            lblPF2.Text = "0.0";
            lblPF3.Text = "0.0";
            lblSystemPF.Text = "0.0";

            lblVA1.Text = "0.0";
            lblVA2.Text = "0.0";
            lblVA3.Text = "0.0";
            lblSystemVA.Text = "0.0";

            lblFrequency.Text = "0.0";
            lblTemperature.Text = "0.0";

            lblMeanVolt1.Text = "0.0";
            lblMeanVolt2.Text = "0.0";
            lblMeanVolt3.Text = "0.0";
            lblSystemMeanV.Text = "0.0";

            lblPeakAmp1.Text = "0.0";
            lblPeakAmp2.Text = "0.0";
            lblPeakAmp3.Text = "0.0";
            lblSystemPeakA.Text = "0.0";

            lblPeakVolt1.Text = "0.0";
            lblPeakVolt2.Text = "0.0";
            lblPeakVolt3.Text = "0.0";
            lblSystemPeakV.Text = "0.0";

            lblMeanAmp1.Text = "0.0";
            lblMeanAmp2.Text = "0.0";
            lblMeanAmp3.Text = "0.0";
            lblSystemMeanA.Text = "0.0";

            lblFFPH1V.Text = "0.0";
            lblFFPH2V.Text = "0.0";
            lblFFPH3V.Text = "0.0";
            lblFFPHSYSV.Text = "0.0";

            lblCFPH1V.Text = "0.0";
            lblCFPH2V.Text = "0.0";
            lblCFPH3V.Text = "0.0";
            lblCFPHSYSV.Text = "0.0";


            lblFFPH1A.Text = "0.0";
            lblFFPH2A.Text = "0.0";
            lblFFPH3A.Text = "0.0";
            lblFFPHSYSA.Text = "0.0";

            lblCFPH1A.Text = "0.0";
            lblCFPH2A.Text = "0.0";
            lblCFPH3A.Text = "0.0";
            lblCFPHSYSA.Text = "0.0";
        }

        #endregion

        #region Data Reading

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                int bytesToRead = _serialPort.BytesToRead;
                byte[] tempBuffer = new byte[bytesToRead];
                _serialPort.Read(tempBuffer, 0, bytesToRead);

                lock (_buffer) // thread-safe
                {
                    _buffer.AddRange(tempBuffer);

                    while (true)
                    {

                        int startIndex = Helper.FindSequence(_buffer, startMarker);
                        if (startIndex == -1) break;

                        int endIndex = Helper.FindSequence(_buffer, endMarker, startIndex + startMarker.Length);
                        if (endIndex == -1) break;

                        int length = endIndex + endMarker.Length - startIndex;
                        byte[] frame = _buffer.Skip(startIndex).Take(length).ToArray();

                        // Remove processed bytes from buffer
                        _buffer.RemoveRange(0, startIndex + length);

                        // Convert to HEX string
                        string hexFrame = BitConverter.ToString(frame).Replace("-", " ");
                        // byte[] byteArray = Helper.HexStringToByteArray(hexFrame);

                        this.BeginInvoke(new Action(() =>
                        {
                            // Console.WriteLine(hexString);
                            string[] hexStringArray = hexFrame.Split(' ');

                            Console.Write(hexStringArray[225]+" -- "+hexStringArray[226]+" ====== ");
                            Console.WriteLine(Helper.HexStringToDecimal(hexStringArray[225], "0")+" ===== "+ Helper.HexStringToDecimal(hexStringArray[226], "0")+"\n");

                            #region Dynamic Data Reading

                            try
                            {
                                // Loop - Dp Value Setup
                                for (int i = 0; i < hexStringArray.Length; i = i + 5)
                                {
                                    string checkByte = Helper.HexStringToDecimal(hexStringArray[i], "0");

                                    switch (checkByte)
                                    {
                                        case "15":
                                            pf1Dp = hexStringArray[i + 1];
                                            break;
                                        case "16":
                                            pf2Dp = hexStringArray[i + 1];
                                            break;
                                        case "17":
                                            pf3Dp = hexStringArray[i + 1];
                                            break;
                                        case "18":
                                            frequencyDp = hexStringArray[i + 1];
                                            break;
                                        case "22":
                                            sysPfDp = hexStringArray[i + 1];
                                            break;
                                        case "40":
                                            temperatureDp = hexStringArray[i + 1];
                                            break;
                                        case "49":
                                            v1Dp = hexStringArray[i + 4];
                                            v2Dp = hexStringArray[i + 3];
                                            v3Dp = hexStringArray[i + 2];
                                            sysVDp = hexStringArray[i + 2];
                                            break;
                                        case "50":
                                            i1Dp = hexStringArray[i + 4];
                                            i2Dp = hexStringArray[i + 3];
                                            i3Dp = hexStringArray[i + 2];
                                            sysIDp = hexStringArray[i + 1];                                            
                                            break;
                                        case "51":
                                            p1Dp = hexStringArray[i + 4];
                                            p2Dp = hexStringArray[i + 3];
                                            p3Dp = hexStringArray[i + 2];
                                            sysPDp = hexStringArray[i + 1];
                                            break;
                                        case "58":
                                            ffPh1Vdp = hexStringArray[i + 1];
                                            break;
                                        case "59":
                                            ffPh2Vdp = hexStringArray[i + 1];
                                            break;
                                        case "60":
                                            ffPh3Vdp = hexStringArray[i + 1];
                                            break;
                                        case "61":
                                            ffPhSysVdp = hexStringArray[i + 1];
                                            break;
                                        case "62":
                                            cfPh1Vdp = hexStringArray[i + 1];
                                            break;
                                        case "63":
                                            cfPh2Vdp = hexStringArray[i + 1];
                                            break;
                                        case "64":
                                            cfPh3Vdp = hexStringArray[i + 1];
                                            break;
                                        case "65":
                                            cfPhSysVdp = hexStringArray[i + 1];
                                            break;
                                        case "66":
                                            ffPh1Adp = hexStringArray[i + 1];
                                            break;
                                        case "67":
                                            ffPh2Adp = hexStringArray[i + 1];
                                            break;
                                        case "68":
                                            ffPh3Adp = hexStringArray[i + 1];
                                            break;
                                        case "69":
                                            ffPhSysAdp = hexStringArray[i + 1];
                                            break;
                                        case "70":
                                            cfPh1Adp = hexStringArray[i + 1];
                                            break;
                                        case "71":
                                            cfPh2Adp= hexStringArray[i + 1];
                                            break;
                                        case "72":
                                            cfPh3Adp= hexStringArray[i + 1];
                                            break;
                                        case "73":
                                            cfPhSysAdp= hexStringArray[i + 1];
                                            break;

                                        default:
                                            break;
                                    }                                    
                                }
                                

                                // Loop - Data Value Setup
                                for (int i = 0; i < hexStringArray.Length; i = i + 5)
                                {                                    
                                    string checkByte = Helper.HexStringToDecimal(hexStringArray[i], "0");
                                    // Console.Write(checkByte+" --- ");

                                    string next4ByteDataString = hexStringArray[i+1] + hexStringArray[i+2] + hexStringArray[i+3] + hexStringArray[i+4];
                                    string next3ByteAfterDPDataString = hexStringArray[i + 2] + hexStringArray[i + 3] + hexStringArray[i + 4];

                                    switch (checkByte)
                                    {
                                        case "6":
                                            v1 = Helper.HexStringToDecimal(next4ByteDataString, v1Dp);
                                            lblV1.Text = v1;
                                            break;
                                        case "7":
                                            v2 = Helper.HexStringToDecimal(next4ByteDataString, v2Dp);
                                            lblV2.Text = v2;
                                            break;
                                        case "8":
                                            v3 = Helper.HexStringToDecimal(next4ByteDataString, v3Dp);
                                            lblV2.Text = v3;
                                            break;
                                        case "9":
                                            amp1 = Helper.HexStringToDecimal(next4ByteDataString, i1Dp);
                                            lblAmp1.Text = amp1;
                                            break;
                                        case "10":
                                            amp2 = Helper.HexStringToDecimal(next4ByteDataString, i2Dp);
                                            lblAmp2.Text = amp2;
                                            break;
                                        case "11":
                                            amp3 = Helper.HexStringToDecimal(next4ByteDataString, i3Dp);
                                            lblAmp3.Text = amp3;
                                            break;
                                        case "12":
                                            watt1 = Helper.HexStringToDecimal(next4ByteDataString, p1Dp);
                                            lblWatt1.Text = watt1;
                                            break;
                                        case "13":
                                            watt2 = Helper.HexStringToDecimal(next4ByteDataString, p2Dp);
                                            lblWatt2.Text = watt2;
                                            break;
                                        case "14":
                                            watt3 = Helper.HexStringToDecimal(next4ByteDataString, p3Dp);
                                            lblWatt3.Text = watt3;
                                            break;
                                        case "15":
                                            pf1 = Helper.HexStringToDecimal(next3ByteAfterDPDataString, pf1Dp);
                                            lblPF1.Text = pf1;
                                            break;
                                        case "16":
                                            pf2 = Helper.HexStringToDecimal(next3ByteAfterDPDataString, pf2Dp);
                                            lblPF2.Text = pf2;
                                            break;
                                        case "17":
                                            pf3 = Helper.HexStringToDecimal(next3ByteAfterDPDataString, pf3Dp);
                                            lblPF3.Text = pf3;
                                            break;
                                        case "18":
                                            frequency = Helper.HexStringToDecimal(next3ByteAfterDPDataString, frequencyDp);
                                            lblPF3.Text = frequency;
                                            break;
                                        case "19":
                                            systemVolt = Helper.HexStringToDecimal(next4ByteDataString, sysVDp);
                                            lblSystemVolt.Text = systemVolt;
                                            break;
                                        case "20":
                                            systemAmp = Helper.HexStringToDecimal(next4ByteDataString, sysIDp);
                                            lblSystemAmp.Text = systemAmp;
                                            break;
                                        case "21":
                                            systemWatt = Helper.HexStringToDecimal(next4ByteDataString, sysPDp);
                                            lblSystemWatt.Text = systemWatt;
                                            break;
                                        case "22":
                                            systemPF = Helper.HexStringToDecimal(next3ByteAfterDPDataString, sysPfDp);
                                            lblSystemPF.Text = systemPF;
                                            break;
                                        case "23":                                            
                                            systemVA = Helper.HexStringToDecimal(next4ByteDataString, sysVDp);
                                            // lblSystemVA.Text = hexStringArray[i+1]+ "--"+ hexStringArray[i + 2] + "--"+ hexStringArray[i + 3] + "--" + hexStringArray[i + 4] + "-" + sysVDp;
                                            lblSystemVA.Text = systemVA;
                                            break;
                                        case "24":

                                            break;
                                        case "25":
                                            peakVolt1 = Helper.HexStringToDecimal(next4ByteDataString, v1Dp);
                                            lblPeakVolt1.Text = peakVolt1;
                                            break;
                                        case "26":
                                            peakVolt2 = Helper.HexStringToDecimal(next4ByteDataString, v2Dp);
                                            lblPeakVolt2.Text = peakVolt2;
                                            break;
                                        case "27":
                                            peakVolt3 = Helper.HexStringToDecimal(next4ByteDataString, v3Dp);
                                            lblPeakVolt3.Text = peakVolt3;
                                            break;
                                        case "28":
                                            meanVolt1 = Helper.HexStringToDecimal(next4ByteDataString, v1Dp);
                                            lblMeanVolt1.Text = meanVolt1;
                                            break;
                                        case "29":
                                            meanVolt2 = Helper.HexStringToDecimal(next4ByteDataString, v2Dp);
                                            lblMeanVolt2.Text = meanVolt2;
                                            break;
                                        case "30":
                                            meanVolt3 = Helper.HexStringToDecimal(next4ByteDataString, v3Dp);
                                            lblMeanVolt3.Text = meanVolt3;
                                            break;
                                        case "31":
                                            va1 = Helper.HexStringToDecimal(next4ByteDataString, v1Dp);
                                            lblVA1.Text = va1;
                                            break;
                                        case "32":
                                            va2 = Helper.HexStringToDecimal(next4ByteDataString, v2Dp);
                                            lblVA2.Text = va2;
                                            break;
                                        case "33":
                                            va3 = Helper.HexStringToDecimal(next4ByteDataString, v3Dp);
                                            lblVA3.Text = va3;
                                            break;
                                        case "34":
                                            peakAmp1 = Helper.HexStringToDecimal(next4ByteDataString, i1Dp);
                                            lblPeakAmp1.Text = peakAmp1;
                                            break;
                                        case "35":
                                            peakAmp2 = Helper.HexStringToDecimal(next4ByteDataString, i2Dp);
                                            lblPeakAmp2.Text = peakAmp2;
                                            break;
                                        case "36":
                                            peakAmp3 = Helper.HexStringToDecimal(next4ByteDataString, i3Dp);
                                            lblPeakAmp3.Text = peakAmp3;
                                            break;
                                        case "37":
                                            meanAmp1 = Helper.HexStringToDecimal(next4ByteDataString, i1Dp);
                                            lblMeanAmp1.Text = meanAmp1;
                                            break;
                                        case "38":
                                            meanAmp2 = Helper.HexStringToDecimal(next4ByteDataString, i2Dp);
                                            lblMeanAmp2.Text = meanAmp2;
                                            break;
                                        case "39":
                                            meanAmp3 = Helper.HexStringToDecimal(next4ByteDataString, i3Dp);
                                            lblMeanAmp3.Text = meanAmp3;
                                            break;
                                        case "40":
                                            string temperatureDataString = hexStringArray[i + 3] + hexStringArray[i + 4];
                                            temperature = Helper.HexStringToDecimal(temperatureDataString, temperatureDp);
                                            lblTemperature.Text = temperature;
                                            break;
                                        case "41":
                                            break;
                                        case "42":
                                            break;
                                        case "43":
                                            break;
                                        case "44":
                                            break;
                                        case "45":
                                            break;
                                        case "46":
                                            break;
                                        case "47":
                                            break;
                                        case "48":
                                            break;
                                        case "49":
                                            break;
                                        case "50":
                                            break;
                                        case "51":
                                            break;
                                        case "52":
                                            break;
                                        case "53":
                                            break;
                                        case "54":
                                            systemMeanVolt = Helper.HexStringToDecimal(next4ByteDataString, sysVDp);
                                            lblSystemMeanV.Text = systemMeanVolt;
                                            break;
                                        case "55":
                                            systemPeakAmp = Helper.HexStringToDecimal(next4ByteDataString, sysIDp);
                                            lblSystemPeakA.Text = systemPeakAmp;
                                            break;
                                        case "56":
                                            systemPeakVolt = Helper.HexStringToDecimal(next4ByteDataString, sysVDp);
                                            lblSystemPeakV.Text = systemPeakVolt;
                                            break;
                                        case "57":
                                            systemMeanAmp = Helper.HexStringToDecimal(next4ByteDataString, sysIDp);
                                            lblSystemMeanA.Text = systemMeanAmp;
                                            break;
                                        case "58":
                                            FfPh1V = Helper.HexStringToDecimal(next3ByteAfterDPDataString, ffPh1Vdp);
                                            lblFFPH1V.Text = FfPh1V;
                                            break;
                                        case "59":
                                            FfPh2V = Helper.HexStringToDecimal(next3ByteAfterDPDataString, ffPh2Vdp);
                                            lblFFPH2V.Text = FfPh2V;
                                            break;
                                        case "60":
                                            FfPh3V = Helper.HexStringToDecimal(next3ByteAfterDPDataString, ffPh3Vdp);
                                            lblFFPH3V.Text = FfPh3V;
                                            break;
                                        case "61":
                                            FfPhSysV = Helper.HexStringToDecimal(next3ByteAfterDPDataString, ffPhSysVdp);
                                            lblFFPHSYSV.Text = FfPhSysV;
                                            break;
                                        case "62":
                                            CfPh1V = Helper.HexStringToDecimal(next3ByteAfterDPDataString, cfPh1Vdp);
                                            lblCFPH1V.Text = CfPh1V;
                                            break;
                                        case "63":
                                            CfPh2V = Helper.HexStringToDecimal(next3ByteAfterDPDataString, cfPh2Vdp);
                                            lblCFPH2V.Text = CfPh2V;
                                            break;
                                        case "64":
                                            CfPh3V = Helper.HexStringToDecimal(next3ByteAfterDPDataString, cfPh3Vdp);
                                            lblCFPH3V.Text = CfPh3V;
                                            break;
                                        case "65":
                                            CfPhSysV = Helper.HexStringToDecimal(next3ByteAfterDPDataString, cfPhSysVdp);
                                            lblCFPHSYSV.Text = CfPhSysV;
                                            break;
                                        case "66":
                                            FfPh1A = Helper.HexStringToDecimal(next3ByteAfterDPDataString, ffPh1Adp);
                                            lblFFPH1A.Text = FfPh1A;
                                            break;
                                        case "67":
                                            FfPh2A = Helper.HexStringToDecimal(next3ByteAfterDPDataString, ffPh2Adp);
                                            lblFFPH2A.Text = FfPh2A;
                                            break;
                                        case "68":
                                            FfPh3A = Helper.HexStringToDecimal(next3ByteAfterDPDataString, ffPh3Adp);
                                            lblFFPH3A.Text = FfPh3A;
                                            break;
                                        case "69":
                                            FfPhSysA = Helper.HexStringToDecimal(next3ByteAfterDPDataString, ffPhSysAdp);
                                            lblFFPHSYSA.Text = FfPhSysA;
                                            break;
                                        case "70":
                                            CfPh1A = Helper.HexStringToDecimal(next3ByteAfterDPDataString, cfPh1Adp);
                                            lblCFPH1A.Text = CfPh1A;
                                            break;
                                        case "71":
                                            CfPh2A = Helper.HexStringToDecimal(next3ByteAfterDPDataString, cfPh2Adp);
                                            lblCFPH2A.Text = CfPh2A;
                                            break;
                                        case "72":
                                            CfPh3A = Helper.HexStringToDecimal(next3ByteAfterDPDataString, cfPh3Adp);
                                            lblCFPH3A.Text = CfPh3A;
                                            break;
                                        case "73":
                                            CfPhSysA = Helper.HexStringToDecimal(next3ByteAfterDPDataString, cfPhSysAdp);
                                            lblCFPHSYSA.Text = CfPhSysA;
                                            break;
                                        case "74":
                                            break;
                                        case "75":
                                            break;
                                        case "76":
                                            break;
                                        case "77":
                                            break;
                                        case "78":
                                            break;
                                        case "79":
                                            break;
                                        case "80":
                                            break;
                                        default:
                                            break;
                                    }

                                }

                            }
                            catch (IndexOutOfRangeException ex)
                            {
                                Console.WriteLine("Error: Index out of range. " + ex.Message);
                            }
                            //catch(Exception ex)
                            //{
                            //    Console.WriteLine("Exception : " + ex.Message);
                            //}
                            
                            #endregion

                        }));
                    }
                }
            }
            catch (Exception ex)
            {
                this.BeginInvoke(new Action(() =>
                {
                    // txtHex.AppendText("Error: " + ex.Message + Environment.NewLine);
                    Console.WriteLine("Error : " + ex.Message);
                }));
            }
        }

        #endregion

        #region Log Functionality

        private void btnLog_Click(object sender, EventArgs e)
        {
            dataGrid.Rows.Add(v1, v2, v3, systemVolt, amp1, amp2, amp3, systemAmp, watt1, watt2, watt3, systemWatt, pf1, pf2, pf3, systemPF, va1, va2, va3, systemVA, frequency, temperature, meanVolt1, meanVolt2, meanVolt3, systemMeanVolt, peakAmp1, peakAmp2, peakAmp3, systemPeakAmp, peakVolt1, peakVolt2, peakVolt3, systemPeakVolt, meanAmp1, meanAmp2, meanAmp3, systemMeanAmp, FfPh1V, FfPh2V, FfPh3V, FfPhSysV, CfPh1V, CfPh2V, CfPh3V, CfPhSysV, FfPh1A, FfPh2A, FfPh3A, FfPhSysA, CfPh1A, CfPh2A, CfPh3A, CfPhSysA);
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            dataGrid.Rows.Clear();
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (dataGrid.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Convert DataGridView to DataTable
            DataTable dt = new DataTable();

            foreach (DataGridViewColumn col in dataGrid.Columns)
            {
                dt.Columns.Add(col.HeaderText);
            }

            foreach (DataGridViewRow row in dataGrid.Rows)
            {
                if (!row.IsNewRow)
                {
                    DataRow dRow = dt.NewRow();
                    for (int i = 0; i < dataGrid.Columns.Count; i++)
                    {
                        dRow[i] = row.Cells[i].Value?.ToString();
                    }
                    dt.Rows.Add(dRow);
                }
            }

            // Save File Dialog
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Workbook|*.xlsx",
                Title = "Save Excel File"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dt, "Sheet1");
                    wb.SaveAs(sfd.FileName);
                }

                MessageBox.Show("Data Exported Successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion

        #region Preference

        public void LoadPreference()
        {
            for(int i=1;i<=54;i++)
            {
                Label lbl = Controls.Find("titleM" + i, true).FirstOrDefault() as Label;

                if(lbl != null)
                {
                    lbl.Text = Prefs.Get("m" + i);
                }
            }

            //titleM1.Text = Prefs.Get("m1", "Volt 1");
            //titleM2.Text = Prefs.Get("m2", "Volt 2");
            //titleM3.Text = Prefs.Get("m3", "Volt 3");
            //titleM4.Text = Prefs.Get("m4", "System Volt");
            //titleM5.Text = Prefs.Get("m5", "Amp 1");
            //titleM6.Text = Prefs.Get("m6", "Amp 2");
            //titleM7.Text = Prefs.Get("m7", "Amp 3");
            //titleM8.Text = Prefs.Get("m8", "System Amp");
            //titleM9.Text = Prefs.Get("m9", "Watt 1");
            //titleM10.Text = Prefs.Get("m10", "Watt 2");
            //titleM11.Text = Prefs.Get("m11", "Watt 3");
            //titleM12.Text = Prefs.Get("m12", "System Watt");
            //titleM13.Text = Prefs.Get("m13", "PF - 1");
            //titleM14.Text = Prefs.Get("m14", "PF - 2");
            //titleM15.Text = Prefs.Get("m15", "PF - 3");
            //titleM16.Text = Prefs.Get("m16", "System PF");
            //titleM17.Text = Prefs.Get("m17", "Frequency");
            //titleM18.Text = Prefs.Get("m18", "VA 1");
            //titleM19.Text = Prefs.Get("m19", "VA 2");
            //titleM20.Text = Prefs.Get("m20", "VA 3");
            //titleM21.Text = Prefs.Get("m21", "System VA");
            //titleM22.Text = Prefs.Get("m22", "Mean Volt 1");
            //titleM23.Text = Prefs.Get("m23", "Mean Volt 2");
            //titleM24.Text = Prefs.Get("m24", "Mean Volt 3");
            //titleM25.Text = Prefs.Get("m25", "System Mean V");
            //titleM26.Text = Prefs.Get("m26", "Peak Amp 1");
            //titleM27.Text = Prefs.Get("m27", "Peak Amp 2");
            //titleM28.Text = Prefs.Get("m28", "Peak Amp 3");
            //titleM29.Text = Prefs.Get("m29", "System Peak A");
            //titleM30.Text = Prefs.Get("m30", "Peak Volt 1");
            //titleM31.Text = Prefs.Get("m31", "Peak Volt 2");
            //titleM32.Text = Prefs.Get("m32", "Peak Volt 3");
            //titleM33.Text = Prefs.Get("m33", "System Peak V");
            //titleM34.Text = Prefs.Get("m34", "Mean Amp 1");
            //titleM35.Text = Prefs.Get("m35", "Mean Amp 2");
            //titleM36.Text = Prefs.Get("m36", "Mean Amp 3");
            //titleM37.Text = Prefs.Get("m37", "System Mean A");
            //titleM38.Text = Prefs.Get("m38", "FF-PH-1-V");
            //titleM39.Text = Prefs.Get("m39", "FF-PH-2-V");
            //titleM40.Text = Prefs.Get("m40", "FF-PH-3-V");
            //titleM41.Text = Prefs.Get("m41", "FF-SYS-V");
            //titleM42.Text = Prefs.Get("m42", "CF-PH-1-V");
            //titleM43.Text = Prefs.Get("m43", "CF-PH-2-V");
            //titleM44.Text = Prefs.Get("m44", "CF-PH-3-V");
            //titleM45.Text = Prefs.Get("m45", "CF-SYS-V");
            //titleM46.Text = Prefs.Get("m46", "FF-PH-1-A");
            //titleM47.Text = Prefs.Get("m47", "FF-PH-2-A");
            //titleM48.Text = Prefs.Get("m48", "FF-PH-3-A");
            //titleM49.Text = Prefs.Get("m49", "FF-SYS-A");
            //titleM50.Text = Prefs.Get("m50", "CF-PH-1-A");
            //titleM51.Text = Prefs.Get("m51", "CF-PH-2-A");
            //titleM52.Text = Prefs.Get("m52", "CF-PH-3-A");
            //titleM53.Text = Prefs.Get("m53", "Temperature");
            //titleM54.Text = Prefs.Get("m54", "CF-SYS-A");

            SetDataGridColumnTitle();
        }

        private void SetDataGridColumnTitle()
        {
            for(int i=0;i<=dataGrid.Columns.Count-1;i++)
            {
                Label lbl = Controls.Find("titleM" + (i+1), true).FirstOrDefault() as Label;
                if (lbl != null)
                {
                    dataGrid.Columns[i].HeaderText = lbl.Text;
                }
            }
        }

        private void ResetPreference()
        {
            Prefs.Set("m1", "Volt 1");
            Prefs.Set("m2", "Volt 2");
            Prefs.Set("m3", "Volt 3");
            Prefs.Set("m4", "System Volt");
            Prefs.Set("m5", "Amp 1");
            Prefs.Set("m6", "Amp 2");
            Prefs.Set("m7", "Amp 3");
            Prefs.Set("m8", "System Amp");
            Prefs.Set("m9", "Watt 1");
            Prefs.Set("m10", "Watt 2");
            Prefs.Set("m11", "Watt 3");
            Prefs.Set("m12", "System Watt");
            Prefs.Set("m13", "PF - 1");
            Prefs.Set("m14", "PF - 2");
            Prefs.Set("m15", "PF - 3");
            Prefs.Set("m16", "System PF");
            Prefs.Set("m17", "Frequency");
            Prefs.Set("m18", "VA 1");
            Prefs.Set("m19", "VA 2");
            Prefs.Set("m20", "VA 3");
            Prefs.Set("m21", "System VA");
            Prefs.Set("m22", "Mean Volt 1");
            Prefs.Set("m23", "Mean Volt 2");
            Prefs.Set("m24", "Mean Volt 3");
            Prefs.Set("m25", "System Mean V");
            Prefs.Set("m26", "Peak Amp 1");
            Prefs.Set("m27", "Peak Amp 2");
            Prefs.Set("m28", "Peak Amp 3");
            Prefs.Set("m29", "System Peak A");
            Prefs.Set("m30", "Peak Volt 1");
            Prefs.Set("m31", "Peak Volt 2");
            Prefs.Set("m32", "Peak Volt 3");
            Prefs.Set("m33", "System Peak V");
            Prefs.Set("m34", "Mean Amp 1");
            Prefs.Set("m35", "Mean Amp 2");
            Prefs.Set("m36", "Mean Amp 3");
            Prefs.Set("m37", "System Mean A");
            Prefs.Set("m38", "FF-PH-1-V");
            Prefs.Set("m39", "FF-PH-2-V");
            Prefs.Set("m40", "FF-PH-3-V");
            Prefs.Set("m41", "FF-SYS-V");
            Prefs.Set("m42", "CF-PH-1-V");
            Prefs.Set("m43", "CF-PH-2-V");
            Prefs.Set("m44", "CF-PH-3-V");
            Prefs.Set("m45", "CF-SYS-V");
            Prefs.Set("m46", "FF-PH-1-A");
            Prefs.Set("m47", "FF-PH-2-A");
            Prefs.Set("m48", "FF-PH-3-A");
            Prefs.Set("m49", "FF-SYS-A");
            Prefs.Set("m50", "CF-PH-1-A");
            Prefs.Set("m51", "CF-PH-2-A");
            Prefs.Set("m52", "CF-PH-3-A");
            Prefs.Set("m53", "Temperature");
            Prefs.Set("m54", "CF-SYS-A");
        }


        #endregion

        #region Button - Refresh, Settings, Fast, Hold, Accurate

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            toggleSwitchConnectDisconnect.Checked = false;            
            LoadAvailablePorts();
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            SettingsForm settingsForm = new SettingsForm(this);
            settingsForm.ShowDialog();
        }

        private void btnHold_Click(object sender, EventArgs e)
        {

        }

        private void btnAccurate_Click(object sender, EventArgs e)
        {

        }

        private void btnFast_Click(object sender, EventArgs e)
        {

        }

        #endregion
    }
}
