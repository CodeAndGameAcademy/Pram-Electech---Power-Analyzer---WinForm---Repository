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
                
                btnLog.Enabled = true;
                btnClearAll.Enabled = true;

                btnFast.Enabled = true;
                btnAccurate.Enabled = true;
                btnHold.Enabled = true;
            }
            else
            {
                btnLog.Enabled = false;
                btnClearAll.Enabled = false;

                btnFast.Enabled = false;
                btnAccurate.Enabled = false;
                btnHold.Enabled = false;

                Disconnect();
            }
        }

        private void Connect()
        {
            try
            {
                if (cmbPort.SelectedItem == null || cmbPort.SelectedItem.ToString() != "COM3")
                {
                    toggleSwitchConnectDisconnect.Checked = false;
                    MessageBox.Show("No valid COM port selected!");
                    return;
                }

                if (_serialPort.IsOpen) _serialPort.Close();

                _serialPort.PortName = cmbPort.SelectedItem.ToString();
                _serialPort.Open();
                
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

                            #region DP                                      
                            string v1DP = hexStringArray[224];
                            string v2DP = hexStringArray[223];
                            string v3DP = hexStringArray[222];
                            string systemVDP = hexStringArray[222];

                            string amp1DP = hexStringArray[229];
                            string amp2DP = hexStringArray[228];
                            string amp3DP = hexStringArray[227];
                            string systemAmpDP = hexStringArray[227];

                            string watt1DP = hexStringArray[234];
                            string watt2DP = hexStringArray[233];
                            string watt3DP = hexStringArray[232];
                            string systemWattDP = hexStringArray[232];

                            #endregion

                            // v1, v2, v3
                            string v1String = hexStringArray[6] + hexStringArray[7] + hexStringArray[8] + hexStringArray[9];
                            v1 = Helper.HexStringToDecimal(v1String, v1DP);
                            lblV1.Text = v1;

                            string v2String = hexStringArray[11] + hexStringArray[12] + hexStringArray[13] + hexStringArray[14];
                            v2 = Helper.HexStringToDecimal(v2String, v2DP);                            
                            lblV2.Text = v2;
                            

                            string v3String = hexStringArray[16] + hexStringArray[17] + hexStringArray[18] + hexStringArray[19];                            
                            v3 = Helper.HexStringToDecimal(v3String, v3DP);                            
                            lblV3.Text = v3;

                            string systemVoltString = hexStringArray[71] + hexStringArray[72] + hexStringArray[73] + hexStringArray[74];
                            systemVolt = Helper.HexStringToDecimal(systemVoltString, systemVDP);
                            lblSystemVolt.Text = systemVolt;


                            // Amp1, Amp2, Amp3
                            string amp1String = hexStringArray[21] + hexStringArray[22] + hexStringArray[23] + hexStringArray[24];
                            amp1 = Helper.HexStringToDecimal(amp1String, amp1DP);
                            lblAmp1.Text = amp1;

                            string amp2String = hexStringArray[26] + hexStringArray[27] + hexStringArray[28] + hexStringArray[29];
                            amp2 = Helper.HexStringToDecimal(amp2String, amp2DP);
                            lblAmp2.Text = amp2;

                            string amp3String = hexStringArray[31] + hexStringArray[32] + hexStringArray[33] + hexStringArray[34];
                            amp3 = Helper.HexStringToDecimal(amp3String, amp3DP);
                            lblAmp3.Text = amp3;

                            string systemAmpString = hexStringArray[76] + hexStringArray[77] + hexStringArray[78] + hexStringArray[79];
                            systemAmp = Helper.HexStringToDecimal(systemAmpString, systemAmpDP);
                            lblSystemAmp.Text = systemAmp;


                            // watt1, watt2, watt3
                            string watt1String = hexStringArray[36] + hexStringArray[37] + hexStringArray[38] + hexStringArray[39];
                            watt1 = Helper.HexStringToDecimal(watt1String, watt1DP);
                            lblWatt1.Text = watt1;

                            string watt2String = hexStringArray[41] + hexStringArray[42] + hexStringArray[43] + hexStringArray[44];
                            watt2 = Helper.HexStringToDecimal(watt2String, watt2DP);
                            lblWatt2.Text = watt2;

                            string watt3String = hexStringArray[46] + hexStringArray[47] + hexStringArray[48] + hexStringArray[49];
                            watt3 = Helper.HexStringToDecimal(watt3String, watt3DP);
                            lblWatt3.Text = watt3;

                            string systemWattString = hexStringArray[81] + hexStringArray[82] + hexStringArray[83] + hexStringArray[84];
                            systemWatt = Helper.HexStringToDecimal(systemWattString, systemWattDP);
                            lblSystemWatt.Text = systemWatt;

                            
                            // pf1, pf2, pf3
                            string pf1String = hexStringArray[52] + hexStringArray[53] + hexStringArray[54];
                            pf1 = Helper.HexStringToDecimal(pf1String, hexStringArray[51]);
                            lblPF1.Text = pf1;

                            string pf2String = hexStringArray[57] + hexStringArray[58] + hexStringArray[59];
                            pf2 = Helper.HexStringToDecimal(pf2String, hexStringArray[56]);
                            lblPF2.Text = pf2;

                            string pf3String = hexStringArray[62] + hexStringArray[63] + hexStringArray[64];
                            pf3 = Helper.HexStringToDecimal(pf3String, hexStringArray[61]);
                            lblPF3.Text = pf3;

                            string systemPFString = hexStringArray[87] + hexStringArray[88] + hexStringArray[89];
                            systemPF = Helper.HexStringToDecimal(systemWattString, hexStringArray[86]);
                            lblSystemPF.Text = systemPF;


                            // va1, va2, va3
                            string va1String = hexStringArray[131] + hexStringArray[132] + hexStringArray[133] + hexStringArray[134];
                            va1 = Helper.HexStringToDecimal(va1String, v1DP);
                            lblVA1.Text = va1;

                            string va2String = hexStringArray[136] + hexStringArray[137] + hexStringArray[138] + hexStringArray[139];
                            va2 = Helper.HexStringToDecimal(va2String, v2DP);
                            lblVA2.Text = va2;

                            string va3String = hexStringArray[141] + hexStringArray[142] + hexStringArray[143] + hexStringArray[144];
                            va3 = Helper.HexStringToDecimal(va3String, v3DP);
                            lblVA3.Text = va3;

                            string systemVAString = hexStringArray[91] + hexStringArray[92] + hexStringArray[93] + hexStringArray[94];
                            systemVA = Helper.HexStringToDecimal(systemVAString, systemVDP);
                            lblSystemVA.Text = systemVA;


                            // Frequency, Temperature                            
                            string frequencyString = hexStringArray[67] + hexStringArray[68] + hexStringArray[69];
                            frequency = Helper.HexStringToDecimal(frequencyString, hexStringArray[66]);
                            lblFrequency.Text = frequency;

                            string temperatureString = hexStringArray[178] + hexStringArray[179];
                            temperature = Helper.HexStringToDecimal(temperatureString, hexStringArray[176]);
                            lblTemperature.Text = temperature;


                            // Mean-Volt1, Mean-Volt2, Mean-Volt3
                            string meanVolt1String = hexStringArray[116] + hexStringArray[117] + hexStringArray[118] + hexStringArray[119];
                            meanVolt1 = Helper.HexStringToDecimal(meanVolt1String, v1DP);
                            lblMeanVolt1.Text = meanVolt1.ToString();

                            string meanVolt2String = hexStringArray[121] + hexStringArray[122] + hexStringArray[123] + hexStringArray[124];
                            meanVolt2 = Helper.HexStringToDecimal(meanVolt2String, v2DP);
                            lblMeanVolt2.Text = meanVolt2.ToString();

                            string meanVolt3String = hexStringArray[126] + hexStringArray[127] + hexStringArray[128] + hexStringArray[129];
                            meanVolt3 = Helper.HexStringToDecimal(meanVolt3String, v3DP);
                            lblMeanVolt3.Text = meanVolt3.ToString();

                            string systemMeanVoltString = hexStringArray[246] + hexStringArray[247] + hexStringArray[248] + hexStringArray[249];
                            systemMeanVolt = Helper.HexStringToDecimal(systemMeanVoltString, systemVDP);
                            lblSystemMeanV.Text = systemMeanVolt;

                            
                            // Peak-Amp1, Peak-Amp2, Peak-Amp3 
                            string peakAmp1String = hexStringArray[146] + hexStringArray[147] + hexStringArray[148] + hexStringArray[149];
                            peakAmp1 = Helper.HexStringToDecimal(peakAmp1String, amp1DP);
                            lblPeakAmp1.Text = peakAmp1;

                            string peakAmp2String = hexStringArray[151] + hexStringArray[152] + hexStringArray[153] + hexStringArray[154];
                            peakAmp2 = Helper.HexStringToDecimal(peakAmp2String, amp2DP);
                            lblPeakAmp2.Text = peakAmp2;

                            string peakAmp3String = hexStringArray[156] + hexStringArray[157] + hexStringArray[158] + hexStringArray[159];
                            peakAmp3 = Helper.HexStringToDecimal(peakAmp3String, amp3DP);
                            lblPeakAmp3.Text = peakAmp3;

                            string systemPeakAmpString = hexStringArray[251] + hexStringArray[252] + hexStringArray[253] + hexStringArray[254];
                            systemPeakAmp = Helper.HexStringToDecimal(systemPeakAmpString, systemAmpDP);
                            lblSystemPeakA.Text = systemPeakAmp;


                            // Peak-Volt1, Peak-Volt2, Peak-Volt3
                            string peakVolt1String = hexStringArray[101] + hexStringArray[102] + hexStringArray[103] + hexStringArray[104];
                            peakVolt1 = Helper.HexStringToDecimal(peakVolt1String, v1DP);
                            lblPeakVolt1.Text = peakVolt1;

                            string peakVolt2String = hexStringArray[106] + hexStringArray[107] + hexStringArray[108] + hexStringArray[109];
                            peakVolt2 = Helper.HexStringToDecimal(peakVolt2String, v2DP);
                            lblPeakVolt2.Text = peakVolt2;

                            string peakVolt3String = hexStringArray[111] + hexStringArray[112] + hexStringArray[113] + hexStringArray[114];
                            peakVolt3 = Helper.HexStringToDecimal(peakVolt3String, v3DP);
                            lblPeakVolt3.Text = peakVolt3;

                            string systemPeakVoltString = hexStringArray[256] + hexStringArray[257] + hexStringArray[258] + hexStringArray[259];
                            systemPeakVolt = Helper.HexStringToDecimal(systemPeakVoltString, systemVDP);
                            lblSystemPeakV.Text = systemPeakVolt;


                            // Mean-Amp1, Mean-Amp2, Mean-Amp3
                            string meanAmp1String = hexStringArray[161] + hexStringArray[162] + hexStringArray[163] + hexStringArray[164];
                            meanAmp1 = Helper.HexStringToDecimal(meanAmp1String, amp1DP);
                            lblMeanAmp1.Text = meanAmp1;

                            string meanAmp2String = hexStringArray[166] + hexStringArray[167] + hexStringArray[168] + hexStringArray[169];
                            meanAmp2 = Helper.HexStringToDecimal(meanAmp2String, amp2DP);
                            lblMeanAmp2.Text = meanAmp2;

                            string meanAmp3String = hexStringArray[171] + hexStringArray[172] + hexStringArray[173] + hexStringArray[174];
                            meanAmp3 = Helper.HexStringToDecimal(meanAmp3String, amp3DP);
                            lblMeanAmp3.Text = meanAmp3;

                            string systemMeanAmpString = hexStringArray[261] + hexStringArray[262] + hexStringArray[263] + hexStringArray[264];
                            systemMeanAmp = Helper.HexStringToDecimal(systemMeanAmpString, systemAmpDP);
                            lblSystemMeanA.Text = systemMeanAmp;


                            // FF-PH-1-V, FF-PH-2-V, FF-PH-3-V
                            string FfPh1VString = hexStringArray[267] + hexStringArray[268] + hexStringArray[269];
                            FfPh1V = Helper.HexStringToDecimal(FfPh1VString, hexStringArray[266]);
                            lblFFPH1V.Text = FfPh1V;

                            string FfPh2VString = hexStringArray[272] + hexStringArray[273] + hexStringArray[274];
                            FfPh2V = Helper.HexStringToDecimal(FfPh2VString, hexStringArray[271]);
                            lblFFPH2V.Text = FfPh2V;

                            string FfPh3VString = hexStringArray[277] + hexStringArray[278] + hexStringArray[279];
                            FfPh3V = Helper.HexStringToDecimal(FfPh3VString, hexStringArray[276]);
                            lblFFPH3V.Text = FfPh3V;

                            string FfPhSysVString = hexStringArray[282] + hexStringArray[283] + hexStringArray[284];
                            FfPhSysV = Helper.HexStringToDecimal(FfPhSysVString, hexStringArray[281]);
                            lblFFPHSYSV.Text = FfPhSysV;


                            // CF-PH-1-V, CF-PH-2-V, CF-PH-3-V
                            string CfPh1VString = hexStringArray[287] + hexStringArray[288] + hexStringArray[289];
                            CfPh1V = Helper.HexStringToDecimal(CfPh1VString, hexStringArray[286]);
                            lblCFPH1V.Text = CfPh1V;

                            string CfPh2VString = hexStringArray[292] + hexStringArray[293] + hexStringArray[294];
                            CfPh2V = Helper.HexStringToDecimal(CfPh2VString, hexStringArray[291]);
                            lblCFPH2V.Text = CfPh2V;

                            string CfPh3VString = hexStringArray[297] + hexStringArray[298] + hexStringArray[299];
                            CfPh3V = Helper.HexStringToDecimal(CfPh3VString, hexStringArray[296]);
                            lblCFPH3V.Text = CfPh3V;

                            string CfPhSysVString = hexStringArray[302] + hexStringArray[303] + hexStringArray[304];
                            CfPhSysV = Helper.HexStringToDecimal(CfPhSysVString, hexStringArray[301]);
                            lblCFPHSYSV.Text = CfPhSysV;


                            // FF-PH-1-A, FF-PH-2-A, FF-PH-3-A
                            string FfPh1AString = hexStringArray[307] + hexStringArray[308] + hexStringArray[309];
                            FfPh1A = Helper.HexStringToDecimal(FfPh1AString, hexStringArray[306]);
                            lblFFPH1A.Text = FfPh1A;

                            string FfPh2AString = hexStringArray[312] + hexStringArray[313] + hexStringArray[314];
                            FfPh2A = Helper.HexStringToDecimal(FfPh2AString, hexStringArray[311]);
                            lblFFPH2A.Text = FfPh2A;

                            string FfPh3AString = hexStringArray[317] + hexStringArray[318] + hexStringArray[319];
                            FfPh3A = Helper.HexStringToDecimal(FfPh3AString, hexStringArray[316]);
                            lblFFPH3A.Text = FfPh3A;

                            string FfPhSysAString = hexStringArray[322] + hexStringArray[323] + hexStringArray[324];
                            FfPhSysA = Helper.HexStringToDecimal(FfPhSysAString, hexStringArray[321]);
                            lblFFPHSYSA.Text = FfPhSysA;


                            // CF-PH-1-A, CF-PH-2-A, CF-PH-3-A
                            string CfPh1AString = hexStringArray[327] + hexStringArray[328] + hexStringArray[329];
                            CfPh1A = Helper.HexStringToDecimal(CfPh1AString, hexStringArray[326]);
                            lblCFPH1A.Text = CfPh1A;

                            string CfPh2AString = hexStringArray[332] + hexStringArray[333] + hexStringArray[334];
                            CfPh2A = Helper.HexStringToDecimal(CfPh2AString, hexStringArray[331]);
                            lblCFPH2A.Text = CfPh2A;

                            string CfPh3AString = hexStringArray[337] + hexStringArray[338] + hexStringArray[339];
                            CfPh3A = Helper.HexStringToDecimal(CfPh3AString, hexStringArray[336]);
                            lblCFPH3A.Text = CfPh3A;

                            string CfPhSysAString = hexStringArray[342] + hexStringArray[343] + hexStringArray[344];
                            CfPhSysA = Helper.HexStringToDecimal(CfPhSysAString, hexStringArray[341]);
                            lblCFPHSYSA.Text = CfPhSysA;

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

        #endregion

        #region Preference

        private void LoadPreference()
        {
            titleM1.Text = Prefs.Get("m1", "Volt 1");
            titleM2.Text = Prefs.Get("m2", "Volt 2");
            titleM3.Text = Prefs.Get("m3", "Volt 3");
            titleM4.Text = Prefs.Get("m4", "System Volt");
            titleM5.Text = Prefs.Get("m5", "Amp 1");
            titleM6.Text = Prefs.Get("m6", "Amp 2");
            titleM7.Text = Prefs.Get("m7", "Amp 3");
            titleM8.Text = Prefs.Get("m8", "System Amp");
            titleM9.Text = Prefs.Get("m9", "Watt 1");
            titleM10.Text = Prefs.Get("m10", "Watt 2");
            titleM11.Text = Prefs.Get("m11", "Watt 3");
            titleM12.Text = Prefs.Get("m12", "System Watt");
            titleM13.Text = Prefs.Get("m13", "PF - 1");
            titleM14.Text = Prefs.Get("m14", "PF - 2");
            titleM15.Text = Prefs.Get("m15", "PF - 3");
            titleM16.Text = Prefs.Get("m16", "System PF");
            titleM17.Text = Prefs.Get("m17", "Frequency");
            titleM18.Text = Prefs.Get("m18", "VA 1");
            titleM19.Text = Prefs.Get("m19", "VA 2");
            titleM20.Text = Prefs.Get("m20", "VA 3");
            titleM21.Text = Prefs.Get("m21", "System VA");
            titleM22.Text = Prefs.Get("m22", "Mean Volt 1");
            titleM23.Text = Prefs.Get("m23", "Mean Volt 2");
            titleM24.Text = Prefs.Get("m24", "Mean Volt 3");
            titleM25.Text = Prefs.Get("m25", "System Mean V");
            titleM26.Text = Prefs.Get("m26", "Peak Amp 1");
            titleM27.Text = Prefs.Get("m27", "Peak Amp 2");
            titleM28.Text = Prefs.Get("m28", "Peak Amp 3");
            titleM29.Text = Prefs.Get("m29", "System Peak A");
            titleM30.Text = Prefs.Get("m30", "Peak Volt 1");
            titleM31.Text = Prefs.Get("m31", "Peak Volt 2");
            titleM32.Text = Prefs.Get("m32", "Peak Volt 3");
            titleM33.Text = Prefs.Get("m33", "System Peak V");
            titleM34.Text = Prefs.Get("m34", "Mean Amp 1");
            titleM35.Text = Prefs.Get("m35", "Mean Amp 2");
            titleM36.Text = Prefs.Get("m36", "Mean Amp 3");
            titleM37.Text = Prefs.Get("m37", "System Mean A");
            titleM38.Text = Prefs.Get("m38", "FF-PH-1-V");
            titleM39.Text = Prefs.Get("m39", "FF-PH-2-V");
            titleM40.Text = Prefs.Get("m40", "FF-PH-3-V");
            titleM41.Text = Prefs.Get("m41", "FF-SYS-V");
            titleM42.Text = Prefs.Get("m42", "CF-PH-1-V");
            titleM43.Text = Prefs.Get("m43", "CF-PH-2-V");
            titleM44.Text = Prefs.Get("m44", "CF-PH-3-V");
            titleM45.Text = Prefs.Get("m45", "CF-SYS-V");
            titleM46.Text = Prefs.Get("m46", "FF-PH-1-A");
            titleM47.Text = Prefs.Get("m47", "FF-PH-2-A");
            titleM48.Text = Prefs.Get("m48", "FF-PH-3-A");
            titleM49.Text = Prefs.Get("m49", "FF-SYS-A");
            titleM50.Text = Prefs.Get("m50", "CF-PH-1-A");
            titleM51.Text = Prefs.Get("m51", "CF-PH-2-A");
            titleM52.Text = Prefs.Get("m52", "CF-PH-3-A");
            titleM53.Text = Prefs.Get("m53", "Temperature");
            titleM54.Text = Prefs.Get("m54", "CF-SYS-A");
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

        #region Button - Refresh, Fast, Hold, Accurate

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            toggleSwitchConnectDisconnect.Checked = false;            
            LoadAvailablePorts();
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
