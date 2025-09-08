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
                            string v1 = Helper.HexStringToDecimal(v1String, v1DP);
                            lblV1.Text = v1;

                            string v2String = hexStringArray[11] + hexStringArray[12] + hexStringArray[13] + hexStringArray[14];
                            string v2 = Helper.HexStringToDecimal(v2String, v2DP);                            
                            lblV2.Text = v2;
                            

                            string v3String = hexStringArray[16] + hexStringArray[17] + hexStringArray[18] + hexStringArray[19];                            
                            string v3 = Helper.HexStringToDecimal(v3String, v3DP);                            
                            lblV3.Text = v3;

                            string systemVoltString = hexStringArray[71] + hexStringArray[72] + hexStringArray[73] + hexStringArray[74];
                            string systemVolt = Helper.HexStringToDecimal(systemVoltString, systemVDP);
                            lblSystemVolt.Text = systemVolt;


                            // Amp1, Amp2, Amp3
                            string amp1String = hexStringArray[21] + hexStringArray[22] + hexStringArray[23] + hexStringArray[24];
                            string amp1 = Helper.HexStringToDecimal(amp1String, amp1DP);
                            lblAmp1.Text = amp1;

                            string amp2String = hexStringArray[26] + hexStringArray[27] + hexStringArray[28] + hexStringArray[29];
                            string amp2 = Helper.HexStringToDecimal(amp2String, amp2DP);
                            lblAmp2.Text = amp2;

                            string amp3String = hexStringArray[31] + hexStringArray[32] + hexStringArray[33] + hexStringArray[34];
                            string amp3 = Helper.HexStringToDecimal(amp3String, amp3DP);
                            lblAmp3.Text = amp3;

                            string systemAmpString = hexStringArray[76] + hexStringArray[77] + hexStringArray[78] + hexStringArray[79];
                            string systemAmp = Helper.HexStringToDecimal(systemAmpString, systemAmpDP);
                            lblSystemAmp.Text = systemAmp;


                            // watt1, watt2, watt3
                            string watt1String = hexStringArray[36] + hexStringArray[37] + hexStringArray[38] + hexStringArray[39];
                            string watt1 = Helper.HexStringToDecimal(watt1String, watt1DP);
                            lblWatt1.Text = watt1;

                            string watt2String = hexStringArray[41] + hexStringArray[42] + hexStringArray[43] + hexStringArray[44];
                            string watt2 = Helper.HexStringToDecimal(watt2String, watt2DP);
                            lblWatt2.Text = watt2;

                            string watt3String = hexStringArray[46] + hexStringArray[47] + hexStringArray[48] + hexStringArray[49];
                            string watt3 = Helper.HexStringToDecimal(watt3String, watt3DP);
                            lblWatt3.Text = watt3;

                            string systemWattString = hexStringArray[81] + hexStringArray[82] + hexStringArray[83] + hexStringArray[84];
                            string systemWatt = Helper.HexStringToDecimal(systemWattString, systemWattDP);
                            lblSystemWatt.Text = systemWatt;

                            
                            // pf1, pf2, pf3
                            string pf1String = hexStringArray[52] + hexStringArray[53] + hexStringArray[54];
                            string pf1 = Helper.HexStringToDecimal(pf1String, hexStringArray[51]);
                            lblPF1.Text = pf1;

                            string pf2String = hexStringArray[57] + hexStringArray[58] + hexStringArray[59];
                            string pf2 = Helper.HexStringToDecimal(pf2String, hexStringArray[56]);
                            lblPF2.Text = pf2;

                            string pf3String = hexStringArray[62] + hexStringArray[63] + hexStringArray[64];
                            string pf3 = Helper.HexStringToDecimal(pf3String, hexStringArray[61]);
                            lblPF3.Text = pf3;

                            string systemPFString = hexStringArray[87] + hexStringArray[88] + hexStringArray[89];
                            string systemPF = Helper.HexStringToDecimal(systemWattString, hexStringArray[86]);
                            lblSystemPF.Text = systemPF;


                            // va1, va2, va3
                            string va1String = hexStringArray[131] + hexStringArray[132] + hexStringArray[133] + hexStringArray[134];
                            string va1 = Helper.HexStringToDecimal(va1String, v1DP);
                            lblVA1.Text = va1;

                            string va2String = hexStringArray[136] + hexStringArray[137] + hexStringArray[138] + hexStringArray[139];
                            string va2 = Helper.HexStringToDecimal(va2String, v2DP);
                            lblVA2.Text = va2;

                            string va3String = hexStringArray[141] + hexStringArray[142] + hexStringArray[143] + hexStringArray[144];
                            string va3 = Helper.HexStringToDecimal(va3String, v3DP);
                            lblVA3.Text = va3;

                            string systemVAString = hexStringArray[91] + hexStringArray[92] + hexStringArray[93] + hexStringArray[94];
                            string systemVA = Helper.HexStringToDecimal(systemVAString, systemVDP);
                            lblSystemVA.Text = systemVA;


                            // Frequency, Temperature                            
                            string frequencyString = hexStringArray[67] + hexStringArray[68] + hexStringArray[69];
                            string frequency = Helper.HexStringToDecimal(frequencyString, hexStringArray[66]);
                            lblFrequency.Text = frequency;

                            string temperatureString = hexStringArray[178] + hexStringArray[179];
                            string temperature = Helper.HexStringToDecimal(temperatureString, hexStringArray[176]);
                            lblTemperature.Text = temperature;


                            // Mean-Volt1, Mean-Volt2, Mean-Volt3
                            string meanVolt1String = hexStringArray[116] + hexStringArray[117] + hexStringArray[118] + hexStringArray[119];
                            string meanVolt1 = Helper.HexStringToDecimal(meanVolt1String, v1DP);
                            lblMeanVolt1.Text = meanVolt1.ToString();

                            string meanVolt2String = hexStringArray[121] + hexStringArray[122] + hexStringArray[123] + hexStringArray[124];
                            string meanVolt2 = Helper.HexStringToDecimal(meanVolt2String, v2DP);
                            lblMeanVolt2.Text = meanVolt2.ToString();

                            string meanVolt3String = hexStringArray[126] + hexStringArray[127] + hexStringArray[128] + hexStringArray[129];
                            string meanVolt3 = Helper.HexStringToDecimal(meanVolt3String, v3DP);
                            lblMeanVolt3.Text = meanVolt3.ToString();

                            string systemMeanVoltString = hexStringArray[246] + hexStringArray[247] + hexStringArray[248] + hexStringArray[249];
                            string systemMeanVolt = Helper.HexStringToDecimal(systemMeanVoltString, systemVDP);
                            lblSystemMeanV.Text = systemMeanVolt;

                            
                            // Peak-Amp1, Peak-Amp2, Peak-Amp3 
                            string peakAmp1String = hexStringArray[146] + hexStringArray[147] + hexStringArray[148] + hexStringArray[149];
                            string peakAmp1 = Helper.HexStringToDecimal(peakAmp1String, amp1DP);
                            lblPeakAmp1.Text = peakAmp1;

                            string peakAmp2String = hexStringArray[151] + hexStringArray[152] + hexStringArray[153] + hexStringArray[154];
                            string peakAmp2 = Helper.HexStringToDecimal(peakAmp2String, amp2DP);
                            lblPeakAmp2.Text = peakAmp2;

                            string peakAmp3String = hexStringArray[156] + hexStringArray[157] + hexStringArray[158] + hexStringArray[159];
                            string peakAmp3 = Helper.HexStringToDecimal(peakAmp3String, amp3DP);
                            lblPeakAmp3.Text = peakAmp3;

                            string systemPeakAmpString = hexStringArray[251] + hexStringArray[252] + hexStringArray[253] + hexStringArray[254];
                            string systemPeakAmp = Helper.HexStringToDecimal(systemPeakAmpString, systemAmpDP);
                            lblSystemPeakA.Text = systemPeakAmp;


                            // Peak-Volt1, Peak-Volt2, Peak-Volt3
                            string peakVolt1String = hexStringArray[101] + hexStringArray[102] + hexStringArray[103] + hexStringArray[104];
                            string peakVolt1 = Helper.HexStringToDecimal(peakVolt1String, v1DP);
                            lblPeakVolt1.Text = peakVolt1;

                            string peakVolt2String = hexStringArray[106] + hexStringArray[107] + hexStringArray[108] + hexStringArray[109];
                            string peakVolt2 = Helper.HexStringToDecimal(peakVolt2String, v2DP);
                            lblPeakVolt2.Text = peakVolt2;

                            string peakVolt3String = hexStringArray[111] + hexStringArray[112] + hexStringArray[113] + hexStringArray[114];
                            string peakVolt3 = Helper.HexStringToDecimal(peakVolt3String, v3DP);
                            lblPeakVolt3.Text = peakVolt3;

                            string systemPeakVoltString = hexStringArray[256] + hexStringArray[257] + hexStringArray[258] + hexStringArray[259];
                            string systemPeakVolt = Helper.HexStringToDecimal(systemPeakVoltString, systemVDP);
                            lblSystemPeakV.Text = systemPeakVolt;


                            // Mean-Amp1, Mean-Amp2, Mean-Amp3
                            string meanAmp1String = hexStringArray[161] + hexStringArray[162] + hexStringArray[163] + hexStringArray[164];
                            string meanAmp1 = Helper.HexStringToDecimal(meanAmp1String, amp1DP);
                            lblMeanAmp1.Text = meanAmp1;

                            string meanAmp2String = hexStringArray[166] + hexStringArray[167] + hexStringArray[168] + hexStringArray[169];
                            string meanAmp2 = Helper.HexStringToDecimal(meanAmp2String, amp2DP);
                            lblMeanAmp2.Text = meanAmp2;

                            string meanAmp3String = hexStringArray[171] + hexStringArray[172] + hexStringArray[173] + hexStringArray[174];
                            string meanAmp3 = Helper.HexStringToDecimal(meanAmp3String, amp3DP);
                            lblMeanAmp3.Text = meanAmp3;

                            string systemMeanAmpString = hexStringArray[261] + hexStringArray[262] + hexStringArray[263] + hexStringArray[264];
                            string systemMeanAmp = Helper.HexStringToDecimal(systemMeanAmpString, systemAmpDP);
                            lblSystemMeanA.Text = systemMeanAmp;


                            // FF-PH-1-V, FF-PH-2-V, FF-PH-3-V
                            string FfPh1VString = hexStringArray[267] + hexStringArray[268] + hexStringArray[269];
                            string FfPh1V = Helper.HexStringToDecimal(FfPh1VString, hexStringArray[266]);
                            lblFFPH1V.Text = FfPh1V;

                            string FfPh2VString = hexStringArray[272] + hexStringArray[273] + hexStringArray[274];
                            string FfPh2V = Helper.HexStringToDecimal(FfPh2VString, hexStringArray[271]);
                            lblFFPH2V.Text = FfPh2V;

                            string FfPh3VString = hexStringArray[277] + hexStringArray[278] + hexStringArray[279];
                            string FfPh3V = Helper.HexStringToDecimal(FfPh3VString, hexStringArray[276]);
                            lblFFPH3V.Text = FfPh3V;

                            string FfPhSysVString = hexStringArray[282] + hexStringArray[283] + hexStringArray[284];
                            string FfPhSysV = Helper.HexStringToDecimal(FfPhSysVString, hexStringArray[281]);
                            lblFFPHSYSV.Text = FfPhSysV;


                            // CF-PH-1-V, CF-PH-2-V, CF-PH-3-V
                            string CfPh1VString = hexStringArray[287] + hexStringArray[288] + hexStringArray[289];
                            string CfPh1V = Helper.HexStringToDecimal(CfPh1VString, hexStringArray[286]);
                            lblCFPH1V.Text = CfPh1V;

                            string CfPh2VString = hexStringArray[292] + hexStringArray[293] + hexStringArray[294];
                            string CfPh2V = Helper.HexStringToDecimal(CfPh2VString, hexStringArray[291]);
                            lblCFPH2V.Text = CfPh2V;

                            string CfPh3VString = hexStringArray[297] + hexStringArray[298] + hexStringArray[299];
                            string CfPh3V = Helper.HexStringToDecimal(CfPh3VString, hexStringArray[296]);
                            lblCFPH3V.Text = CfPh3V;

                            string CfPhSysVString = hexStringArray[302] + hexStringArray[303] + hexStringArray[304];
                            string CfPhSysV = Helper.HexStringToDecimal(CfPhSysVString, hexStringArray[301]);
                            lblCFPHSYSV.Text = CfPhSysV;


                            // FF-PH-1-A, FF-PH-2-A, FF-PH-3-A
                            string FfPh1AString = hexStringArray[307] + hexStringArray[308] + hexStringArray[309];
                            string FfPh1A = Helper.HexStringToDecimal(FfPh1AString, hexStringArray[306]);
                            lblFFPH1A.Text = FfPh1A;

                            string FfPh2AString = hexStringArray[312] + hexStringArray[313] + hexStringArray[314];
                            string FfPh2A = Helper.HexStringToDecimal(FfPh2AString, hexStringArray[311]);
                            lblFFPH2A.Text = FfPh2A;

                            string FfPh3AString = hexStringArray[317] + hexStringArray[318] + hexStringArray[319];
                            string FfPh3A = Helper.HexStringToDecimal(FfPh3AString, hexStringArray[316]);
                            lblFFPH3A.Text = FfPh3A;

                            string FfPhSysAString = hexStringArray[322] + hexStringArray[323] + hexStringArray[324];
                            string FfPhSysA = Helper.HexStringToDecimal(FfPhSysAString, hexStringArray[321]);
                            lblFFPHSYSA.Text = FfPhSysA;


                            // CF-PH-1-A, CF-PH-2-A, CF-PH-3-A
                            string CfPh1AString = hexStringArray[327] + hexStringArray[328] + hexStringArray[329];
                            string CfPh1A = Helper.HexStringToDecimal(CfPh1AString, hexStringArray[326]);
                            lblCFPH1A.Text = CfPh1A;

                            string CfPh2AString = hexStringArray[332] + hexStringArray[333] + hexStringArray[334];
                            string CfPh2A = Helper.HexStringToDecimal(CfPh2AString, hexStringArray[331]);
                            lblCFPH2A.Text = CfPh2A;

                            string CfPh3AString = hexStringArray[337] + hexStringArray[338] + hexStringArray[339];
                            string CfPh3A = Helper.HexStringToDecimal(CfPh3AString, hexStringArray[336]);
                            lblCFPH3A.Text = CfPh3A;

                            string CfPhSysAString = hexStringArray[342] + hexStringArray[343] + hexStringArray[344];
                            string CfPhSysA = Helper.HexStringToDecimal(CfPhSysAString, hexStringArray[341]);
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
    }
}
