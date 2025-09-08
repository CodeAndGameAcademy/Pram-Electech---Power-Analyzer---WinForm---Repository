using PowerAnalyzer.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace PowerAnalyzer
{
    public partial class MainForm : Form
    {
        private SerialPort _serialPort;
        private List<byte> _buffer = new List<byte>();

        private readonly byte[] startMarker = { 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
        private readonly byte[] endMarker = { 0xFE, 0xFE, 0xFE, 0xFE, 0xFE };

        public MainForm()
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
                        int startIndex = FindSequence(_buffer, startMarker);
                        if (startIndex == -1) break;

                        int endIndex = FindSequence(_buffer, endMarker, startIndex + startMarker.Length);
                        if (endIndex == -1) break;

                        int length = endIndex + endMarker.Length - startIndex;
                        byte[] frame = _buffer.Skip(startIndex).Take(length).ToArray();

                        // Remove processed bytes from buffer
                        _buffer.RemoveRange(0, startIndex + length);

                        // Convert to HEX string
                        string hexFrame = BitConverter.ToString(frame).Replace("-", " ");
                        byte[] byteArray = HexStringToByteArray(hexFrame);

                        this.BeginInvoke(new Action(() =>
                        {
                            // Console.WriteLine(hexString);
                            string[] hexStringArray = hexFrame.Split(' ');

                            #region DP                                      
                            //byte v1DP = Convert.ToByte(hexStringArray[224], 16);
                            //byte v2DP = Convert.ToByte(hexStringArray[223], 16);
                            //byte v3DP = Convert.ToByte(hexStringArray[222], 16);

                            //byte i1DP = Convert.ToByte(hexStringArray[229], 16);
                            //byte i2DP = Convert.ToByte(hexStringArray[228], 16);
                            //byte i3DP = Convert.ToByte(hexStringArray[227], 16);

                            //byte p1DP = Convert.ToByte(hexStringArray[234], 16);
                            //byte p2DP = Convert.ToByte(hexStringArray[233], 16);
                            //byte p3DP = Convert.ToByte(hexStringArray[232], 16);


                            string v1DP = hexStringArray[224];
                            string v2DP = hexStringArray[223];
                            string v3DP = hexStringArray[222];

                            string i1DP = hexStringArray[229];
                            string i2DP = hexStringArray[228];
                            string i3DP = hexStringArray[227];

                            string p1DP = hexStringArray[234];
                            string p2DP = hexStringArray[233];
                            string p3DP = hexStringArray[232];

                            string freqDP = "03";
                            string temperatureDP = "02";

                            #endregion

                            // v1, v2, v3
                            string v1String = hexStringArray[6] + hexStringArray[7] + hexStringArray[8] + hexStringArray[9];                            
                            string v1 = HexDecimalConverter.ToConvert(v1String, v1DP);
                            lblV1.Text = v1;

                            string v2String = hexStringArray[11] + hexStringArray[12] + hexStringArray[13] + hexStringArray[14];
                            string v2 = HexDecimalConverter.ToConvert(v2String, v2DP);
                            lblV2.Text = v2;

                            string v3String = hexStringArray[16] + hexStringArray[17] + hexStringArray[18] + hexStringArray[19];
                            string v3 = HexDecimalConverter.ToConvert(v3String, v3DP);
                            lblV3.Text = v3;


                            // i1, i2, i3
                            string i1String = hexStringArray[21] + hexStringArray[22] + hexStringArray[23] + hexStringArray[24];
                            string i1 = HexDecimalConverter.ToConvert(i1String, i1DP);
                            lblI1.Text = i1;
                            
                            string i2String = hexStringArray[26] + hexStringArray[27] + hexStringArray[28] + hexStringArray[29];
                            string i2 = HexDecimalConverter.ToConvert(i2String, i2DP);
                            lblI2.Text = i2;

                            string i3String = hexStringArray[31] + hexStringArray[32] + hexStringArray[33] + hexStringArray[34];
                            string i3 = HexDecimalConverter.ToConvert(i3String, i3DP);
                            lblI3.Text = i3;


                            // p1, p2, p3
                            string p1String = hexStringArray[36] + hexStringArray[37] + hexStringArray[38] + hexStringArray[39];
                            string p1 = HexDecimalConverter.ToConvert(p1String, p1DP);
                            lblP1.Text = p1;

                            string p2String = hexStringArray[41] + hexStringArray[42] + hexStringArray[43] + hexStringArray[44];
                            string p2 = HexDecimalConverter.ToConvert(p2String, p2DP);
                            lblP2.Text = p2;

                            string p3String = hexStringArray[46] + hexStringArray[47] + hexStringArray[48] + hexStringArray[49];
                            string p3 = HexDecimalConverter.ToConvert(p3String, p3DP);
                            lblP3.Text = p3;


                            // pf1, pf2, pf3
                            string pf1String = hexStringArray[51] + hexStringArray[52] + hexStringArray[53] + hexStringArray[54];
                            string pf1 = HexDecimalConverter.ToConvert(pf1String, p1DP);
                            lblPF1.Text = pf1;

                            string pf2String = hexStringArray[56] + hexStringArray[57] + hexStringArray[58] + hexStringArray[59];
                            string pf2 = HexDecimalConverter.ToConvert(pf2String, p2DP);
                            lblPF2.Text = pf2;

                            string pf3String = hexStringArray[61] + hexStringArray[62] + hexStringArray[63] + hexStringArray[64];
                            string pf3 = HexDecimalConverter.ToConvert(pf3String, p3DP);
                            lblPF3.Text = pf3;

                        }));
                    }
                }
            }
            catch (Exception ex)
            {
                this.BeginInvoke(new Action(() =>
                {
                    // txtHex.AppendText("Error: " + ex.Message + Environment.NewLine);
                    Console.WriteLine("Error : "+ex.Message);
                }));
            }
        }

        // Helper method: Find byte sequence in list
        private int FindSequence(List<byte> buffer, byte[] sequence, int startIndex = 0)
        {
            for (int i = startIndex; i <= buffer.Count - sequence.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < sequence.Length; j++)
                {
                    if (buffer[i + j] != sequence[j])
                    {
                        match = false;
                        break;
                    }
                }
                if (match) return i;
            }
            return -1;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_serialPort != null && _serialPort.IsOpen)
                _serialPort.Close();

            base.OnFormClosing(e);
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (_serialPort.IsOpen)
                {
                    _serialPort.Close();
                    Console.WriteLine("Disconnected");
                }

                ClearAllData();
                btnConnect.Enabled = true;
                btnDisconnect.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error disconnecting: " + ex.Message);
            }
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbPort.SelectedItem == null || cmbPort.SelectedItem.ToString() != "COM3")
                {
                    MessageBox.Show("No valid COM port selected!");
                    return;
                }

                if (_serialPort.IsOpen) _serialPort.Close();

                _serialPort.PortName = cmbPort.SelectedItem.ToString();
                _serialPort.Open();

                btnConnect.Enabled = false;
                btnDisconnect.Enabled = true;

                // txtHex.AppendText($"Connected to {_serialPort.PortName}{Environment.NewLine}");
                Console.Write("Connected to "+_serialPort.PortName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error connecting: " + ex.Message);
            }
        }

        public static byte[] HexStringToByteArray(string hex)
        {
            // Remove spaces and dashes
            hex = hex.Replace(" ", "").Replace("-", "");

            // Must be even length
            if (hex.Length % 2 != 0)
                throw new ArgumentException("Invalid hex string length.");

            byte[] bytes = new byte[hex.Length / 2];

            for (int i = 0; i < bytes.Length; i++)
            {
                bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
            }

            return bytes;
        }
    
        private void ClearAllData()
        {
            lblV1.Text = "0.0";
            lblV2.Text = "0.0";
            lblV3.Text = "0.0";

            lblI1.Text = "0.0";
            lblI2.Text = "0.0";
            lblI3.Text = "0.0";

            lblP1.Text = "0.0";
            lblP2.Text = "0.0";
            lblP3.Text = "0.0";

            lblPF1.Text = "0.0";
            lblPF2.Text = "0.0";
            lblPF3.Text = "0.0";
        }
    }
}
