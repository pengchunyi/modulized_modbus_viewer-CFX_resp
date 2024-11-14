// ModbusViewer.cs

using System;
using System.IO.Ports;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AmqpModbusIntegration
{
    public class ModbusViewer : Form
    {
        private SerialPort serialPort;
        private ComboBox portSelector;
        private TextBox stationNumberTextBox;
        private Button connectButton, readButton;
        private byte stationNumber = 0xFF;

        public int currentStatus1, currentStatus2, leakageCurrent, tempA, tempB, tempC, tempN;
        public int voltageA, voltageB, voltageC, currentA, currentB, currentC;
        public int powerFactorA, powerFactorB, powerFactorC, activePowerA, activePowerB, activePowerC;
        public int reactivePowerA, reactivePowerB, reactivePowerC, breakerTimes, energyHighByte, energyLowByte;
        public int switchStatus, apparentPowerA, apparentPowerB, apparentPowerC, totalApparentPower, totalActivePower, totalReactivePower;
        public int combinedPowerFactor, lineFrequency, deviceType, historicalLeakage, historicalCurrentA, historicalCurrentB, historicalCurrentC;

        public ModbusViewer()
        {
            InitializeUI();
        }

        private void InitializeUI()
        {
            Text = "Modbus Control";
            Width = 400;
            Height = 300;

            var portLabel = new Label { Text = "COM口:", Location = new System.Drawing.Point(10, 10), AutoSize = true };
            Controls.Add(portLabel);

            portSelector = new ComboBox { Location = new System.Drawing.Point(70, 10), Width = 100 };
            portSelector.Items.AddRange(SerialPort.GetPortNames());
            Controls.Add(portSelector);

            var stationNumberLabel = new Label { Text = "站號:", Location = new System.Drawing.Point(10, 50), AutoSize = true };
            Controls.Add(stationNumberLabel);

            stationNumberTextBox = new TextBox { Location = new System.Drawing.Point(70, 50), Width = 100 };
            Controls.Add(stationNumberTextBox);

            connectButton = new Button { Text = "連接", Location = new System.Drawing.Point(200, 10), Width = 80 };
            connectButton.Click += (s, e) => InitializeSerialPort(portSelector.SelectedItem?.ToString());
            Controls.Add(connectButton);

            readButton = new Button { Text = "讀取數值", Location = new System.Drawing.Point(200, 50), Width = 80 };
            readButton.Click += (s, e) => ReadAllParameters();
            Controls.Add(readButton);
        }

        private void InitializeSerialPort(string portName)
        {
            if (serialPort != null && serialPort.IsOpen)
            {
                serialPort.Close();
                serialPort.Dispose();
            }

            if (!byte.TryParse(stationNumberTextBox.Text, out stationNumber))
            {
                MessageBox.Show("請輸入有效的站號 (0-255)");
                return;
            }

            serialPort = new SerialPort(portName, 9600, Parity.None, 8, StopBits.One);
            try
            {
                serialPort.Open();
                MessageBox.Show($"串口 {portName} 連接成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"串口連接失敗: {ex.Message}");
            }
        }

        public void ReadAllParameters()
        {
            if (serialPort != null && serialPort.IsOpen)
            {
                byte[] readCommand = { stationNumber, 0x03, 0x00, 0x00, 0x00, 0x30 };
                ushort crc = CalculateCRC(readCommand);
                byte[] crcBytes = BitConverter.GetBytes(crc);
                byte[] fullCommand = new byte[readCommand.Length + 2];
                Array.Copy(readCommand, fullCommand, readCommand.Length);
                fullCommand[fullCommand.Length - 2] = crcBytes[0];
                fullCommand[fullCommand.Length - 1] = crcBytes[1];

                serialPort.Write(fullCommand, 0, fullCommand.Length);

                Task.Run(() =>
                {
                    try
                    {
                        byte[] buffer = new byte[256];
                        int bytesRead = serialPort.Read(buffer, 0, buffer.Length);

                        if (bytesRead > 5)
                        {
                            currentStatus1 = (buffer[3] << 8) | buffer[4];
                            currentStatus2 = (buffer[5] << 8) | buffer[6];
                            leakageCurrent = (buffer[7] << 8) | buffer[8];
                            tempA = (buffer[9] << 8) | buffer[10];
                            tempB = (buffer[11] << 8) | buffer[12];
                            tempC = (buffer[13] << 8) | buffer[14];
                            tempN = (buffer[15] << 8) | buffer[16];
                            voltageA = (buffer[17] << 8) | buffer[18];
                            voltageB = (buffer[19] << 8) | buffer[20];
                            voltageC = (buffer[21] << 8) | buffer[22];
                            currentA = (buffer[23] << 8) | buffer[24];
                            currentB = (buffer[25] << 8) | buffer[26];
                            currentC = (buffer[27] << 8) | buffer[28];
                            powerFactorA = (buffer[29] << 8) | buffer[30];
                            powerFactorB = (buffer[31] << 8) | buffer[32];
                            powerFactorC = (buffer[33] << 8) | buffer[34];
                            activePowerA = (buffer[35] << 8) | buffer[36];
                            activePowerB = (buffer[37] << 8) | buffer[38];
                            activePowerC = (buffer[39] << 8) | buffer[40];
                            reactivePowerA = (buffer[41] << 8) | buffer[42];
                            reactivePowerB = (buffer[43] << 8) | buffer[44];
                            reactivePowerC = (buffer[45] << 8) | buffer[46];
                            breakerTimes = (buffer[47] << 8) | buffer[48];
                            energyHighByte = (buffer[49] << 8) | buffer[50];
                            energyLowByte = (buffer[51] << 8) | buffer[52];
                            switchStatus = (buffer[53] << 8) | buffer[54];
                            apparentPowerA = (buffer[55] << 8) | buffer[56];
                            apparentPowerB = (buffer[57] << 8) | buffer[58];
                            apparentPowerC = (buffer[59] << 8) | buffer[60];
                            totalApparentPower = (buffer[61] << 8) | buffer[62];
                            totalActivePower = (buffer[63] << 8) | buffer[64];
                            totalReactivePower = (buffer[65] << 8) | buffer[66];
                            combinedPowerFactor = (buffer[67] << 8) | buffer[68];
                            lineFrequency = (buffer[69] << 8) | buffer[70];
                            deviceType = (buffer[71] << 8) | buffer[72];
                            historicalLeakage = (buffer[73] << 8) | buffer[74];
                            historicalCurrentA = (buffer[75] << 8) | buffer[76];
                            historicalCurrentB = (buffer[77] << 8) | buffer[78];
                            historicalCurrentC = (buffer[79] << 8) | buffer[80];
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"讀取參數時發生錯誤: {ex.Message}");
                    }
                });
            }
        }

        private ushort CalculateCRC(byte[] data)
        {
            ushort crc = 0xFFFF;
            for (int pos = 0; pos < data.Length; pos++)
            {
                crc ^= (ushort)data[pos];
                for (int i = 8; i != 0; i--)
                {
                    if ((crc & 0x0001) != 0)
                    {
                        crc >>= 1;
                        crc ^= 0xA001;
                    }
                    else
                    {
                        crc >>= 1;
                    }
                }
            }
            return crc;
        }
    }
}
