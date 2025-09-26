using System;
using System.IO.Ports;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;
using Guna.UI2.WinForms;

namespace RS485_WinForms
{
    public partial class RS485_protocol : MaterialForm
    {
        private readonly MaterialSkinManager materialSkinManager;
        private SerialPort serialPort;

        private Guna2ComboBox cmbPort;
        private Guna2Button btnConnect;
        private Guna2Button btnDisconnect;
        private Guna2Button btnSend;
        private Guna2TextBox txtSend;
        private Guna2TextBox txtLog;

        public RS485_protocol()
        {
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.EnforceBackcolorOnAllComponents = true;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Blue600, Primary.Blue700,
                Primary.Blue200, Accent.LightBlue200,
                TextShade.WHITE);

            InitUI();
            LoadPorts();
        }

        private void InitUI()
        {
            this.Text = "RS485 통신 테스트";
            this.Width = 700;
            this.Height = 500;

            cmbPort = new Guna2ComboBox { Location = new System.Drawing.Point(20, 80), Width = 200, Height = 40 };
            this.Controls.Add(cmbPort);

            btnConnect = new Guna2Button { Text = "연결", Location = new System.Drawing.Point(240, 80), Width = 100, Height = 40 };
            btnConnect.Click += BtnConnect_Click;
            this.Controls.Add(btnConnect);

            btnDisconnect = new Guna2Button { Text = "끊기", Location = new System.Drawing.Point(360, 80), Width = 100, Height = 40 };
            btnDisconnect.Click += BtnDisconnect_Click;
            this.Controls.Add(btnDisconnect);

            txtSend = new Guna2TextBox { PlaceholderText = "보낼 데이터 입력", Location = new System.Drawing.Point(20, 140), Width = 400, Height = 40 };
            this.Controls.Add(txtSend);

            btnSend = new Guna2Button { Text = "보내기", Location = new System.Drawing.Point(440, 140), Width = 100, Height = 40 };
            btnSend.Click += BtnSend_Click;
            this.Controls.Add(btnSend);

            txtLog = new Guna2TextBox { Multiline = true, Location = new System.Drawing.Point(20, 200), Width = 640, Height = 220, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
            this.Controls.Add(txtLog);
        }

        private void LoadPorts()
        {
            cmbPort.Items.Clear();
            cmbPort.Items.AddRange(SerialPort.GetPortNames());
            if (cmbPort.Items.Count > 0) cmbPort.SelectedIndex = 0;
        }

        private void BtnConnect_Click(object sender, EventArgs e)
        {
            if (cmbPort.SelectedItem == null) { Log("⚠ 포트를 선택하세요."); return; }

            try
            {
                if (serialPort == null || !serialPort.IsOpen)
                {
                    serialPort = new SerialPort(cmbPort.SelectedItem.ToString(), 115200, Parity.None, 8, StopBits.One);
                    serialPort.DataReceived += SerialPort_DataReceived;
                    serialPort.Open();
                    Log($"✅ Port {cmbPort.SelectedItem} opened.");
                }
                else { Log("⚠ 이미 연결된 상태입니다."); }
            }
            catch (Exception ex) { Log($"❌ 연결 실패: {ex.Message}"); }
        }

        private void BtnDisconnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort != null && serialPort.IsOpen)
                {
                    string portName = serialPort.PortName;
                    serialPort.Close();
                    Log($"❌ Port {portName} closed.");
                }
                else { Log("⚠ 이미 끊긴 상태입니다."); }
            }
            catch (Exception ex) { Log($"❌ 연결 해제 실패: {ex.Message}"); }
        }

        private void BtnSend_Click(object sender, EventArgs e)
        {
            if (serialPort == null || !serialPort.IsOpen)
            {
                Log("⚠ 포트가 열려있지 않습니다.");
                return;
            }

            Task.Run(() =>
            {
                try
                {
                    byte stx = 0x02;
                    byte etx = 0x03;
                    byte addr = 0x01;
                    byte cmd;
                    byte[] data;

                    if (string.IsNullOrWhiteSpace(txtSend.Text))
                    {
                        // 상태 요청
                        cmd = 0x20;
                        data = new byte[0];
                        this.Invoke(() => Log("📤 상태 요청 패킷 전송"));
                    }
                    else
                    {
                        cmd = 0x10;
                        data = System.Text.Encoding.ASCII.GetBytes(txtSend.Text.PadRight(10, ' '));
                    }

                    // XOR CRC
                    byte crc = addr;
                    crc ^= cmd;
                    foreach (byte b in data) crc ^= b;

                    byte[] packet = new byte[6 + data.Length];
                    int idx = 0;
                    packet[idx++] = stx;
                    packet[idx++] = (byte)data.Length;
                    packet[idx++] = addr;
                    packet[idx++] = cmd;
                    Array.Copy(data, 0, packet, idx, data.Length);
                    idx += data.Length;
                    packet[idx++] = crc;
                    packet[idx++] = etx;

                    serialPort.Write(packet, 0, packet.Length);
                    this.Invoke(() => { Log("📤 Sent: " + BitConverter.ToString(packet)); txtSend.Clear(); });
                }
                catch (Exception ex)
                {
                    this.Invoke(() => Log($"❌ 전송 실패: {ex.Message}"));
                }
            });
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                int bytesToRead = serialPort.BytesToRead;
                if (bytesToRead > 0)
                {
                    byte[] buffer = new byte[bytesToRead];
                    serialPort.Read(buffer, 0, bytesToRead);

                    this.Invoke(() =>
                    {
                        Log("📥 Received: " + BitConverter.ToString(buffer));

                        // 상태 요청 응답 분석 예시
                        if (buffer.Length >= 6 && buffer[3] == 0x20) // cmd가 상태 응답
                        {
                            // 예시: 장치가 0x01이면 정상, 0x00이면 오류
                            string status = buffer.Length > 4 && buffer[4] == 0x01 ? "✅ 장치 정상" : "❌ 장치 오류";
                            Log(status);
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                this.Invoke(() => Log($"❌ 수신 오류: {ex.Message}"));
            }
        }

        private void Log(string message)
        {
            txtLog.AppendText($"{DateTime.Now:HH:mm:ss} {message}\r\n");
        }
    }
}
