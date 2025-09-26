using System;
using System.IO.Ports;
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

        // Guna 컨트롤
        private Guna2ComboBox comboBox;
        private Guna2ComboBox cmbPort;
        private Guna2Button btnConnect;
        private Guna2Button btnDisconnect;
        private Guna2Button btnSend;
        private Guna2TextBox txtSend;
        private Guna2TextBox txtLog;

        public RS485_protocol()
        {
            // InitializeComponent(); // ← 제거
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.EnforceBackcolorOnAllComponents = true;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Blue600, Primary.Blue700,
                Primary.Blue200, Accent.LightBlue200,
                TextShade.WHITE);

            InitUI();   // 여기서 컨트롤 생성
            LoadPorts();
        }

        private void InitUI()
        {
            this.Text = "RS485 통신 테스트";
            this.Width = 700;
            this.Height = 500;

            // 포트 선택 콤보박스
            cmbPort = new Guna2ComboBox
            {
                Location = new System.Drawing.Point(20, 80),
                Width = 200,
                Height = 40
            };
            this.Controls.Add(cmbPort);

            // 연결 버튼
            btnConnect = new Guna2Button
            {
                Text = "연결",
                Location = new System.Drawing.Point(240, 80),
                Width = 100,
                Height = 40
            };
            btnConnect.Click += BtnConnect_Click;
            this.Controls.Add(btnConnect);

            // 연결 해제 버튼
            btnDisconnect = new Guna2Button
            {
                Text = "끊기",
                Location = new System.Drawing.Point(360, 80),
                Width = 100,
                Height = 40
            };
            btnDisconnect.Click += BtnDisconnect_Click;
            this.Controls.Add(btnDisconnect);

            // 송신 텍스트 박스
            txtSend = new Guna2TextBox
            {
                PlaceholderText = "보낼 데이터 입력",
                Location = new System.Drawing.Point(20, 140),
                Width = 400,
                Height = 40
            };
            this.Controls.Add(txtSend);

            // 송신 버튼
            btnSend = new Guna2Button
            {
                Text = "보내기",
                Location = new System.Drawing.Point(440, 140),
                Width = 100,
                Height = 40
            };
            btnSend.Click += BtnSend_Click;
            this.Controls.Add(btnSend);

            // 로그 텍스트 박스
            txtLog = new Guna2TextBox
            {
                Multiline = true,
                Location = new System.Drawing.Point(20, 200),
                Width = 640,
                Height = 220,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical
            };
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
            try
            {
                if (serialPort == null || !serialPort.IsOpen)
                {
                    serialPort = new SerialPort(cmbPort.SelectedItem.ToString(), 115200, Parity.None, 8, StopBits.One);
                    serialPort.DataReceived += SerialPort_DataReceived;
                    serialPort.Open();
                    Log($"✅ Port {cmbPort.SelectedItem} opened.");
                }
                else
                {
                    Log("⚠ 이미 연결된 상태입니다.");
                }
            }
            catch (Exception ex)
            {
                Log($"❌ 연결 실패: {ex.Message}");
            }
        }

        private void BtnDisconnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort != null && serialPort.IsOpen)
                {
                    serialPort.Close();
                    Log($"❌ Port {serialPort.PortName} closed.");
                }
                else
                {
                    Log("⚠ 이미 끊긴 상태입니다.");
                }
            }
            catch (Exception ex)
            {
                Log($"❌ 연결 해제 실패: {ex.Message}");
            }
        }

        private void BtnSend_Click(object sender, EventArgs e)
        {
            if (serialPort != null && serialPort.IsOpen)
            {
                try
                {
                    byte stx = 0x02;
                    byte etx = 0x03;
                    byte addr = 0x01;
                    byte cmd = 0x10;
                    byte[] data = System.Text.Encoding.ASCII.GetBytes(txtSend.Text.PadRight(10, ' '));

                    // 간단 XOR CRC
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
                    Log("📤 Sent: " + BitConverter.ToString(packet));
                    txtSend.Clear();
                }
                catch (Exception ex)
                {
                    Log($"❌ 전송 실패: {ex.Message}");
                }
            }
            else
            {
                Log("⚠ 포트가 열려있지 않습니다.");
            }
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string data = serialPort.ReadExisting();
                this.Invoke(new Action(() =>
                {
                    Log("📥 Received: " + data);
                }));
            }
            catch (Exception ex)
            {
                this.Invoke(new Action(() =>
                {
                    Log($"❌ 수신 오류: {ex.Message}");
                }));
            }
        }

        private void Log(string message)
        {
            txtLog.AppendText($"{DateTime.Now:HH:mm:ss} {message}\r\n");
        }


    }
}
