using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;

namespace RS485_WinForms_Improved
{
    // 엑셀 시트의 한 줄에 해당하는 명령 데이터를 저장하기 위한 클래스
    public class CommandData
    {
        public string Description { get; set; } // 설명 (예: 이젝터 1 진공)
        public byte[] Data { get; set; } // 실제 전송될 데이터 부분 (9-byte)
    }

    /// <summary>
    /// CRC-16/BUYPASS 계산을 위한 정적 클래스 (조회 테이블 방식)
    /// </summary>
    public static class Crc16Buypass
    {
        private const ushort Polynomial = 0x8005;
        private static readonly ushort[] Table = new ushort[256];

        static Crc16Buypass()
        {
            for (ushort i = 0; i < 256; ++i)
            {
                ushort value = 0;
                ushort temp = (ushort)(i << 8);
                for (byte j = 0; j < 8; ++j)
                {
                    if (((value ^ temp) & 0x8000) != 0)
                    {
                        value = (ushort)((value << 1) ^ Polynomial);
                    }
                    else
                    {
                        value <<= 1;
                    }
                    temp <<= 1;
                }
                Table[i] = value;
            }
        }

        public static ushort ComputeChecksum(byte[] bytes)
        {
            ushort crc = 0;
            foreach (byte b in bytes)
            {
                crc = (ushort)((crc << 8) ^ Table[((crc >> 8) ^ b) & 0xFF]);
            }
            return crc;
        }
    }


    public partial class RS485_ImprovedForm : MaterialForm
    {
        private readonly MaterialSkinManager materialSkinManager;
        private SerialPort serialPort;

        // --- 수신 데이터 조각을 모으기 위한 버퍼 ---
        private List<byte> receiveBuffer = new List<byte>();

        // UI 컨트롤 선언
        private ComboBox cmbPort;
        private Button btnConnect, btnDisconnect, btnRefresh;
        private DataGridView dgvCommands;
        private RichTextBox txtLog;
        private TextBox txtCustomData;
        private Button btnSendCustomData;

        // --- 상태 분석 UI 컨트롤 ---
        private Label[] lblDataBytes = new Label[9];
        private Label[] lblDataHex = new Label[9];
        private Label[] lblDataBin = new Label[9];

        // 프로토콜 상수 정의 (명령 전송용)
        private const byte SEND_STX = 0x22;
        private const byte SEND_ETX = 0x33;
        private const byte SEND_ADDR = 0x03;
        private const byte SEND_CMD = 0x85;

        // 프로토콜 상수 정의 (상태 수신용)
        private const byte RECV_STX = 0x44;
        private const int PACKET_LENGTH = 16;


        public RS485_ImprovedForm()
        {
            InitializeComponent();

            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.EnforceBackcolorOnAllComponents = true;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Indigo500, Primary.Indigo700,
                Primary.Indigo100, Accent.Pink200,
                TextShade.WHITE);

            LoadPorts();
            InitializeCommandGrid();
        }

        private void InitializeComponent()
        {
            this.Text = "RS-485 제어 및 프로토콜 분석 도구";
            this.ClientSize = new System.Drawing.Size(900, 640);

            // 포트 선택
            var lblPort = new Label { Text = "COM 포트:", Location = new Point(20, 80), AutoSize = true };
            cmbPort = new ComboBox { Location = new Point(100, 78), Width = 120 };
            btnRefresh = new Button { Text = "새로고침", Location = new Point(230, 76), Width = 80 };
            btnRefresh.Click += (s, e) => LoadPorts();

            // 연결/해제 버튼
            btnConnect = new Button { Text = "연결", Location = new Point(320, 76), Width = 100 };
            btnConnect.Click += BtnConnect_Click;
            btnDisconnect = new Button { Text = "해제", Location = new Point(430, 76), Width = 100 };
            btnDisconnect.Click += BtnDisconnect_Click;

            // 명령어 그리드
            dgvCommands = new DataGridView
            {
                Location = new Point(20, 120),
                Width = 550,
                Height = 300,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvCommands.CellContentClick += DgvCommands_CellContentClick;

            // --- 사용자 정의 데이터 입력 UI ---
            var lblCustomData = new Label { Text = "사용자 정의 데이터 (9-byte hex, 공백으로 구분):", Location = new Point(20, 430), AutoSize = true };
            txtCustomData = new TextBox { Location = new Point(20, 450), Width = 440, Font = new Font("Consolas", 9) };
            btnSendCustomData = new Button { Text = "사용자 전송", Location = new Point(470, 448), Width = 100 };
            btnSendCustomData.Click += BtnSendCustomData_Click;

            // 로그 창
            txtLog = new RichTextBox
            {
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Location = new Point(20, 490),
                Width = 860,
                Height = 330,
                Font = new Font("Consolas", 9)
            };

            // --- 실시간 상태 분석 UI ---
            var analysisGroup = new GroupBox { Text = "실시간 상태 분석 (수신 DATA)", Location = new Point(580, 120), Width = 300, Height = 360 };
            var fontBold = new Font("Consolas", 10, FontStyle.Bold);
            var fontRegular = new Font("Consolas", 10);

            for (int i = 0; i < 9; i++)
            {
                int yPos = 30 + (i * 35);
                lblDataBytes[i] = new Label { Text = $"Data[{i}]:", Location = new Point(15, yPos), Font = fontBold, AutoSize = true };
                lblDataHex[i] = new Label { Text = "00", Location = new Point(100, yPos), Font = fontRegular, ForeColor = Color.Blue, AutoSize = true };
                lblDataBin[i] = new Label { Text = "00000000", Location = new Point(150, yPos), Font = fontRegular, ForeColor = Color.DarkGreen, AutoSize = true };

                analysisGroup.Controls.Add(lblDataBytes[i]);
                analysisGroup.Controls.Add(lblDataHex[i]);
                analysisGroup.Controls.Add(lblDataBin[i]);
            }

            // 컨트롤들을 폼에 추가
            this.Controls.Add(lblPort);
            this.Controls.Add(cmbPort);
            this.Controls.Add(btnRefresh);
            this.Controls.Add(btnConnect);
            this.Controls.Add(btnDisconnect);
            this.Controls.Add(dgvCommands);
            this.Controls.Add(lblCustomData);
            this.Controls.Add(txtCustomData);
            this.Controls.Add(btnSendCustomData);
            this.Controls.Add(txtLog);
            this.Controls.Add(analysisGroup);
        }

        private void InitializeCommandGrid()
        {
            // '쓰기'를 위한 명령어 목록 (STX=0x22 프로토콜)
            var commands = new List<CommandData>
            {
                new CommandData { Description = "이젝터 1 진공", Data = new byte[] { 0, 0x01, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 1 파기", Data = new byte[] { 0, 0x02, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 2 진공", Data = new byte[] { 0, 0x04, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 2 파기", Data = new byte[] { 0, 0x08, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3 진공", Data = new byte[] { 0, 0x10, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3 파기", Data = new byte[] { 0, 0x20, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 4 진공", Data = new byte[] { 0, 0x40, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 4 파기", Data = new byte[] { 0, 0x80, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5 진공", Data = new byte[] { 0, 0, 0x01, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5 파기", Data = new byte[] { 0, 0, 0x02, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 6 진공", Data = new byte[] { 0, 0, 0x04, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 6 파기", Data = new byte[] { 0, 0, 0x08, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 1 후진", Data = new byte[] { 0, 0, 0x10, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 1 전진", Data = new byte[] { 0, 0, 0x20, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 2 후진", Data = new byte[] { 0, 0, 0x40, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 2 전진", Data = new byte[] { 0, 0, 0x80, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 3 후진", Data = new byte[] { 0, 0, 0, 0x01, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 3 전진", Data = new byte[] { 0, 0, 0, 0x02, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 4 후진", Data = new byte[] { 0, 0, 0, 0x04, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 4 전진", Data = new byte[] { 0, 0, 0, 0x08, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 5 UP",   Data = new byte[] { 0, 0, 0, 0x10, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 5 DOWN", Data = new byte[] { 0, 0, 0, 0x20, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 6 후진", Data = new byte[] { 0, 0, 0, 0x40, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 6 전진", Data = new byte[] { 0, 0, 0, 0x80, 0, 0, 0, 0, 0 } },
            };

            dgvCommands.DataSource = commands;
            dgvCommands.Columns["Description"].HeaderText = "명령 설명";
            dgvCommands.Columns["Data"].Visible = false;

            var sendButtonColumn = new DataGridViewButtonColumn
            {
                Name = "SendButton",
                HeaderText = "동작",
                Text = "전송",
                UseColumnTextForButtonValue = true
            };
            dgvCommands.Columns.Add(sendButtonColumn);
            dgvCommands.Columns["SendButton"].Width = 80;
        }

        private void LoadPorts()
        {
            cmbPort.Items.Clear();
            cmbPort.Items.AddRange(SerialPort.GetPortNames());
            if (cmbPort.Items.Count > 0) cmbPort.SelectedIndex = 0;
            else cmbPort.Text = "포트 없음";
        }

        private void BtnConnect_Click(object sender, EventArgs e)
        {
            if (cmbPort.SelectedItem == null) { Log("포트를 선택하세요.", Color.Red); return; }
            try
            {
                if (serialPort != null && serialPort.IsOpen) { Log("이미 연결되어 있습니다.", Color.Orange); return; }
                serialPort = new SerialPort(cmbPort.SelectedItem.ToString(), 115200, Parity.None, 8, StopBits.One);
                serialPort.DataReceived += SerialPort_DataReceived;
                serialPort.Open();
                Log($"✅ {serialPort.PortName} 포트가 연결되었습니다.", Color.Green);
            }
            catch (Exception ex) { Log($"❌ 연결 실패: {ex.Message}", Color.Red); }
        }

        private void BtnDisconnect_Click(object sender, EventArgs e)
        {
            if (serialPort == null || !serialPort.IsOpen) { Log("연결 상태가 아닙니다.", Color.Orange); return; }
            try
            {
                serialPort.Close();
                Log($"🔌 {serialPort.PortName} 포트 연결이 해제되었습니다.", Color.Black);
            }
            catch (Exception ex) { Log($"❌ 해제 실패: {ex.Message}", Color.Red); }
        }

        private void DgvCommands_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dgvCommands.Columns["SendButton"].Index && e.RowIndex >= 0)
            {
                var selectedCommand = (dgvCommands.Rows[e.RowIndex].DataBoundItem as CommandData);
                if (selectedCommand != null) { SendPacket(selectedCommand); }
            }
        }

        private void BtnSendCustomData_Click(object sender, EventArgs e)
        {
            string inputText = txtCustomData.Text.Trim();
            if (string.IsNullOrEmpty(inputText)) { Log("⚠ 전송할 데이터를 입력하세요.", Color.Orange); return; }
            string[] hexValues = inputText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (hexValues.Length != 9) { Log($"❌ 데이터는 반드시 9-byte여야 합니다. (입력된 바이트: {hexValues.Length})", Color.Red); return; }
            try
            {
                byte[] dataBytes = hexValues.Select(hex => Convert.ToByte(hex, 16)).ToArray();
                var customCommand = new CommandData { Description = "사용자 정의 데이터", Data = dataBytes };
                SendPacket(customCommand);
            }
            catch (FormatException) { Log("❌ 잘못된 Hex 값 형식입니다. (예: 01 23 AB)", Color.Red); }
            catch (Exception ex) { Log($"❌ 데이터 변환 오류: {ex.Message}", Color.Red); }
        }

        private void SendPacket(CommandData command)
        {
            if (serialPort == null || !serialPort.IsOpen) { Log("⚠ 포트가 열려있지 않습니다.", Color.Red); return; }
            try
            {
                byte[] packet = new byte[16];
                byte[] data = command.Data;

                byte[] crcData = new byte[12];
                crcData[0] = 0x10;
                crcData[1] = SEND_ADDR;
                crcData[2] = SEND_CMD;
                Buffer.BlockCopy(data, 0, crcData, 3, data.Length);

                ushort crcValue = Crc16Buypass.ComputeChecksum(crcData);
                byte crcHigh = (byte)((crcValue >> 8) & 0xFF);
                byte crcLow = (byte)(crcValue & 0xFF);

                packet[0] = SEND_STX;
                packet[1] = 0x10;
                packet[2] = SEND_ADDR;
                packet[3] = SEND_CMD;
                Buffer.BlockCopy(data, 0, packet, 4, data.Length);
                packet[13] = crcHigh;
                packet[14] = crcLow;
                packet[15] = SEND_ETX;

                serialPort.Write(packet, 0, packet.Length);
                Log($"📤 [{command.Description}] 전송: {BitConverter.ToString(packet).Replace("-", " ")}", Color.Blue);
            }
            catch (Exception ex) { Log($"❌ 전송 실패: {ex.Message}", Color.Red); }
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

                    receiveBuffer.AddRange(buffer);

                    this.Invoke((MethodInvoker)delegate {
                        while (receiveBuffer.Count >= PACKET_LENGTH)
                        {
                            int packetStartIndex = receiveBuffer.FindIndex(b => b == RECV_STX);
                            if (packetStartIndex == -1) { receiveBuffer.Clear(); return; }
                            if (packetStartIndex > 0) { receiveBuffer.RemoveRange(0, packetStartIndex); }
                            if (receiveBuffer.Count < PACKET_LENGTH) { break; }

                            byte[] completePacket = receiveBuffer.GetRange(0, PACKET_LENGTH).ToArray();
                            Log($"📥 수신: {BitConverter.ToString(completePacket).Replace("-", " ")}", Color.DarkGreen);

                            ParseReceivedData(completePacket);

                            receiveBuffer.RemoveRange(0, PACKET_LENGTH);
                        }
                    });
                }
            }
            catch (Exception ex) { this.Invoke((MethodInvoker)delegate { Log($"❌ 수신 오류: {ex.Message}", Color.Red); }); }
        }

        /// <summary>
        /// 수신된 데이터 패킷을 분석하여 UI에 Hex와 Binary 값으로 표시합니다.
        /// </summary>
        private void ParseReceivedData(byte[] packet)
        {
            byte[] data = new byte[9];
            Buffer.BlockCopy(packet, 4, data, 0, 9);

            for (int i = 0; i < 9; i++)
            {
                lblDataHex[i].Text = data[i].ToString("X2");
                lblDataBin[i].Text = Convert.ToString(data[i], 2).PadLeft(8, '0');
            }
        }

        private void Log(string message, Color color)
        {
            if (txtLog.InvokeRequired) { txtLog.Invoke((MethodInvoker)delegate { Log(message, color); }); }
            else
            {
                txtLog.SelectionStart = txtLog.TextLength;
                txtLog.SelectionLength = 0;
                txtLog.SelectionColor = color;
                txtLog.AppendText($"{DateTime.Now:HH:mm:ss} - {message}\r\n");
                txtLog.SelectionColor = txtLog.ForeColor;
                txtLog.ScrollToCaret();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (serialPort != null && serialPort.IsOpen) { serialPort.Close(); }
            base.OnFormClosing(e);
        }
    }
}

