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

        // UI 컨트롤 선언
        private ComboBox cmbPort;
        private Button btnConnect, btnDisconnect, btnRefresh;
        private DataGridView dgvCommands;
        private RichTextBox txtLog;
        private TextBox txtCustomData; // 사용자 입력 텍스트박스
        private Button btnSendCustomData; // 사용자 입력 전송 버튼

        // 프로토콜 상수 정의 (엑셀 시트 기반)
        private const byte STX = 0x22;
        private const byte ETX = 0x33; // 사용자 코드 참고하여 0x33으로 수정
        private const byte ADDR = 0x03;
        private const byte CMD = 0x85;

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
            this.Text = "RS-485 제어 프로그램";
            this.ClientSize = new System.Drawing.Size(800, 600);

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
                Location = new Point(20, 130),
                Width = 760,
                Height = 250,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvCommands.CellContentClick += DgvCommands_CellContentClick;

            // --- 사용자 정의 데이터 입력 UI 추가 ---
            var lblCustomData = new Label { Text = "사용자 정의 데이터 (9-byte hex, 공백으로 구분):", Location = new Point(20, 390), AutoSize = true };
            txtCustomData = new TextBox { Location = new Point(20, 410), Width = 650, Font = new Font("Consolas", 9) };
            btnSendCustomData = new Button { Text = "사용자 전송", Location = new Point(680, 408), Width = 100 };
            btnSendCustomData.Click += BtnSendCustomData_Click;

            // 로그 창
            txtLog = new RichTextBox
            {
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Location = new Point(20, 450),
                Width = 760,
                Height = 130,
                Font = new Font("Consolas", 9)
            };

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
        }

        private void InitializeCommandGrid()
        {
            // 엑셀 시트에 정의된 명령어 목록 (엑셀 시트와 정확히 일치하도록 수정)
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

            // DataGridView 설정
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
            if (cmbPort.Items.Count > 0)
                cmbPort.SelectedIndex = 0;
            else
                cmbPort.Text = "포트 없음";
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
                // 패킷 구조: [STX][Byte][ADDR][CMD][DATA(9)][CRC(2)][ETX]
                // 전체 길이: 1 + 1 + 1 + 1 + 9 + 2 + 1 = 16 bytes
                byte[] packet = new byte[16];
                byte[] data = command.Data;

                // CRC 계산을 위한 데이터 배열 (Byte + ADDR + CMD + DATA) = 12 bytes
                byte[] crcData = new byte[12];
                crcData[0] = 0x10; // Byte 필드
                crcData[1] = ADDR;
                crcData[2] = CMD;
                Buffer.BlockCopy(data, 0, crcData, 3, data.Length);

                // CRC-16/BUYPASS (Table-lookup) 계산
                ushort crcValue = Crc16Buypass.ComputeChecksum(crcData);
                byte crcHigh = (byte)((crcValue >> 8) & 0xFF);
                byte crcLow = (byte)(crcValue & 0xFF);

                // 패킷 조립
                packet[0] = STX;
                packet[1] = 0x10;
                packet[2] = ADDR;
                packet[3] = CMD;
                Buffer.BlockCopy(data, 0, packet, 4, data.Length); // Data (인덱스 4~12)
                packet[13] = crcHigh; // CRC High Byte (인덱스 13)
                packet[14] = crcLow;  // CRC Low Byte (인덱스 14)
                packet[15] = ETX;     // ETX (인덱스 15)

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
                    this.Invoke((MethodInvoker)delegate { Log($"📥 수신: {BitConverter.ToString(buffer).Replace("-", " ")}", Color.DarkGreen); });
                }
            }
            catch (Exception ex) { this.Invoke((MethodInvoker)delegate { Log($"❌ 수신 오류: {ex.Message}", Color.Red); }); }
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

