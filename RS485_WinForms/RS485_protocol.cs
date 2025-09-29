using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        private System.Windows.Forms.Timer pollingTimer; // 상태 요청을 위한 폴링 타이머

        // --- 자동 모드 상태 관리를 위한 변수 ---
        private enum AutoModeState { Idle, VacuumsOn, ReadyForBreak, Breaking }
        private AutoModeState currentAutoModeState1 = AutoModeState.Idle;
        private AutoModeState currentAutoModeState2 = AutoModeState.Idle;
        private volatile bool isAutoModeRunning = false;

        // --- 자동 모드 트리거 신호 정의 ---
        private readonly byte[] triggerSignal1 = { 0x44, 0x10, 0x03, 0x85, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFA, 0x12, 0x00, 0x00, 0x1B, 0x99, 0x55 };

        // --- 수신 데이터 조각을 모으기 위한 버퍼 ---
        private List<byte> receiveBuffer = new List<byte>();

        // UI 컨트롤 선언
        private ComboBox cmbPort;
        private Button btnConnect, btnDisconnect, btnRefresh;
        private DataGridView dgvCommands;
        private RichTextBox txtLog;
        private TextBox txtCustomData;
        private Button btnSendCustomData;
        private Button btnStartAutoMode, btnStopAutoMode;
        private GroupBox sendGroup;
        private Label lblLastPacket;

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
            InitializePollingTimer(); // 폴링 타이머 초기화
        }

        private void InitializePollingTimer()
        {
            pollingTimer = new System.Windows.Forms.Timer();
            pollingTimer.Interval = 100;
            pollingTimer.Tick += PollingTimer_Tick;
        }

        private void PollingTimer_Tick(object sender, EventArgs e)
        {
            if (isAutoModeRunning && serialPort != null && serialPort.IsOpen)
            {
                CommandData commandToSend = null;

                if (currentAutoModeState1 != AutoModeState.Idle)
                {
                    if (currentAutoModeState1 == AutoModeState.VacuumsOn || currentAutoModeState1 == AutoModeState.ReadyForBreak)
                    {
                        commandToSend = new CommandData { Description = "상태 유지 1 (진공 ON)", Data = new byte[] { 0, 0x50, 0, 0, 0, 0, 0, 0, 0 } };
                    }
                }
                else if (currentAutoModeState2 != AutoModeState.Idle)
                {
                    if (currentAutoModeState2 == AutoModeState.VacuumsOn || currentAutoModeState2 == AutoModeState.ReadyForBreak)
                    {
                        commandToSend = new CommandData { Description = "상태 유지 2 (진공 ON)", Data = new byte[] { 0, 0, 0x05, 0, 0, 0, 0, 0, 0 } };
                    }
                }

                // 어떤 자동 모드든 Idle 상태이거나, 실행 중인 모드가 아니면 기본 상태 요청
                if (commandToSend == null)
                {
                    commandToSend = new CommandData { Description = "상태 요청 (Polling)", Data = new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 } };
                }

                SendPacket(commandToSend);
            }
        }

        private void InitializeComponent()
        {
            this.Text = "RS-485 자동화 제어 프로그램";
            this.ClientSize = new System.Drawing.Size(900, 850);

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
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvCommands.CellContentClick += DgvCommands_CellContentClick;

            // --- 자동 모드 UI (통합) ---
            var autoModeGroup = new GroupBox { Text = "자동 모드 (핸드툴 1 & 2)", Location = new Point(20, 430), Width = 550, Height = 55 };
            btnStartAutoMode = new Button { Text = "시작", Location = new Point(15, 20), Width = 150 };
            btnStartAutoMode.Click += BtnStartAutoMode_Click;
            btnStopAutoMode = new Button { Text = "중지", Location = new Point(180, 20), Width = 150, Enabled = false };
            btnStopAutoMode.Click += BtnStopAutoMode_Click;
            autoModeGroup.Controls.Add(btnStartAutoMode);
            autoModeGroup.Controls.Add(btnStopAutoMode);

            // --- 추가 명령 전송 UI ---
            sendGroup = new GroupBox { Text = "추가 명령 전송", Location = new Point(20, 490), Width = 550, Height = 100 };
            var lblCustomData = new Label { Text = "전체 패킷 직접 전송 (16-byte hex, 공백으로 구분):", Location = new Point(15, 25), AutoSize = true };
            txtCustomData = new TextBox { Location = new Point(15, 45), Width = 400, Font = new Font("Consolas", 9) };
            btnSendCustomData = new Button { Text = "사용자 전송", Location = new Point(425, 43), Width = 100 };
            btnSendCustomData.Click += BtnSendCustomData_Click;
            sendGroup.Controls.Add(lblCustomData);
            sendGroup.Controls.Add(txtCustomData);
            sendGroup.Controls.Add(btnSendCustomData);

            // 로그 창
            txtLog = new RichTextBox
            {
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Location = new Point(20, 600),
                Width = 860,
                Height = 230,
                Font = new Font("Consolas", 9)
            };

            // --- 실시간 상태 분석 UI ---
            var analysisGroup = new GroupBox { Text = "실시간 상태 분석", Location = new Point(580, 120), Width = 300, Height = 470 };
            var fontBold = new Font("Consolas", 10, FontStyle.Bold);
            var fontRegular = new Font("Consolas", 10);
            analysisGroup.Controls.Add(new Label { Text = "최종 수신 패킷:", Location = new Point(15, 30), Font = fontBold, AutoSize = true });
            lblLastPacket = new Label { Text = "대기 중...", Location = new Point(15, 55), Font = new Font("Consolas", 9, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Width = 270, Height = 40 };
            analysisGroup.Controls.Add(new Label { Text = "수신 DATA (Hex / Binary):", Location = new Point(15, 110), Font = fontBold, AutoSize = true });
            for (int i = 0; i < 9; i++)
            {
                int yPos = 140 + (i * 28);
                lblDataBytes[i] = new Label { Text = $"Data[{i}]:", Location = new Point(15, yPos), Font = fontBold, AutoSize = true };
                lblDataHex[i] = new Label { Text = "00", Location = new Point(100, yPos), Font = fontRegular, ForeColor = Color.Blue, AutoSize = true };
                lblDataBin[i] = new Label { Text = "00000000", Location = new Point(150, yPos), Font = fontRegular, ForeColor = Color.DarkGreen, AutoSize = true };
                analysisGroup.Controls.Add(lblDataBytes[i]);
                analysisGroup.Controls.Add(lblDataHex[i]);
                analysisGroup.Controls.Add(lblDataBin[i]);
            }

            this.Controls.Add(lblPort); this.Controls.Add(cmbPort); this.Controls.Add(btnRefresh);
            this.Controls.Add(btnConnect); this.Controls.Add(btnDisconnect); this.Controls.Add(dgvCommands);
            this.Controls.Add(autoModeGroup); this.Controls.Add(sendGroup);
            this.Controls.Add(txtLog); this.Controls.Add(analysisGroup);
        }

        private void InitializeCommandGrid()
        {
            var commands = new List<CommandData>
            {
                new CommandData { Description = "🚨 전체 정지", Data = new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 1 진공", Data = new byte[] { 0, 0x01, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 1 파기", Data = new byte[] { 0, 0x02, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 2 진공", Data = new byte[] { 0, 0x04, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 2 파기", Data = new byte[] { 0, 0x08, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3 진공", Data = new byte[] { 0, 0x10, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3 파기", Data = new byte[] { 0, 0x20, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 4 진공", Data = new byte[] { 0, 0x40, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 4 파기", Data = new byte[] { 0, 0x80, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3+4 동시 진공", Data = new byte[] { 0, 0x50, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3+4 동시 파기", Data = new byte[] { 0, 0xA0, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5 진공", Data = new byte[] { 0, 0, 0x01, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5 파기", Data = new byte[] { 0, 0, 0x02, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 6 진공", Data = new byte[] { 0, 0, 0x04, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 6 파기", Data = new byte[] { 0, 0, 0x08, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5+6 동시 진공", Data = new byte[] { 0, 0, 0x05, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5+6 동시 파기", Data = new byte[] { 0, 0, 0x0A, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 1 후진", Data = new byte[] { 0, 0, 0x10, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 1 전진", Data = new byte[] { 0, 0, 0x20, 0, 0, 0, 0, 0, 0 } },
            };
            dgvCommands.DataSource = commands;
            dgvCommands.Columns["Description"].HeaderText = "명령 설명";
            dgvCommands.Columns["Data"].Visible = false;
            var sendButtonColumn = new DataGridViewButtonColumn { Name = "SendButton", HeaderText = "동작", Text = "전송", UseColumnTextForButtonValue = true };
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
                lblLastPacket.Text = "신호 대기 중...";
            }
            catch (Exception ex) { Log($"❌ 연결 실패: {ex.Message}", Color.Red); }
        }

        private void BtnDisconnect_Click(object sender, EventArgs e)
        {
            if (serialPort == null || !serialPort.IsOpen) { Log("연결 상태가 아닙니다.", Color.Orange); return; }
            try
            {
                if (isAutoModeRunning) BtnStopAutoMode_Click(sender, e);
                serialPort.Close();
                Log($"🔌 {serialPort.PortName} 포트가 해제되었습니다.", Color.Black);
                lblLastPacket.Text = "연결 해제됨";
            }
            catch (Exception ex) { Log($"❌ 해제 실패: {ex.Message}", Color.Red); }
        }

        private void BtnStartAutoMode_Click(object sender, EventArgs e)
        {
            if (serialPort == null || !serialPort.IsOpen) { Log("⚠ 자동 모드를 시작하려면 먼저 포트를 연결하세요.", Color.Red); return; }
            isAutoModeRunning = true;
            currentAutoModeState1 = AutoModeState.Idle;
            currentAutoModeState2 = AutoModeState.Idle;

            btnStartAutoMode.Enabled = false;
            btnStopAutoMode.Enabled = true;
            dgvCommands.Enabled = false;
            sendGroup.Enabled = false;

            pollingTimer.Start();
            Log("▶️ 통합 자동 모드 시작", Color.Purple);
        }

        private void BtnStopAutoMode_Click(object sender, EventArgs e)
        {
            isAutoModeRunning = false;
            pollingTimer.Stop();

            btnStartAutoMode.Enabled = true;
            btnStopAutoMode.Enabled = false;
            dgvCommands.Enabled = true;
            sendGroup.Enabled = true;

            Log("⏹️ 통합 자동 모드 중지", Color.Purple);
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
            if (hexValues.Length != 16) { Log($"❌ 데이터는 반드시 16-byte여야 합니다. (입력된 바이트: {hexValues.Length})", Color.Red); return; }
            try
            {
                byte[] fullPacket = hexValues.Select(hex => Convert.ToByte(hex, 16)).ToArray();
                if (serialPort != null && serialPort.IsOpen)
                {
                    serialPort.Write(fullPacket, 0, fullPacket.Length);
                    Log($"📤 [사용자 정의 패킷] 전송: {BitConverter.ToString(fullPacket).Replace("-", " ")}", Color.Blue);
                }
                else { Log("⚠ 포트가 열려있지 않습니다.", Color.Red); }
            }
            catch (FormatException) { Log("❌ 잘못된 Hex 값 형식입니다.", Color.Red); }
            catch (Exception ex) { Log($"❌ 데이터 변환 오류: {ex.Message}", Color.Red); }
        }

        private void SendPacket(CommandData command)
        {
            if (serialPort == null || !serialPort.IsOpen) { return; }
            try
            {
                byte[] packet = new byte[16];
                byte[] data = command.Data;
                byte[] crcData = new byte[12];
                crcData[0] = 0x10; crcData[1] = SEND_ADDR; crcData[2] = SEND_CMD;
                Buffer.BlockCopy(data, 0, crcData, 3, data.Length);
                ushort crcValue = Crc16Buypass.ComputeChecksum(crcData);
                packet[0] = SEND_STX; packet[1] = 0x10; packet[2] = SEND_ADDR; packet[3] = SEND_CMD;
                Buffer.BlockCopy(data, 0, packet, 4, data.Length);
                packet[13] = (byte)((crcValue >> 8) & 0xFF); packet[14] = (byte)(crcValue & 0xFF);
                packet[15] = SEND_ETX;
                serialPort.Write(packet, 0, packet.Length);
                if (command.Description.Contains("Polling") == false && command.Description.Contains("유지") == false)
                {
                    Log($"📤 [{command.Description}] 전송: {BitConverter.ToString(packet).Replace("-", " ")}", Color.Blue);
                }
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
                            string packetString = BitConverter.ToString(completePacket).Replace("-", " ");
                            Log($"📥 수신: {packetString}", Color.DarkGreen);
                            UpdateAnalysisUI(completePacket);
                            UpdateLastPacketLabel(packetString);
                            if (isAutoModeRunning)
                            {
                                HandleAutoMode1(completePacket);
                                HandleAutoMode2(completePacket);
                            }
                            receiveBuffer.RemoveRange(0, PACKET_LENGTH);
                        }
                    });
                }
            }
            catch (Exception ex) { this.Invoke((MethodInvoker)delegate { Log($"❌ 수신 오류: {ex.Message}", Color.Red); }); }
        }

        private void HandleAutoMode1(byte[] currentPacket)
        {
            // 다른 모드가 동작 중일 때는 실행하지 않음
            if (currentAutoModeState2 != AutoModeState.Idle) return;

            switch (currentAutoModeState1)
            {
                case AutoModeState.Idle:
                    if (currentPacket.SequenceEqual(triggerSignal1))
                    {
                        Log("✨ 핸드툴 1차 눌림 감지! (모드1)", Color.Magenta);
                        SendPacket(new CommandData { Description = "이젝터 3+4 동시 진공", Data = new byte[] { 0, 0x50, 0, 0, 0, 0, 0, 0, 0 } });
                        currentAutoModeState1 = AutoModeState.VacuumsOn;
                    }
                    break;
                case AutoModeState.VacuumsOn:
                    if (!currentPacket.SequenceEqual(triggerSignal1))
                    {
                        Log("...핸드툴 릴리즈 감지. (모드1)", Color.CornflowerBlue);
                        currentAutoModeState1 = AutoModeState.ReadyForBreak;
                    }
                    break;
                case AutoModeState.ReadyForBreak:
                    if (currentPacket.SequenceEqual(triggerSignal1))
                    {
                        Log("✨ 핸드툴 2차 눌림 감지! (모드1)", Color.Magenta);
                        currentAutoModeState1 = AutoModeState.Breaking;
                        Task.Run(async () =>
                        {
                            SendPacket(new CommandData { Description = "이젝터 3+4 동시 파기", Data = new byte[] { 0, 0xA0, 0, 0, 0, 0, 0, 0, 0 } });
                            await Task.Delay(1000);
                            SendPacket(new CommandData { Description = "명령 리셋", Data = new byte[9] { 0, 0, 0, 0, 0, 0, 0, 0, 0 } });
                            currentAutoModeState1 = AutoModeState.Idle;
                            this.Invoke((MethodInvoker)delegate { Log("✅ 자동 모드 1 사이클 완료.", Color.Green); });
                        });
                    }
                    break;
                case AutoModeState.Breaking: break;
            }
        }

        private void HandleAutoMode2(byte[] currentPacket)
        {
            // 다른 모드가 동작 중일 때는 실행하지 않음
            if (currentAutoModeState1 != AutoModeState.Idle) return;

            byte triggerByte = currentPacket[10]; // Data[6]

            switch (currentAutoModeState2)
            {
                case AutoModeState.Idle:
                    if (triggerByte == 0x04)
                    {
                        Log("✨ 핸드툴 1차 눌림 감지! (모드2)", Color.Tomato);
                        SendPacket(new CommandData { Description = "이젝터 5+6 동시 진공", Data = new byte[] { 0, 0, 0x05, 0, 0, 0, 0, 0, 0 } });
                        currentAutoModeState2 = AutoModeState.VacuumsOn;
                    }
                    break;
                case AutoModeState.VacuumsOn:
                    if (triggerByte != 0x04)
                    {
                        Log("...핸드툴 릴리즈 감지. (모드2)", Color.LightSalmon);
                        currentAutoModeState2 = AutoModeState.ReadyForBreak;
                    }
                    break;
                case AutoModeState.ReadyForBreak:
                    if (triggerByte == 0x04)
                    {
                        Log("✨ 핸드툴 2차 눌림 감지! (모드2)", Color.Tomato);
                        currentAutoModeState2 = AutoModeState.Breaking;
                        Task.Run(async () =>
                        {
                            SendPacket(new CommandData { Description = "이젝터 5+6 동시 파기", Data = new byte[] { 0, 0, 0x0A, 0, 0, 0, 0, 0, 0 } });
                            await Task.Delay(1000);
                            SendPacket(new CommandData { Description = "명령 리셋", Data = new byte[9] { 0, 0, 0, 0, 0, 0, 0, 0, 0 } });
                            currentAutoModeState2 = AutoModeState.Idle;
                            this.Invoke((MethodInvoker)delegate { Log("✅ 자동 모드 2 사이클 완료.", Color.Green); });
                        });
                    }
                    break;
                case AutoModeState.Breaking: break;
            }
        }


        private void UpdateLastPacketLabel(string packetString)
        {
            if (lblLastPacket.InvokeRequired) { lblLastPacket.Invoke((MethodInvoker)delegate { lblLastPacket.Text = packetString; }); }
            else { lblLastPacket.Text = packetString; }
        }

        private void UpdateAnalysisUI(byte[] packet)
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
            if (pollingTimer != null) pollingTimer.Stop();
            if (serialPort != null && serialPort.IsOpen) { serialPort.Close(); }
            base.OnFormClosing(e);
        }
    }
}

