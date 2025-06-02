using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace XPRBalanceDataLogger
{
    public partial class MainForm : Form
    {
        // 측정 항목 열거형
        public enum MeasurementType
        {
            InitialPlate,              // 초기 기판
            InitialPlateSample,        // 초기 기판 + 시료
            FinalPlateSample,          // 최종 기판 + 시료
            FinalKeepPlateSample,      // 최종 보관 기판 + 시료
            InitialCollector,          // 최초 응집판
            FinalCollector             // 최종 응집판
        }
        private void dgvMeasurements_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // 수정 버튼
            if (e.ColumnIndex == 5)
            {
                var data = measurements[e.RowIndex];
                // 무게 표시를 동적으로 포맷팅
                string weightText = data.Weight.ToString("F10").TrimEnd('0');
                if (weightText.EndsWith("."))
                    weightText += "0";

                using (var editForm = new EditDataForm(data.SampleNumber, weightText))
                {
                    if (editForm.ShowDialog() == DialogResult.OK)
                    {
                        data.SampleNumber = editForm.SampleNumber;
                        data.Weight = double.Parse(editForm.Weight, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture);
                        UpdateDataGridView();
                        UpdateStatus($"데이터 수정 완료: {data.SampleNumber}");
                    }
                }
            }
            // 삭제 버튼
            else if (e.ColumnIndex == 6)
            {
                if (MessageBox.Show("선택한 데이터를 삭제하시겠습니까?", "확인",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    measurements.RemoveAt(e.RowIndex);
                    // 인덱스 재정렬
                    for (int i = 0; i < measurements.Count; i++)
                    {
                        measurements[i].Index = i + 1;
                    }
                    UpdateDataGridView();
                    UpdateStatus("데이터 삭제 완료");
                }
            }
        }

        // 측정 항목별 구분자 딕셔너리
        private readonly Dictionary<MeasurementType, string> measurementPrefixes = new Dictionary<MeasurementType, string>
        {
            { MeasurementType.InitialPlate, "IP" },
            { MeasurementType.InitialPlateSample, "IPS" },
            { MeasurementType.FinalPlateSample, "FPS" },
            { MeasurementType.FinalKeepPlateSample, "FKPS" },
            { MeasurementType.InitialCollector, "IC" },
            { MeasurementType.FinalCollector, "FC" }
        };

        // 측정 항목별 표시 이름
        private readonly Dictionary<MeasurementType, string> measurementTypeNames = new Dictionary<MeasurementType, string>
        {
            { MeasurementType.InitialPlate, "초기 기판" },
            { MeasurementType.InitialPlateSample, "초기 기판 + 시료" },
            { MeasurementType.FinalPlateSample, "최종 기판 + 시료" },
            { MeasurementType.FinalKeepPlateSample, "최종 보관 기판 + 시료" },
            { MeasurementType.InitialCollector, "최초 응집판" },
            { MeasurementType.FinalCollector, "최종 응집판" }
        };

        private TcpClient client;
        private NetworkStream stream;
        private bool isConnected = false;
        private List<MeasurementData> measurements;
        private double currentWeight = 0;
        private string currentUnit = "g";
        private bool isStable = false;

        // 자동 저장 관련 필드
        private bool isAutoMeasureEnabled = false;
        private int stabilityTimeSeconds = 3;
        private DateTime doorClosedTime = DateTime.Now;
        private bool isDoorOpen = false;
        private bool wasDoorOpen = false;
        private bool isWaitingForStability = false;
        private int doorStatus = -1; // -1: unknown, 0: closed, 1: open

        public MainForm()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            measurements = new List<MeasurementData>();
            UpdateDataGridView();
            UpdateStatus("프로그램이 시작되었습니다.");

            // 카테고리 콤보박스 초기화
            InitializeCategoryComboBox();

            // 안정 시간 설정 초기화
            nudStabilityTime.Value = 3;
            btnStartMeasurement.Enabled = true;

            // DataGridView 행 색상 설정
            dgvMeasurements.RowsDefaultCellStyle.BackColor = Color.White;
            dgvMeasurements.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
        }

        // 카테고리 콤보박스 초기화
        private void InitializeCategoryComboBox()
        {
            cboCategory.Items.Clear();
            cboCategory.Items.Add("카테고리 선택");
            cboCategory.Items.Add("Sample");
            cboCategory.Items.Add("Collector");
            cboCategory.SelectedIndex = 0;

            // 측정 항목 콤보박스 초기화
            cboMeasurementType.Items.Clear();
            cboMeasurementType.Enabled = false;
        }

        // 카테고리 선택 변경 이벤트
        private void cboCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboMeasurementType.Items.Clear();
            cboMeasurementType.Enabled = false;

            // 샘플 번호 초기화
            txtSampleNumber.Text = "";

            if (cboCategory.SelectedIndex == 1) // Sample
            {
                cboMeasurementType.Items.Add("항목 선택");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.InitialPlate]} ({measurementPrefixes[MeasurementType.InitialPlate]})");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.InitialPlateSample]} ({measurementPrefixes[MeasurementType.InitialPlateSample]})");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.FinalPlateSample]} ({measurementPrefixes[MeasurementType.FinalPlateSample]})");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.FinalKeepPlateSample]} ({measurementPrefixes[MeasurementType.FinalKeepPlateSample]})");
                cboMeasurementType.Enabled = true;
                cboMeasurementType.SelectedIndex = 0;
            }
            else if (cboCategory.SelectedIndex == 2) // Collector
            {
                cboMeasurementType.Items.Add("항목 선택");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.InitialCollector]} ({measurementPrefixes[MeasurementType.InitialCollector]})");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.FinalCollector]} ({measurementPrefixes[MeasurementType.FinalCollector]})");
                cboMeasurementType.Enabled = true;
                cboMeasurementType.SelectedIndex = 0;
            }
        }

        private async void btnConnect_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                try
                {
                    string ipAddress = txtIPAddress.Text;
                    int port = int.Parse(txtPort.Text);

                    client = new TcpClient();
                    await client.ConnectAsync(ipAddress, port);
                    stream = client.GetStream();

                    // 초기 명령 전송 (@)
                    string response = await SendCommand("@");
                    if (!string.IsNullOrEmpty(response))
                    {
                        isConnected = true;
                        btnConnect.Text = "연결 끊기";
                        lblConnectionStatus.Text = "연결됨";
                        lblConnectionStatus.ForeColor = Color.Green;

                        // 무게 읽기 타이머 시작
                        timerWeight.Start();

                        // 초기 도어 상태 확인
                        await CheckDoorStatusOnce();

                        UpdateStatus($"저울 연결 성공: {response}");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"연결 실패: {ex.Message}", "오류",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatus($"연결 실패: {ex.Message}");
                }
            }
            else
            {
                Disconnect();
            }
        }

        // 도어 상태 한 번 확인
        private async Task CheckDoorStatusOnce()
        {
            try
            {
                string response = await SendCommand("WS");
                if (!string.IsNullOrEmpty(response) && response.StartsWith("WS "))
                {
                    doorStatus = int.Parse(response.Substring(5, 1));
                    isDoorOpen = doorStatus > 0;
                    UpdateDoorStatus(isDoorOpen);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"도어 상태 확인 오류: {ex.Message}");
            }
        }

        private void Disconnect()
        {
            try
            {
                timerWeight.Stop();
                timerAutoSave.Stop();
                timerDoorCheck.Stop();

                stream?.Close();
                client?.Close();

                isConnected = false;
                btnConnect.Text = "연결";
                lblConnectionStatus.Text = "연결 끊김";
                lblConnectionStatus.ForeColor = Color.Red;

                UpdateStatus("저울 연결이 끊어졌습니다.");
            }
            catch (Exception ex)
            {
                UpdateStatus($"연결 끊기 오류: {ex.Message}");
            }
        }

        private async Task<string> SendCommand(string command)
        {
            try
            {
                if (stream == null || !stream.CanWrite)
                    return string.Empty;

                byte[] data = Encoding.ASCII.GetBytes(command + "\r\n");
                await stream.WriteAsync(data, 0, data.Length);

                byte[] buffer = new byte[1024];
                int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length);

                return Encoding.ASCII.GetString(buffer, 0, bytesRead).Trim();
            }
            catch (Exception ex)
            {
                UpdateStatus($"명령 전송 오류: {ex.Message}");
                return string.Empty;
            }
        }

        private async void timerWeight_Tick(object sender, EventArgs e)
        {
            if (!isConnected || client == null || !client.Connected)
                return;

            try
            {
                string response = await SendCommand("SI");
                if (!string.IsNullOrEmpty(response))
                {
                    ParseWeightResponse(response);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"무게 읽기 오류: {ex.Message}");
            }
        }

        private void ParseWeightResponse(string response)
        {
            try
            {
                // 응답 형식: S S     1.2345678 g
                // 처음 2자리는 command ID, 3번째는 status, 4-14는 weight value, 15부터는 unit
                if (response.Length >= 14)
                {
                    string status = response.Substring(2, 1);
                    isStable = (status == "S");

                    // 무게값은 4번째 문자부터 시작하여 공백 제거
                    string weightStr = response.Substring(3).Trim();

                    // 단위 분리
                    int spaceIndex = weightStr.IndexOf(' ');
                    if (spaceIndex > 0)
                    {
                        string weightPart = weightStr.Substring(0, spaceIndex).Trim();
                        currentUnit = weightStr.Substring(spaceIndex + 1).Trim();

                        if (double.TryParse(weightPart, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out double weight))
                        {
                            currentWeight = weight;
                        }
                    }

                    UpdateWeightDisplay();
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"응답 파싱 오류: {ex.Message}");
            }
        }

        private void UpdateWeightDisplay()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(UpdateWeightDisplay));
                return;
            }

            // 소수점 이하 자리수를 동적으로 결정
            string weightText = currentWeight.ToString("F10").TrimEnd('0');
            if (weightText.EndsWith("."))
                weightText += "0";

            lblWeight.Text = weightText;
            lblUnit.Text = currentUnit;
            lblStability.Text = isStable ? "안정" : "불안정";
            lblStability.ForeColor = isStable ? Color.Lime : Color.Yellow;
        }

        private async void btnZero_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("저울이 연결되지 않았습니다.", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string response = await SendCommand("Z");
            if (response.Contains("A"))
            {
                UpdateStatus("영점 설정 완료");
            }
            else
            {
                UpdateStatus($"영점 설정 실패: {response}");
            }
        }

        private async void btnTare_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("저울이 연결되지 않았습니다.", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string response = await SendCommand("T");
            if (response.Contains("S"))
            {
                UpdateStatus("용기무게 설정 완료");
            }
            else
            {
                UpdateStatus($"용기무게 설정 실패: {response}");
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("저울이 연결되지 않았습니다.", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 카테고리 선택 확인
            if (cboCategory.SelectedIndex <= 0)
            {
                MessageBox.Show("카테고리를 선택하세요.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboCategory.Focus();
                return;
            }

            // 측정 항목 선택 확인
            if (cboMeasurementType.SelectedIndex <= 0)
            {
                MessageBox.Show("측정 항목을 선택하세요.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboMeasurementType.Focus();
                return;
            }

            // 샘플 번호 확인
            string sampleNumber = txtSampleNumber.Text.Trim();
            if (string.IsNullOrEmpty(sampleNumber))
            {
                MessageBox.Show("샘플 번호를 입력하세요.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSampleNumber.Focus();
                return;
            }

            // 측정 항목 가져오기
            string measurementType = cboMeasurementType.SelectedItem.ToString();

            // 데이터 저장
            measurements.Add(new MeasurementData
            {
                Index = measurements.Count + 1,
                SampleNumber = sampleNumber,
                Weight = currentWeight,
                Unit = currentUnit,
                DateTime = DateTime.Now,
                MeasurementType = measurementType
            });

            // DataGridView 업데이트
            UpdateDataGridView();

            // 샘플 번호 자동 증가
            IncrementSampleNumberWithPrefix();

            // 상태 업데이트
            UpdateStatus($"저장 완료: {sampleNumber} - {currentWeight:F7} {currentUnit} ({measurementType})");
        }

        // 콤보박스 선택 변경 이벤트 핸들러
        private void cboMeasurementType_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateSampleNumberWithPrefix();
        }

        // 샘플 번호에 구분자 추가하는 메서드
        private void UpdateSampleNumberWithPrefix()
        {
            if (cboMeasurementType.SelectedIndex <= 0)
                return;

            // 현재 샘플 번호
            string currentSampleNumber = txtSampleNumber.Text.Trim();

            // 기존 구분자 제거 (있는 경우)
            foreach (var prefix in measurementPrefixes.Values)
            {
                if (currentSampleNumber.StartsWith(prefix + "_"))
                {
                    currentSampleNumber = currentSampleNumber.Substring(prefix.Length + 1);
                    break;
                }
            }

            // 카테고리별로 선택된 항목의 실제 인덱스 계산
            MeasurementType selectedType;
            if (cboCategory.SelectedIndex == 1) // Sample
            {
                switch (cboMeasurementType.SelectedIndex)
                {
                    case 1: selectedType = MeasurementType.InitialPlate; break;
                    case 2: selectedType = MeasurementType.InitialPlateSample; break;
                    case 3: selectedType = MeasurementType.FinalPlateSample; break;
                    case 4: selectedType = MeasurementType.FinalKeepPlateSample; break;
                    default: return;
                }
            }
            else if (cboCategory.SelectedIndex == 2) // Collector
            {
                switch (cboMeasurementType.SelectedIndex)
                {
                    case 1: selectedType = MeasurementType.InitialCollector; break;
                    case 2: selectedType = MeasurementType.FinalCollector; break;
                    default: return;
                }
            }
            else
            {
                return;
            }

            string newPrefix = measurementPrefixes[selectedType];

            // 샘플 번호가 비어있으면 자동 번호 생성
            if (string.IsNullOrWhiteSpace(currentSampleNumber))
            {
                int count = measurements.Count(m => m.SampleNumber.StartsWith(newPrefix + "_")) + 1;
                currentSampleNumber = count.ToString("D3");
            }

            txtSampleNumber.Text = $"{newPrefix}_{currentSampleNumber}";
        }

        // 샘플 번호 자동 증가 메서드
        private void IncrementSampleNumberWithPrefix()
        {
            string currentNumber = txtSampleNumber.Text;
            if (string.IsNullOrWhiteSpace(currentNumber))
                return;

            // 구분자와 번호 분리
            int underscoreIndex = currentNumber.IndexOf('_');
            if (underscoreIndex > 0)
            {
                string prefix = currentNumber.Substring(0, underscoreIndex);
                string numberPart = currentNumber.Substring(underscoreIndex + 1);

                // 숫자 부분만 추출
                string numericPart = "";
                string suffix = "";

                for (int i = 0; i < numberPart.Length; i++)
                {
                    if (char.IsDigit(numberPart[i]))
                    {
                        numericPart += numberPart[i];
                    }
                    else
                    {
                        suffix = numberPart.Substring(i);
                        break;
                    }
                }

                if (!string.IsNullOrEmpty(numericPart) && int.TryParse(numericPart, out int number))
                {
                    number++;
                    txtSampleNumber.Text = $"{prefix}_{number:D3}{suffix}";
                }
            }
        }

        // 측정 시작/중지 버튼 클릭 이벤트
        private void btnStartMeasurement_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("저울이 연결되지 않았습니다.", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 카테고리와 측정 항목 확인
            if (cboCategory.SelectedIndex <= 0)
            {
                MessageBox.Show("카테고리를 선택하세요.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboCategory.Focus();
                return;
            }

            if (cboMeasurementType.SelectedIndex <= 0)
            {
                MessageBox.Show("측정 항목을 선택하세요.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboMeasurementType.Focus();
                return;
            }

            if (!isAutoMeasureEnabled)
            {
                // 측정 시작
                isAutoMeasureEnabled = true;
                stabilityTimeSeconds = (int)nudStabilityTime.Value;
                timerDoorCheck.Start();
                timerAutoSave.Start();

                btnStartMeasurement.Text = "측정 중지";
                btnStartMeasurement.BackColor = Color.Red;

                // 컨트롤 비활성화
                cboCategory.Enabled = false;
                cboMeasurementType.Enabled = false;
                nudStabilityTime.Enabled = false;

                UpdateStatus("자동 측정 모드 시작 - 도어를 열고 샘플을 교체하세요.");
            }
            else
            {
                // 측정 중지
                isAutoMeasureEnabled = false;
                timerDoorCheck.Stop();
                timerAutoSave.Stop();

                btnStartMeasurement.Text = "측정 시작";
                btnStartMeasurement.BackColor = Color.FromArgb(0, 192, 0);

                // 컨트롤 활성화
                cboCategory.Enabled = true;
                cboMeasurementType.Enabled = true;
                nudStabilityTime.Enabled = true;

                UpdateStatus("자동 측정 모드 중지");
            }
        }

        // 도어 상태 확인 타이머
        private async void timerDoorCheck_Tick(object sender, EventArgs e)
        {
            if (!isConnected || !isAutoMeasureEnabled)
                return;

            try
            {
                // WS 명령으로 도어 상태 확인
                string response = await SendCommand("WS");
                if (!string.IsNullOrEmpty(response) && response.StartsWith("WS "))
                {
                    // WS 0: 도어 닫힘, WS 1-7: 도어 열림
                    int newDoorStatus = int.Parse(response.Substring(5, 1));

                    if (newDoorStatus != doorStatus)
                    {
                        doorStatus = newDoorStatus;
                        isDoorOpen = doorStatus > 0;

                        // 도어 상태 표시 업데이트
                        UpdateDoorStatus(isDoorOpen);

                        // 도어가 닫혔을 때
                        if (!isDoorOpen && wasDoorOpen)
                        {
                            doorClosedTime = DateTime.Now;
                            isWaitingForStability = true;
                            UpdateStatus("도어 닫힘 - 무게 안정 대기 중...");
                        }

                        wasDoorOpen = isDoorOpen;
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"도어 상태 확인 오류: {ex.Message}");
            }
        }

        // 도어 상태 표시 업데이트
        private void UpdateDoorStatus(bool isOpen)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<bool>(UpdateDoorStatus), isOpen);
                return;
            }

            lblDoorStatus.Text = isOpen ? "도어: 열림" : "도어: 닫힘";
            lblDoorStatus.ForeColor = isOpen ? Color.Red : Color.Blue;
        }

        // 자동 저장 타이머 이벤트
        private void timerAutoSave_Tick(object sender, EventArgs e)
        {
            if (!isConnected || !isAutoMeasureEnabled || !isWaitingForStability)
                return;

            try
            {
                // 도어가 닫혀있고, 무게가 안정되었는지 확인
                if (!isDoorOpen && isStable && currentWeight > 0.1)
                {
                    // 도어가 닫힌 후 설정된 시간이 경과했는지 확인
                    TimeSpan timeSinceDoorClosed = DateTime.Now - doorClosedTime;

                    if (timeSinceDoorClosed.TotalSeconds >= stabilityTimeSeconds)
                    {
                        // 자동 저장 수행
                        PerformAutoSave();
                        isWaitingForStability = false;
                        UpdateStatus("자동 저장 완료 - 도어를 열고 다음 샘플을 준비하세요.");
                    }
                    else
                    {
                        int remainingTime = stabilityTimeSeconds - (int)timeSinceDoorClosed.TotalSeconds;
                        UpdateStatus($"안정 대기 중... {remainingTime}초");
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"자동 저장 오류: {ex.Message}");
            }
        }

        // 자동 저장 수행 메서드
        private void PerformAutoSave()
        {
            // 샘플 번호 자동 생성
            string autoSampleNumber = txtSampleNumber.Text;
            if (string.IsNullOrWhiteSpace(autoSampleNumber))
            {
                // 현재 선택된 측정 타입 가져오기
                MeasurementType selectedType;
                if (cboCategory.SelectedIndex == 1) // Sample
                {
                    switch (cboMeasurementType.SelectedIndex)
                    {
                        case 1: selectedType = MeasurementType.InitialPlate; break;
                        case 2: selectedType = MeasurementType.InitialPlateSample; break;
                        case 3: selectedType = MeasurementType.FinalPlateSample; break;
                        case 4: selectedType = MeasurementType.FinalKeepPlateSample; break;
                        default: return;
                    }
                }
                else if (cboCategory.SelectedIndex == 2) // Collector
                {
                    switch (cboMeasurementType.SelectedIndex)
                    {
                        case 1: selectedType = MeasurementType.InitialCollector; break;
                        case 2: selectedType = MeasurementType.FinalCollector; break;
                        default: return;
                    }
                }
                else
                {
                    return;
                }

                string prefix = measurementPrefixes[selectedType];
                int count = measurements.Count(m => m.SampleNumber.StartsWith(prefix + "_")) + 1;
                autoSampleNumber = $"{prefix}_{count:D3}";
                txtSampleNumber.Text = autoSampleNumber;
            }

            // 측정 항목 가져오기
            string measurementType = cboMeasurementType.SelectedItem.ToString();

            // 데이터 저장
            measurements.Add(new MeasurementData
            {
                Index = measurements.Count + 1,
                SampleNumber = autoSampleNumber,
                Weight = currentWeight,
                Unit = currentUnit,
                DateTime = DateTime.Now,
                MeasurementType = measurementType
            });

            // DataGridView 업데이트
            UpdateDataGridView();

            // 상태 업데이트
            UpdateStatus($"자동 저장 완료: {autoSampleNumber} - {currentWeight:F7} {currentUnit} ({measurementType})");

            // 샘플 번호 자동 증가
            IncrementSampleNumberWithPrefix();
        }

        private void UpdateDataGridView()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(UpdateDataGridView));
                return;
            }

            dgvMeasurements.Rows.Clear();
            foreach (var data in measurements)
            {
                // 무게 표시를 동적으로 포맷팅
                string weightText = data.Weight.ToString("F10").TrimEnd('0');
                if (weightText.EndsWith("."))
                    weightText += "0";

                int rowIndex = dgvMeasurements.Rows.Add(
                    data.Index,
                    data.SampleNumber,
                    weightText,
                    data.Unit,
                    data.DateTime.ToString("yyyy-MM-dd HH:mm:ss"),
                    "수정",
                    "삭제"
                );

                // 측정 항목을 툴팁으로 표시
                if (!string.IsNullOrEmpty(data.MeasurementType))
                {
                    dgvMeasurements.Rows[rowIndex].Cells[1].ToolTipText = data.MeasurementType;
                }
            }
        }

        // DataGridView 셀 더블클릭 이벤트 - 직접 편집
        private void dgvMeasurements_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // 샘플 번호와 무게 열만 편집 가능
            if (e.ColumnIndex == 1 || e.ColumnIndex == 2)
            {
                dgvMeasurements.BeginEdit(true);
            }
        }

        // DataGridView 셀 편집 완료 이벤트
        private void dgvMeasurements_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            try
            {
                var data = measurements[e.RowIndex];

                // 샘플 번호 변경
                if (e.ColumnIndex == 1)
                {
                    string newSampleNumber = dgvMeasurements.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                    if (!string.IsNullOrWhiteSpace(newSampleNumber))
                    {
                        data.SampleNumber = newSampleNumber;
                        UpdateStatus($"샘플 번호 변경 완료: {newSampleNumber}");
                    }
                    else
                    {
                        // 빈 값이면 원래 값으로 복원
                        dgvMeasurements.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = data.SampleNumber;
                    }
                }
                // 무게 변경
                else if (e.ColumnIndex == 2)
                {
                    string weightStr = dgvMeasurements.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                    if (!string.IsNullOrWhiteSpace(weightStr) &&
                        double.TryParse(weightStr, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out double newWeight))
                    {
                        data.Weight = newWeight;
                        UpdateStatus($"무게 변경 완료: {newWeight:F7} {data.Unit}");
                    }
                    else
                    {
                        // 유효하지 않은 값이면 원래 값으로 복원
                        string weightText = data.Weight.ToString("F10").TrimEnd('0');
                        if (weightText.EndsWith("."))
                            weightText += "0";
                        dgvMeasurements.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = weightText;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"데이터 수정 중 오류 발생: {ex.Message}", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateDataGridView();
            }
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            if (measurements.Count == 0)
            {
                MessageBox.Show("삭제할 데이터가 없습니다.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show($"전체 {measurements.Count}개의 데이터를 삭제하시겠습니까?", "확인",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                measurements.Clear();
                UpdateDataGridView();
                UpdateStatus("전체 데이터 삭제 완료");
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (measurements.Count == 0)
            {
                MessageBox.Show("저장할 데이터가 없습니다.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Files|*.xlsx";
                sfd.FileName = $"측정데이터_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var package = new ExcelPackage())
                        {
                            var worksheet = package.Workbook.Worksheets.Add("측정 데이터");

                            // 헤더 추가
                            worksheet.Cells[1, 1].Value = "순번";
                            worksheet.Cells[1, 2].Value = "샘플 번호";
                            worksheet.Cells[1, 3].Value = "무게";
                            worksheet.Cells[1, 4].Value = "단위";
                            worksheet.Cells[1, 5].Value = "측정시간";
                            worksheet.Cells[1, 6].Value = "측정항목";

                            // 헤더 스타일
                            using (var range = worksheet.Cells[1, 1, 1, 6])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            }

                            // 데이터 추가
                            for (int i = 0; i < measurements.Count; i++)
                            {
                                var data = measurements[i];
                                worksheet.Cells[i + 2, 1].Value = data.Index;
                                worksheet.Cells[i + 2, 2].Value = data.SampleNumber;
                                worksheet.Cells[i + 2, 3].Value = data.Weight;
                                worksheet.Cells[i + 2, 4].Value = data.Unit;
                                worksheet.Cells[i + 2, 5].Value = data.DateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                worksheet.Cells[i + 2, 6].Value = data.MeasurementType;
                            }

                            // 열 너비 자동 조정
                            worksheet.Cells.AutoFitColumns();

                            // 파일 저장
                            package.SaveAs(new System.IO.FileInfo(sfd.FileName));
                        }

                        MessageBox.Show("엑셀 파일로 저장되었습니다.", "완료",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdateStatus($"엑셀 파일 저장 완료: {sfd.FileName}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"저장 중 오류 발생: {ex.Message}", "오류",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var package = new ExcelPackage(new System.IO.FileInfo(ofd.FileName)))
                        {
                            var worksheet = package.Workbook.Worksheets[0];
                            int rowCount = worksheet.Dimension.Rows;

                            // 기존 데이터를 덮어쓸지 확인
                            if (measurements.Count > 0)
                            {
                                var result = MessageBox.Show("기존 데이터에 추가하시겠습니까?\n" +
                                    "아니오를 선택하면 기존 데이터가 삭제됩니다.", "확인",
                                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                if (result == DialogResult.Cancel)
                                    return;
                                else if (result == DialogResult.No)
                                    measurements.Clear();
                            }

                            // 데이터 읽기 (헤더 행 제외)
                            for (int row = 2; row <= rowCount; row++)
                            {
                                if (worksheet.Cells[row, 2].Value != null)
                                {
                                    measurements.Add(new MeasurementData
                                    {
                                        Index = measurements.Count + 1,
                                        SampleNumber = worksheet.Cells[row, 2].Value.ToString(),
                                        Weight = Convert.ToDouble(worksheet.Cells[row, 3].Value),
                                        Unit = worksheet.Cells[row, 4].Value?.ToString() ?? "g",
                                        DateTime = DateTime.Parse(worksheet.Cells[row, 5].Value.ToString()),
                                        MeasurementType = worksheet.Cells[row, 6].Value?.ToString() ?? ""
                                    });
                                }
                            }

                            UpdateDataGridView();
                            MessageBox.Show($"{rowCount - 1}개의 데이터를 불러왔습니다.", "완료",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            UpdateStatus($"엑셀 파일 불러오기 완료: {ofd.FileName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"불러오기 중 오류 발생: {ex.Message}", "오류",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void UpdateStatus(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(UpdateStatus), message);
                return;
            }

            toolStripStatusLabel.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 자동 측정 중이면 중지
            if (isAutoMeasureEnabled)
            {
                var result = MessageBox.Show("자동 측정이 진행 중입니다. 종료하시겠습니까?",
                    "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                isAutoMeasureEnabled = false;
                timerDoorCheck.Stop();
                timerAutoSave.Stop();
            }

            if (measurements.Count > 0)
            {
                var result = MessageBox.Show("저장되지 않은 데이터가 있습니다. 종료하시겠습니까?",
                    "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
            }

            Disconnect();
        }

        private void lblUnit_Click(object sender, EventArgs e)
        {

        }
    }

    // 측정 데이터 클래스
    public class MeasurementData
    {
        public int Index { get; set; }
        public string SampleNumber { get; set; }
        public double Weight { get; set; }
        public string Unit { get; set; }
        public DateTime DateTime { get; set; }
        public string MeasurementType { get; set; }
    }
}