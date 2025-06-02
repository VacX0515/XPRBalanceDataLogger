using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace XPRBalanceDataLogger
{
    public partial class MainForm : Form
    {
        private TcpClient tcpClient;
        private NetworkStream stream;
        private Thread readThread;
        private bool isConnected = false;
        private string currentWeight = "0.0000";
        private string currentUnit = "g";
        private bool isStable = false;
        private int measurementIndex = 1;

        public MainForm()
        {
            InitializeComponent();
            // EPPlus 라이센스 설정 (비상업적 사용)
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            UpdateConnectionStatus(false);
            dgvMeasurements.AutoGenerateColumns = false;
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                ConnectToBalance();
            }
            else
            {
                DisconnectFromBalance();
            }
        }

        private async void ConnectToBalance()
        {
            try
            {
                string ipAddress = txtIPAddress.Text.Trim();
                int port = int.Parse(txtPort.Text.Trim());

                tcpClient = new TcpClient();
                await tcpClient.ConnectAsync(ipAddress, port);
                stream = tcpClient.GetStream();

                isConnected = true;
                UpdateConnectionStatus(true);

                // 연결 후 초기화 명령 전송
                SendCommand("@"); // Abort 명령으로 초기화
                Thread.Sleep(500);

                // 저울 정보 요청
                SendCommand("I4"); // 시리얼 번호 요청

                // 무게 읽기 타이머 시작
                timerWeight.Start();

                // 읽기 스레드 시작
                readThread = new Thread(ReadFromBalance);
                readThread.IsBackground = true;
                readThread.Start();

                UpdateStatus("저울에 연결되었습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"연결 실패: {ex.Message}", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateConnectionStatus(false);
            }
        }

        private void DisconnectFromBalance()
        {
            try
            {
                timerWeight.Stop();
                isConnected = false;

                if (readThread != null && readThread.IsAlive)
                {
                    readThread.Join(1000);
                }

                if (stream != null)
                {
                    stream.Close();
                    stream.Dispose();
                }

                if (tcpClient != null)
                {
                    tcpClient.Close();
                    tcpClient.Dispose();
                }

                UpdateConnectionStatus(false);
                UpdateStatus("저울 연결이 해제되었습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"연결 해제 중 오류: {ex.Message}", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SendCommand(string command)
        {
            try
            {
                if (stream != null && stream.CanWrite)
                {
                    byte[] data = Encoding.ASCII.GetBytes(command + "\r\n");
                    stream.Write(data, 0, data.Length);
                    stream.Flush();
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"명령 전송 실패: {ex.Message}");
            }
        }

        private void ReadFromBalance()
        {
            byte[] buffer = new byte[1024];
            StringBuilder sb = new StringBuilder();

            while (isConnected)
            {
                try
                {
                    if (stream.DataAvailable)
                    {
                        int bytesRead = stream.Read(buffer, 0, buffer.Length);
                        string data = Encoding.ASCII.GetString(buffer, 0, bytesRead);
                        sb.Append(data);

                        // 응답이 완전한지 확인 (CRLF로 끝나는지)
                        string response = sb.ToString();
                        if (response.Contains("\r\n"))
                        {
                            string[] lines = response.Split(new[] { "\r\n" },
                                StringSplitOptions.RemoveEmptyEntries);

                            foreach (string line in lines)
                            {
                                ProcessResponse(line);
                            }

                            sb.Clear();
                        }
                    }
                    Thread.Sleep(10);
                }
                catch (Exception ex)
                {
                    if (isConnected)
                    {
                        Invoke(new Action(() => UpdateStatus($"읽기 오류: {ex.Message}")));
                    }
                }
            }
        }

        private void ProcessResponse(string response)
        {
            try
            {
                if (string.IsNullOrEmpty(response)) return;

                string[] parts = response.Split(' ');

                // 무게 응답 처리
                if (response.StartsWith("S "))
                {
                    if (parts.Length >= 4)
                    {
                        string stability = parts[1];
                        string weight = parts[4].Trim();
                        string unit = parts[5];

                        isStable = (stability == "S");
                        currentWeight = weight;
                        currentUnit = unit;

                        Invoke(new Action(() => UpdateWeightDisplay()));
                    }
                }
                // 시리얼 번호 응답
                else if (response.StartsWith("I4 A"))
                {
                    if (parts.Length >= 3)
                    {
                        string serialNumber = parts[2].Trim('"');
                        Invoke(new Action(() =>
                            UpdateStatus($"저울 시리얼 번호: {serialNumber}")));
                    }
                }
                // 영점 설정 응답
                else if (response.StartsWith("Z A"))
                {
                    Invoke(new Action(() => UpdateStatus("영점이 설정되었습니다.")));
                }
                // 용기무게 설정 응답
                else if (response.StartsWith("T S") || response.StartsWith("T A"))
                {
                    Invoke(new Action(() => UpdateStatus("용기무게가 설정되었습니다.")));
                }
                // 오류 응답
                else if (response.StartsWith("ES") || response.StartsWith("ET") ||
                         response.StartsWith("EL"))
                {
                    Invoke(new Action(() => UpdateStatus($"오류: {response}")));
                }
            }
            catch (Exception ex)
            {
                Invoke(new Action(() => UpdateStatus($"응답 처리 오류: {ex.Message}")));
            }
        }

        private void UpdateWeightDisplay()
        {
            lblWeight.Text = currentWeight;
            lblUnit.Text = currentUnit;
            lblStability.Text = isStable ? "안정" : "불안정";
            lblStability.ForeColor = isStable ? Color.Lime : Color.Yellow;
        }

        private void UpdateConnectionStatus(bool connected)
        {
            isConnected = connected;
            lblConnectionStatus.Text = connected ? "연결됨" : "연결 끊김";
            lblConnectionStatus.ForeColor = connected ? Color.Green : Color.Red;
            btnConnect.Text = connected ? "연결 해제" : "연결";

            // 컨트롤 활성화/비활성화
            txtIPAddress.Enabled = !connected;
            txtPort.Enabled = !connected;
            btnZero.Enabled = connected;
            btnTare.Enabled = connected;
            btnSave.Enabled = connected;
        }

        private void UpdateStatus(string message)
        {
            toolStripStatusLabel.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
        }

        private void timerWeight_Tick(object sender, EventArgs e)
        {
            if (isConnected)
            {
                SendCommand("SI"); // 즉시 무게 요청
            }
        }

        private void btnZero_Click(object sender, EventArgs e)
        {
            if (isConnected)
            {
                SendCommand("Z"); // 영점 설정
            }
        }

        private void btnTare_Click(object sender, EventArgs e)
        {
            if (isConnected)
            {
                SendCommand("T"); // 용기무게 설정
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("저울이 연결되지 않았습니다.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sampleNumber = txtSampleNumber.Text.Trim();
            if (string.IsNullOrEmpty(sampleNumber))
            {
                MessageBox.Show("샘플 번호를 입력하세요.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSampleNumber.Focus();
                return;
            }

            // 데이터 저장
            dgvMeasurements.Rows.Add(
                measurementIndex++,
                sampleNumber,
                currentWeight,
                currentUnit,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            );

            // 샘플 번호 자동 증가 (숫자인 경우)
            if (int.TryParse(sampleNumber, out int num))
            {
                txtSampleNumber.Text = (num + 1).ToString();
            }

            txtSampleNumber.SelectAll();
            txtSampleNumber.Focus();

            UpdateStatus($"측정값 저장됨: {sampleNumber} - {currentWeight} {currentUnit}");
        }

        private void dgvMeasurements_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // 수정 버튼
            if (e.ColumnIndex == colEdit.Index)
            {
                DataGridViewRow row = dgvMeasurements.Rows[e.RowIndex];
                string currentSample = row.Cells[colSampleNumber.Index].Value?.ToString();
                string currentWeightValue = row.Cells[colWeight.Index].Value?.ToString();

                using (var editForm = new EditDataForm(currentSample, currentWeightValue))
                {
                    if (editForm.ShowDialog() == DialogResult.OK)
                    {
                        row.Cells[colSampleNumber.Index].Value = editForm.SampleNumber;
                        row.Cells[colWeight.Index].Value = editForm.Weight;
                        UpdateStatus($"데이터 수정됨: {editForm.SampleNumber}");
                    }
                }
            }
            // 삭제 버튼
            else if (e.ColumnIndex == colDelete.Index)
            {
                if (MessageBox.Show("이 데이터를 삭제하시겠습니까?", "확인",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    dgvMeasurements.Rows.RemoveAt(e.RowIndex);
                    // 순번 재정렬
                    for (int i = 0; i < dgvMeasurements.Rows.Count; i++)
                    {
                        dgvMeasurements.Rows[i].Cells[colIndex.Index].Value = i + 1;
                    }
                    UpdateStatus("데이터가 삭제되었습니다.");
                }
            }
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            if (dgvMeasurements.Rows.Count == 0)
            {
                MessageBox.Show("삭제할 데이터가 없습니다.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show("모든 데이터를 삭제하시겠습니까?", "확인",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                dgvMeasurements.Rows.Clear();
                measurementIndex = 1;
                UpdateStatus("모든 데이터가 삭제되었습니다.");
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (dgvMeasurements.Rows.Count == 0)
            {
                MessageBox.Show("저장할 데이터가 없습니다.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "엑셀 파일로 저장",
                FileName = $"XPR_측정데이터_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("측정 데이터");

                        // 헤더 설정
                        worksheet.Cells[1, 1].Value = "순번";
                        worksheet.Cells[1, 2].Value = "샘플 번호";
                        worksheet.Cells[1, 3].Value = "무게";
                        worksheet.Cells[1, 4].Value = "단위";
                        worksheet.Cells[1, 5].Value = "측정시간";

                        // 헤더 스타일
                        using (var range = worksheet.Cells[1, 1, 1, 5])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        // 데이터 입력
                        for (int i = 0; i < dgvMeasurements.Rows.Count; i++)
                        {
                            var row = dgvMeasurements.Rows[i];
                            worksheet.Cells[i + 2, 1].Value = row.Cells[colIndex.Index].Value;
                            worksheet.Cells[i + 2, 2].Value = row.Cells[colSampleNumber.Index].Value;
                            worksheet.Cells[i + 2, 3].Value = row.Cells[colWeight.Index].Value;
                            worksheet.Cells[i + 2, 4].Value = row.Cells[colUnit.Index].Value;
                            worksheet.Cells[i + 2, 5].Value = row.Cells[colDateTime.Index].Value;
                        }

                        // 컬럼 너비 자동 조정
                        worksheet.Cells.AutoFitColumns();

                        // 파일 저장
                        package.SaveAs(new FileInfo(sfd.FileName));
                    }

                    MessageBox.Show("엑셀 파일이 저장되었습니다.", "완료",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdateStatus($"엑셀 파일 저장됨: {sfd.FileName}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"저장 중 오류 발생: {ex.Message}", "오류",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "엑셀 파일 불러오기"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(ofd.FileName)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension?.Rows ?? 0;

                        if (rowCount > 1)
                        {
                            dgvMeasurements.Rows.Clear();
                            measurementIndex = 1;

                            for (int row = 2; row <= rowCount; row++)
                            {
                                string sampleNumber = worksheet.Cells[row, 2].Value?.ToString() ?? "";
                                string weight = worksheet.Cells[row, 3].Value?.ToString() ?? "0";
                                string unit = worksheet.Cells[row, 4].Value?.ToString() ?? "g";
                                string dateTime = worksheet.Cells[row, 5].Value?.ToString() ??
                                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                                dgvMeasurements.Rows.Add(
                                    measurementIndex++,
                                    sampleNumber,
                                    weight,
                                    unit,
                                    dateTime
                                );
                            }

                            MessageBox.Show($"{rowCount - 1}개의 데이터를 불러왔습니다.", "완료",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            UpdateStatus($"엑셀 파일 불러옴: {ofd.FileName}");
                        }
                        else
                        {
                            MessageBox.Show("불러올 데이터가 없습니다.", "알림",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"불러오기 중 오류 발생: {ex.Message}", "오류",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (isConnected)
            {
                DisconnectFromBalance();
            }
        }

        private void lblUnit_Click(object sender, EventArgs e)
        {

        }
    }
}