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
            // EPPlus ���̼��� ���� (������ ���)
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

                // ���� �� �ʱ�ȭ ��� ����
                SendCommand("@"); // Abort ������� �ʱ�ȭ
                Thread.Sleep(500);

                // ���� ���� ��û
                SendCommand("I4"); // �ø��� ��ȣ ��û

                // ���� �б� Ÿ�̸� ����
                timerWeight.Start();

                // �б� ������ ����
                readThread = new Thread(ReadFromBalance);
                readThread.IsBackground = true;
                readThread.Start();

                UpdateStatus("���￡ ����Ǿ����ϴ�.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"���� ����: {ex.Message}", "����",
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
                UpdateStatus("���� ������ �����Ǿ����ϴ�.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"���� ���� �� ����: {ex.Message}", "����",
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
                UpdateStatus($"��� ���� ����: {ex.Message}");
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

                        // ������ �������� Ȯ�� (CRLF�� ��������)
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
                        Invoke(new Action(() => UpdateStatus($"�б� ����: {ex.Message}")));
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

                // ���� ���� ó��
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
                // �ø��� ��ȣ ����
                else if (response.StartsWith("I4 A"))
                {
                    if (parts.Length >= 3)
                    {
                        string serialNumber = parts[2].Trim('"');
                        Invoke(new Action(() =>
                            UpdateStatus($"���� �ø��� ��ȣ: {serialNumber}")));
                    }
                }
                // ���� ���� ����
                else if (response.StartsWith("Z A"))
                {
                    Invoke(new Action(() => UpdateStatus("������ �����Ǿ����ϴ�.")));
                }
                // ��⹫�� ���� ����
                else if (response.StartsWith("T S") || response.StartsWith("T A"))
                {
                    Invoke(new Action(() => UpdateStatus("��⹫�԰� �����Ǿ����ϴ�.")));
                }
                // ���� ����
                else if (response.StartsWith("ES") || response.StartsWith("ET") ||
                         response.StartsWith("EL"))
                {
                    Invoke(new Action(() => UpdateStatus($"����: {response}")));
                }
            }
            catch (Exception ex)
            {
                Invoke(new Action(() => UpdateStatus($"���� ó�� ����: {ex.Message}")));
            }
        }

        private void UpdateWeightDisplay()
        {
            lblWeight.Text = currentWeight;
            lblUnit.Text = currentUnit;
            lblStability.Text = isStable ? "����" : "�Ҿ���";
            lblStability.ForeColor = isStable ? Color.Lime : Color.Yellow;
        }

        private void UpdateConnectionStatus(bool connected)
        {
            isConnected = connected;
            lblConnectionStatus.Text = connected ? "�����" : "���� ����";
            lblConnectionStatus.ForeColor = connected ? Color.Green : Color.Red;
            btnConnect.Text = connected ? "���� ����" : "����";

            // ��Ʈ�� Ȱ��ȭ/��Ȱ��ȭ
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
                SendCommand("SI"); // ��� ���� ��û
            }
        }

        private void btnZero_Click(object sender, EventArgs e)
        {
            if (isConnected)
            {
                SendCommand("Z"); // ���� ����
            }
        }

        private void btnTare_Click(object sender, EventArgs e)
        {
            if (isConnected)
            {
                SendCommand("T"); // ��⹫�� ����
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("������ ������� �ʾҽ��ϴ�.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sampleNumber = txtSampleNumber.Text.Trim();
            if (string.IsNullOrEmpty(sampleNumber))
            {
                MessageBox.Show("���� ��ȣ�� �Է��ϼ���.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSampleNumber.Focus();
                return;
            }

            // ������ ����
            dgvMeasurements.Rows.Add(
                measurementIndex++,
                sampleNumber,
                currentWeight,
                currentUnit,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            );

            // ���� ��ȣ �ڵ� ���� (������ ���)
            if (int.TryParse(sampleNumber, out int num))
            {
                txtSampleNumber.Text = (num + 1).ToString();
            }

            txtSampleNumber.SelectAll();
            txtSampleNumber.Focus();

            UpdateStatus($"������ �����: {sampleNumber} - {currentWeight} {currentUnit}");
        }

        private void dgvMeasurements_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // ���� ��ư
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
                        UpdateStatus($"������ ������: {editForm.SampleNumber}");
                    }
                }
            }
            // ���� ��ư
            else if (e.ColumnIndex == colDelete.Index)
            {
                if (MessageBox.Show("�� �����͸� �����Ͻðڽ��ϱ�?", "Ȯ��",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    dgvMeasurements.Rows.RemoveAt(e.RowIndex);
                    // ���� ������
                    for (int i = 0; i < dgvMeasurements.Rows.Count; i++)
                    {
                        dgvMeasurements.Rows[i].Cells[colIndex.Index].Value = i + 1;
                    }
                    UpdateStatus("�����Ͱ� �����Ǿ����ϴ�.");
                }
            }
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            if (dgvMeasurements.Rows.Count == 0)
            {
                MessageBox.Show("������ �����Ͱ� �����ϴ�.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show("��� �����͸� �����Ͻðڽ��ϱ�?", "Ȯ��",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                dgvMeasurements.Rows.Clear();
                measurementIndex = 1;
                UpdateStatus("��� �����Ͱ� �����Ǿ����ϴ�.");
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (dgvMeasurements.Rows.Count == 0)
            {
                MessageBox.Show("������ �����Ͱ� �����ϴ�.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "���� ���Ϸ� ����",
                FileName = $"XPR_����������_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("���� ������");

                        // ��� ����
                        worksheet.Cells[1, 1].Value = "����";
                        worksheet.Cells[1, 2].Value = "���� ��ȣ";
                        worksheet.Cells[1, 3].Value = "����";
                        worksheet.Cells[1, 4].Value = "����";
                        worksheet.Cells[1, 5].Value = "�����ð�";

                        // ��� ��Ÿ��
                        using (var range = worksheet.Cells[1, 1, 1, 5])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        // ������ �Է�
                        for (int i = 0; i < dgvMeasurements.Rows.Count; i++)
                        {
                            var row = dgvMeasurements.Rows[i];
                            worksheet.Cells[i + 2, 1].Value = row.Cells[colIndex.Index].Value;
                            worksheet.Cells[i + 2, 2].Value = row.Cells[colSampleNumber.Index].Value;
                            worksheet.Cells[i + 2, 3].Value = row.Cells[colWeight.Index].Value;
                            worksheet.Cells[i + 2, 4].Value = row.Cells[colUnit.Index].Value;
                            worksheet.Cells[i + 2, 5].Value = row.Cells[colDateTime.Index].Value;
                        }

                        // �÷� �ʺ� �ڵ� ����
                        worksheet.Cells.AutoFitColumns();

                        // ���� ����
                        package.SaveAs(new FileInfo(sfd.FileName));
                    }

                    MessageBox.Show("���� ������ ����Ǿ����ϴ�.", "�Ϸ�",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdateStatus($"���� ���� �����: {sfd.FileName}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"���� �� ���� �߻�: {ex.Message}", "����",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "���� ���� �ҷ�����"
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

                            MessageBox.Show($"{rowCount - 1}���� �����͸� �ҷ��Խ��ϴ�.", "�Ϸ�",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            UpdateStatus($"���� ���� �ҷ���: {ofd.FileName}");
                        }
                        else
                        {
                            MessageBox.Show("�ҷ��� �����Ͱ� �����ϴ�.", "�˸�",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"�ҷ����� �� ���� �߻�: {ex.Message}", "����",
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