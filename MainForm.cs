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
        // ���� �׸� ������
        public enum MeasurementType
        {
            InitialPlate,              // �ʱ� ����
            InitialPlateSample,        // �ʱ� ���� + �÷�
            FinalPlateSample,          // ���� ���� + �÷�
            FinalKeepPlateSample,      // ���� ���� ���� + �÷�
            InitialCollector,          // ���� ������
            FinalCollector             // ���� ������
        }
        private void dgvMeasurements_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // ���� ��ư
            if (e.ColumnIndex == 5)
            {
                var data = measurements[e.RowIndex];
                // ���� ǥ�ø� �������� ������
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
                        UpdateStatus($"������ ���� �Ϸ�: {data.SampleNumber}");
                    }
                }
            }
            // ���� ��ư
            else if (e.ColumnIndex == 6)
            {
                if (MessageBox.Show("������ �����͸� �����Ͻðڽ��ϱ�?", "Ȯ��",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    measurements.RemoveAt(e.RowIndex);
                    // �ε��� ������
                    for (int i = 0; i < measurements.Count; i++)
                    {
                        measurements[i].Index = i + 1;
                    }
                    UpdateDataGridView();
                    UpdateStatus("������ ���� �Ϸ�");
                }
            }
        }

        // ���� �׸� ������ ��ųʸ�
        private readonly Dictionary<MeasurementType, string> measurementPrefixes = new Dictionary<MeasurementType, string>
        {
            { MeasurementType.InitialPlate, "IP" },
            { MeasurementType.InitialPlateSample, "IPS" },
            { MeasurementType.FinalPlateSample, "FPS" },
            { MeasurementType.FinalKeepPlateSample, "FKPS" },
            { MeasurementType.InitialCollector, "IC" },
            { MeasurementType.FinalCollector, "FC" }
        };

        // ���� �׸� ǥ�� �̸�
        private readonly Dictionary<MeasurementType, string> measurementTypeNames = new Dictionary<MeasurementType, string>
        {
            { MeasurementType.InitialPlate, "�ʱ� ����" },
            { MeasurementType.InitialPlateSample, "�ʱ� ���� + �÷�" },
            { MeasurementType.FinalPlateSample, "���� ���� + �÷�" },
            { MeasurementType.FinalKeepPlateSample, "���� ���� ���� + �÷�" },
            { MeasurementType.InitialCollector, "���� ������" },
            { MeasurementType.FinalCollector, "���� ������" }
        };

        private TcpClient client;
        private NetworkStream stream;
        private bool isConnected = false;
        private List<MeasurementData> measurements;
        private double currentWeight = 0;
        private string currentUnit = "g";
        private bool isStable = false;

        // �ڵ� ���� ���� �ʵ�
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
            UpdateStatus("���α׷��� ���۵Ǿ����ϴ�.");

            // ī�װ� �޺��ڽ� �ʱ�ȭ
            InitializeCategoryComboBox();

            // ���� �ð� ���� �ʱ�ȭ
            nudStabilityTime.Value = 3;
            btnStartMeasurement.Enabled = true;

            // DataGridView �� ���� ����
            dgvMeasurements.RowsDefaultCellStyle.BackColor = Color.White;
            dgvMeasurements.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
        }

        // ī�װ� �޺��ڽ� �ʱ�ȭ
        private void InitializeCategoryComboBox()
        {
            cboCategory.Items.Clear();
            cboCategory.Items.Add("ī�װ� ����");
            cboCategory.Items.Add("Sample");
            cboCategory.Items.Add("Collector");
            cboCategory.SelectedIndex = 0;

            // ���� �׸� �޺��ڽ� �ʱ�ȭ
            cboMeasurementType.Items.Clear();
            cboMeasurementType.Enabled = false;
        }

        // ī�װ� ���� ���� �̺�Ʈ
        private void cboCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboMeasurementType.Items.Clear();
            cboMeasurementType.Enabled = false;

            // ���� ��ȣ �ʱ�ȭ
            txtSampleNumber.Text = "";

            if (cboCategory.SelectedIndex == 1) // Sample
            {
                cboMeasurementType.Items.Add("�׸� ����");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.InitialPlate]} ({measurementPrefixes[MeasurementType.InitialPlate]})");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.InitialPlateSample]} ({measurementPrefixes[MeasurementType.InitialPlateSample]})");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.FinalPlateSample]} ({measurementPrefixes[MeasurementType.FinalPlateSample]})");
                cboMeasurementType.Items.Add($"{measurementTypeNames[MeasurementType.FinalKeepPlateSample]} ({measurementPrefixes[MeasurementType.FinalKeepPlateSample]})");
                cboMeasurementType.Enabled = true;
                cboMeasurementType.SelectedIndex = 0;
            }
            else if (cboCategory.SelectedIndex == 2) // Collector
            {
                cboMeasurementType.Items.Add("�׸� ����");
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

                    // �ʱ� ��� ���� (@)
                    string response = await SendCommand("@");
                    if (!string.IsNullOrEmpty(response))
                    {
                        isConnected = true;
                        btnConnect.Text = "���� ����";
                        lblConnectionStatus.Text = "�����";
                        lblConnectionStatus.ForeColor = Color.Green;

                        // ���� �б� Ÿ�̸� ����
                        timerWeight.Start();

                        // �ʱ� ���� ���� Ȯ��
                        await CheckDoorStatusOnce();

                        UpdateStatus($"���� ���� ����: {response}");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"���� ����: {ex.Message}", "����",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatus($"���� ����: {ex.Message}");
                }
            }
            else
            {
                Disconnect();
            }
        }

        // ���� ���� �� �� Ȯ��
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
                UpdateStatus($"���� ���� Ȯ�� ����: {ex.Message}");
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
                btnConnect.Text = "����";
                lblConnectionStatus.Text = "���� ����";
                lblConnectionStatus.ForeColor = Color.Red;

                UpdateStatus("���� ������ ���������ϴ�.");
            }
            catch (Exception ex)
            {
                UpdateStatus($"���� ���� ����: {ex.Message}");
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
                UpdateStatus($"��� ���� ����: {ex.Message}");
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
                UpdateStatus($"���� �б� ����: {ex.Message}");
            }
        }

        private void ParseWeightResponse(string response)
        {
            try
            {
                // ���� ����: S S     1.2345678 g
                // ó�� 2�ڸ��� command ID, 3��°�� status, 4-14�� weight value, 15���ʹ� unit
                if (response.Length >= 14)
                {
                    string status = response.Substring(2, 1);
                    isStable = (status == "S");

                    // ���԰��� 4��° ���ں��� �����Ͽ� ���� ����
                    string weightStr = response.Substring(3).Trim();

                    // ���� �и�
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
                UpdateStatus($"���� �Ľ� ����: {ex.Message}");
            }
        }

        private void UpdateWeightDisplay()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(UpdateWeightDisplay));
                return;
            }

            // �Ҽ��� ���� �ڸ����� �������� ����
            string weightText = currentWeight.ToString("F10").TrimEnd('0');
            if (weightText.EndsWith("."))
                weightText += "0";

            lblWeight.Text = weightText;
            lblUnit.Text = currentUnit;
            lblStability.Text = isStable ? "����" : "�Ҿ���";
            lblStability.ForeColor = isStable ? Color.Lime : Color.Yellow;
        }

        private async void btnZero_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("������ ������� �ʾҽ��ϴ�.", "����",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string response = await SendCommand("Z");
            if (response.Contains("A"))
            {
                UpdateStatus("���� ���� �Ϸ�");
            }
            else
            {
                UpdateStatus($"���� ���� ����: {response}");
            }
        }

        private async void btnTare_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("������ ������� �ʾҽ��ϴ�.", "����",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string response = await SendCommand("T");
            if (response.Contains("S"))
            {
                UpdateStatus("��⹫�� ���� �Ϸ�");
            }
            else
            {
                UpdateStatus($"��⹫�� ���� ����: {response}");
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("������ ������� �ʾҽ��ϴ�.", "����",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ī�װ� ���� Ȯ��
            if (cboCategory.SelectedIndex <= 0)
            {
                MessageBox.Show("ī�װ��� �����ϼ���.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboCategory.Focus();
                return;
            }

            // ���� �׸� ���� Ȯ��
            if (cboMeasurementType.SelectedIndex <= 0)
            {
                MessageBox.Show("���� �׸��� �����ϼ���.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboMeasurementType.Focus();
                return;
            }

            // ���� ��ȣ Ȯ��
            string sampleNumber = txtSampleNumber.Text.Trim();
            if (string.IsNullOrEmpty(sampleNumber))
            {
                MessageBox.Show("���� ��ȣ�� �Է��ϼ���.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSampleNumber.Focus();
                return;
            }

            // ���� �׸� ��������
            string measurementType = cboMeasurementType.SelectedItem.ToString();

            // ������ ����
            measurements.Add(new MeasurementData
            {
                Index = measurements.Count + 1,
                SampleNumber = sampleNumber,
                Weight = currentWeight,
                Unit = currentUnit,
                DateTime = DateTime.Now,
                MeasurementType = measurementType
            });

            // DataGridView ������Ʈ
            UpdateDataGridView();

            // ���� ��ȣ �ڵ� ����
            IncrementSampleNumberWithPrefix();

            // ���� ������Ʈ
            UpdateStatus($"���� �Ϸ�: {sampleNumber} - {currentWeight:F7} {currentUnit} ({measurementType})");
        }

        // �޺��ڽ� ���� ���� �̺�Ʈ �ڵ鷯
        private void cboMeasurementType_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateSampleNumberWithPrefix();
        }

        // ���� ��ȣ�� ������ �߰��ϴ� �޼���
        private void UpdateSampleNumberWithPrefix()
        {
            if (cboMeasurementType.SelectedIndex <= 0)
                return;

            // ���� ���� ��ȣ
            string currentSampleNumber = txtSampleNumber.Text.Trim();

            // ���� ������ ���� (�ִ� ���)
            foreach (var prefix in measurementPrefixes.Values)
            {
                if (currentSampleNumber.StartsWith(prefix + "_"))
                {
                    currentSampleNumber = currentSampleNumber.Substring(prefix.Length + 1);
                    break;
                }
            }

            // ī�װ����� ���õ� �׸��� ���� �ε��� ���
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

            // ���� ��ȣ�� ��������� �ڵ� ��ȣ ����
            if (string.IsNullOrWhiteSpace(currentSampleNumber))
            {
                int count = measurements.Count(m => m.SampleNumber.StartsWith(newPrefix + "_")) + 1;
                currentSampleNumber = count.ToString("D3");
            }

            txtSampleNumber.Text = $"{newPrefix}_{currentSampleNumber}";
        }

        // ���� ��ȣ �ڵ� ���� �޼���
        private void IncrementSampleNumberWithPrefix()
        {
            string currentNumber = txtSampleNumber.Text;
            if (string.IsNullOrWhiteSpace(currentNumber))
                return;

            // �����ڿ� ��ȣ �и�
            int underscoreIndex = currentNumber.IndexOf('_');
            if (underscoreIndex > 0)
            {
                string prefix = currentNumber.Substring(0, underscoreIndex);
                string numberPart = currentNumber.Substring(underscoreIndex + 1);

                // ���� �κи� ����
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

        // ���� ����/���� ��ư Ŭ�� �̺�Ʈ
        private void btnStartMeasurement_Click(object sender, EventArgs e)
        {
            if (!isConnected)
            {
                MessageBox.Show("������ ������� �ʾҽ��ϴ�.", "����",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ī�װ��� ���� �׸� Ȯ��
            if (cboCategory.SelectedIndex <= 0)
            {
                MessageBox.Show("ī�װ��� �����ϼ���.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboCategory.Focus();
                return;
            }

            if (cboMeasurementType.SelectedIndex <= 0)
            {
                MessageBox.Show("���� �׸��� �����ϼ���.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboMeasurementType.Focus();
                return;
            }

            if (!isAutoMeasureEnabled)
            {
                // ���� ����
                isAutoMeasureEnabled = true;
                stabilityTimeSeconds = (int)nudStabilityTime.Value;
                timerDoorCheck.Start();
                timerAutoSave.Start();

                btnStartMeasurement.Text = "���� ����";
                btnStartMeasurement.BackColor = Color.Red;

                // ��Ʈ�� ��Ȱ��ȭ
                cboCategory.Enabled = false;
                cboMeasurementType.Enabled = false;
                nudStabilityTime.Enabled = false;

                UpdateStatus("�ڵ� ���� ��� ���� - ��� ���� ������ ��ü�ϼ���.");
            }
            else
            {
                // ���� ����
                isAutoMeasureEnabled = false;
                timerDoorCheck.Stop();
                timerAutoSave.Stop();

                btnStartMeasurement.Text = "���� ����";
                btnStartMeasurement.BackColor = Color.FromArgb(0, 192, 0);

                // ��Ʈ�� Ȱ��ȭ
                cboCategory.Enabled = true;
                cboMeasurementType.Enabled = true;
                nudStabilityTime.Enabled = true;

                UpdateStatus("�ڵ� ���� ��� ����");
            }
        }

        // ���� ���� Ȯ�� Ÿ�̸�
        private async void timerDoorCheck_Tick(object sender, EventArgs e)
        {
            if (!isConnected || !isAutoMeasureEnabled)
                return;

            try
            {
                // WS ������� ���� ���� Ȯ��
                string response = await SendCommand("WS");
                if (!string.IsNullOrEmpty(response) && response.StartsWith("WS "))
                {
                    // WS 0: ���� ����, WS 1-7: ���� ����
                    int newDoorStatus = int.Parse(response.Substring(5, 1));

                    if (newDoorStatus != doorStatus)
                    {
                        doorStatus = newDoorStatus;
                        isDoorOpen = doorStatus > 0;

                        // ���� ���� ǥ�� ������Ʈ
                        UpdateDoorStatus(isDoorOpen);

                        // ��� ������ ��
                        if (!isDoorOpen && wasDoorOpen)
                        {
                            doorClosedTime = DateTime.Now;
                            isWaitingForStability = true;
                            UpdateStatus("���� ���� - ���� ���� ��� ��...");
                        }

                        wasDoorOpen = isDoorOpen;
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"���� ���� Ȯ�� ����: {ex.Message}");
            }
        }

        // ���� ���� ǥ�� ������Ʈ
        private void UpdateDoorStatus(bool isOpen)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<bool>(UpdateDoorStatus), isOpen);
                return;
            }

            lblDoorStatus.Text = isOpen ? "����: ����" : "����: ����";
            lblDoorStatus.ForeColor = isOpen ? Color.Red : Color.Blue;
        }

        // �ڵ� ���� Ÿ�̸� �̺�Ʈ
        private void timerAutoSave_Tick(object sender, EventArgs e)
        {
            if (!isConnected || !isAutoMeasureEnabled || !isWaitingForStability)
                return;

            try
            {
                // ��� �����ְ�, ���԰� �����Ǿ����� Ȯ��
                if (!isDoorOpen && isStable && currentWeight > 0.1)
                {
                    // ��� ���� �� ������ �ð��� ����ߴ��� Ȯ��
                    TimeSpan timeSinceDoorClosed = DateTime.Now - doorClosedTime;

                    if (timeSinceDoorClosed.TotalSeconds >= stabilityTimeSeconds)
                    {
                        // �ڵ� ���� ����
                        PerformAutoSave();
                        isWaitingForStability = false;
                        UpdateStatus("�ڵ� ���� �Ϸ� - ��� ���� ���� ������ �غ��ϼ���.");
                    }
                    else
                    {
                        int remainingTime = stabilityTimeSeconds - (int)timeSinceDoorClosed.TotalSeconds;
                        UpdateStatus($"���� ��� ��... {remainingTime}��");
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"�ڵ� ���� ����: {ex.Message}");
            }
        }

        // �ڵ� ���� ���� �޼���
        private void PerformAutoSave()
        {
            // ���� ��ȣ �ڵ� ����
            string autoSampleNumber = txtSampleNumber.Text;
            if (string.IsNullOrWhiteSpace(autoSampleNumber))
            {
                // ���� ���õ� ���� Ÿ�� ��������
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

            // ���� �׸� ��������
            string measurementType = cboMeasurementType.SelectedItem.ToString();

            // ������ ����
            measurements.Add(new MeasurementData
            {
                Index = measurements.Count + 1,
                SampleNumber = autoSampleNumber,
                Weight = currentWeight,
                Unit = currentUnit,
                DateTime = DateTime.Now,
                MeasurementType = measurementType
            });

            // DataGridView ������Ʈ
            UpdateDataGridView();

            // ���� ������Ʈ
            UpdateStatus($"�ڵ� ���� �Ϸ�: {autoSampleNumber} - {currentWeight:F7} {currentUnit} ({measurementType})");

            // ���� ��ȣ �ڵ� ����
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
                // ���� ǥ�ø� �������� ������
                string weightText = data.Weight.ToString("F10").TrimEnd('0');
                if (weightText.EndsWith("."))
                    weightText += "0";

                int rowIndex = dgvMeasurements.Rows.Add(
                    data.Index,
                    data.SampleNumber,
                    weightText,
                    data.Unit,
                    data.DateTime.ToString("yyyy-MM-dd HH:mm:ss"),
                    "����",
                    "����"
                );

                // ���� �׸��� �������� ǥ��
                if (!string.IsNullOrEmpty(data.MeasurementType))
                {
                    dgvMeasurements.Rows[rowIndex].Cells[1].ToolTipText = data.MeasurementType;
                }
            }
        }

        // DataGridView �� ����Ŭ�� �̺�Ʈ - ���� ����
        private void dgvMeasurements_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // ���� ��ȣ�� ���� ���� ���� ����
            if (e.ColumnIndex == 1 || e.ColumnIndex == 2)
            {
                dgvMeasurements.BeginEdit(true);
            }
        }

        // DataGridView �� ���� �Ϸ� �̺�Ʈ
        private void dgvMeasurements_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            try
            {
                var data = measurements[e.RowIndex];

                // ���� ��ȣ ����
                if (e.ColumnIndex == 1)
                {
                    string newSampleNumber = dgvMeasurements.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                    if (!string.IsNullOrWhiteSpace(newSampleNumber))
                    {
                        data.SampleNumber = newSampleNumber;
                        UpdateStatus($"���� ��ȣ ���� �Ϸ�: {newSampleNumber}");
                    }
                    else
                    {
                        // �� ���̸� ���� ������ ����
                        dgvMeasurements.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = data.SampleNumber;
                    }
                }
                // ���� ����
                else if (e.ColumnIndex == 2)
                {
                    string weightStr = dgvMeasurements.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                    if (!string.IsNullOrWhiteSpace(weightStr) &&
                        double.TryParse(weightStr, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out double newWeight))
                    {
                        data.Weight = newWeight;
                        UpdateStatus($"���� ���� �Ϸ�: {newWeight:F7} {data.Unit}");
                    }
                    else
                    {
                        // ��ȿ���� ���� ���̸� ���� ������ ����
                        string weightText = data.Weight.ToString("F10").TrimEnd('0');
                        if (weightText.EndsWith("."))
                            weightText += "0";
                        dgvMeasurements.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = weightText;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"������ ���� �� ���� �߻�: {ex.Message}", "����",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateDataGridView();
            }
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            if (measurements.Count == 0)
            {
                MessageBox.Show("������ �����Ͱ� �����ϴ�.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show($"��ü {measurements.Count}���� �����͸� �����Ͻðڽ��ϱ�?", "Ȯ��",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                measurements.Clear();
                UpdateDataGridView();
                UpdateStatus("��ü ������ ���� �Ϸ�");
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (measurements.Count == 0)
            {
                MessageBox.Show("������ �����Ͱ� �����ϴ�.", "�˸�",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Files|*.xlsx";
                sfd.FileName = $"����������_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var package = new ExcelPackage())
                        {
                            var worksheet = package.Workbook.Worksheets.Add("���� ������");

                            // ��� �߰�
                            worksheet.Cells[1, 1].Value = "����";
                            worksheet.Cells[1, 2].Value = "���� ��ȣ";
                            worksheet.Cells[1, 3].Value = "����";
                            worksheet.Cells[1, 4].Value = "����";
                            worksheet.Cells[1, 5].Value = "�����ð�";
                            worksheet.Cells[1, 6].Value = "�����׸�";

                            // ��� ��Ÿ��
                            using (var range = worksheet.Cells[1, 1, 1, 6])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            }

                            // ������ �߰�
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

                            // �� �ʺ� �ڵ� ����
                            worksheet.Cells.AutoFitColumns();

                            // ���� ����
                            package.SaveAs(new System.IO.FileInfo(sfd.FileName));
                        }

                        MessageBox.Show("���� ���Ϸ� ����Ǿ����ϴ�.", "�Ϸ�",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdateStatus($"���� ���� ���� �Ϸ�: {sfd.FileName}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"���� �� ���� �߻�: {ex.Message}", "����",
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

                            // ���� �����͸� ����� Ȯ��
                            if (measurements.Count > 0)
                            {
                                var result = MessageBox.Show("���� �����Ϳ� �߰��Ͻðڽ��ϱ�?\n" +
                                    "�ƴϿ��� �����ϸ� ���� �����Ͱ� �����˴ϴ�.", "Ȯ��",
                                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                if (result == DialogResult.Cancel)
                                    return;
                                else if (result == DialogResult.No)
                                    measurements.Clear();
                            }

                            // ������ �б� (��� �� ����)
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
                            MessageBox.Show($"{rowCount - 1}���� �����͸� �ҷ��Խ��ϴ�.", "�Ϸ�",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            UpdateStatus($"���� ���� �ҷ����� �Ϸ�: {ofd.FileName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"�ҷ����� �� ���� �߻�: {ex.Message}", "����",
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
            // �ڵ� ���� ���̸� ����
            if (isAutoMeasureEnabled)
            {
                var result = MessageBox.Show("�ڵ� ������ ���� ���Դϴ�. �����Ͻðڽ��ϱ�?",
                    "Ȯ��", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

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
                var result = MessageBox.Show("������� ���� �����Ͱ� �ֽ��ϴ�. �����Ͻðڽ��ϱ�?",
                    "Ȯ��", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

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

    // ���� ������ Ŭ����
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