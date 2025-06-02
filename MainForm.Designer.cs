namespace XPRBalanceDataLogger
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            pnlConnection = new Panel();
            lblConnectionStatus = new Label();
            txtPort = new TextBox();
            lblPort = new Label();
            txtIPAddress = new TextBox();
            lblIPAddress = new Label();
            btnConnect = new Button();
            pnlWeight = new Panel();
            lblStability = new Label();
            lblUnit = new Label();
            lblWeight = new Label();
            pnlControl = new Panel();
            cboCategory = new ComboBox();
            lblCategory = new Label();
            cboMeasurementType = new ComboBox();
            lblMeasurementType = new Label();
            btnStartMeasurement = new Button();
            lblDoorStatus = new Label();
            nudStabilityTime = new NumericUpDown();
            lblStabilityTime = new Label();
            btnTare = new Button();
            btnZero = new Button();
            btnSave = new Button();
            txtSampleNumber = new TextBox();
            lblSampleNumber = new Label();
            pnlData = new Panel();
            dgvMeasurements = new DataGridView();
            colIndex = new DataGridViewTextBoxColumn();
            colSampleNumber = new DataGridViewTextBoxColumn();
            colWeight = new DataGridViewTextBoxColumn();
            colUnit = new DataGridViewTextBoxColumn();
            colDateTime = new DataGridViewTextBoxColumn();
            colEdit = new DataGridViewButtonColumn();
            colDelete = new DataGridViewButtonColumn();
            pnlDataControl = new Panel();
            btnImportExcel = new Button();
            btnExportExcel = new Button();
            btnClearAll = new Button();
            statusStrip = new StatusStrip();
            toolStripStatusLabel = new ToolStripStatusLabel();
            timerWeight = new System.Windows.Forms.Timer(components);
            timerAutoSave = new System.Windows.Forms.Timer(components);
            timerDoorCheck = new System.Windows.Forms.Timer(components);
            pnlConnection.SuspendLayout();
            pnlWeight.SuspendLayout();
            pnlControl.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)nudStabilityTime).BeginInit();
            pnlData.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvMeasurements).BeginInit();
            pnlDataControl.SuspendLayout();
            statusStrip.SuspendLayout();
            SuspendLayout();
            // 
            // pnlConnection
            // 
            pnlConnection.BorderStyle = BorderStyle.FixedSingle;
            pnlConnection.Controls.Add(lblConnectionStatus);
            pnlConnection.Controls.Add(txtPort);
            pnlConnection.Controls.Add(lblPort);
            pnlConnection.Controls.Add(txtIPAddress);
            pnlConnection.Controls.Add(lblIPAddress);
            pnlConnection.Controls.Add(btnConnect);
            pnlConnection.Dock = DockStyle.Top;
            pnlConnection.Location = new Point(0, 0);
            pnlConnection.Margin = new Padding(3, 4, 3, 4);
            pnlConnection.Name = "pnlConnection";
            pnlConnection.Size = new Size(1400, 75);
            pnlConnection.TabIndex = 0;
            // 
            // lblConnectionStatus
            // 
            lblConnectionStatus.AutoSize = true;
            lblConnectionStatus.Font = new Font("맑은 고딕", 12F, FontStyle.Bold, GraphicsUnit.Point, 129);
            lblConnectionStatus.ForeColor = Color.Red;
            lblConnectionStatus.Location = new Point(500, 28);
            lblConnectionStatus.Name = "lblConnectionStatus";
            lblConnectionStatus.Size = new Size(80, 21);
            lblConnectionStatus.TabIndex = 5;
            lblConnectionStatus.Text = "연결 끊김";
            lblConnectionStatus.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // txtPort
            // 
            txtPort.Location = new Point(260, 25);
            txtPort.Margin = new Padding(3, 4, 3, 4);
            txtPort.Name = "txtPort";
            txtPort.Size = new Size(60, 23);
            txtPort.TabIndex = 4;
            txtPort.Text = "8001";
            // 
            // lblPort
            // 
            lblPort.AutoSize = true;
            lblPort.Location = new Point(220, 29);
            lblPort.Name = "lblPort";
            lblPort.Size = new Size(34, 15);
            lblPort.TabIndex = 3;
            lblPort.Text = "포트:";
            // 
            // txtIPAddress
            // 
            txtIPAddress.Location = new Point(80, 25);
            txtIPAddress.Margin = new Padding(3, 4, 3, 4);
            txtIPAddress.Name = "txtIPAddress";
            txtIPAddress.Size = new Size(120, 23);
            txtIPAddress.TabIndex = 2;
            txtIPAddress.Text = "192.168.1.100";
            // 
            // lblIPAddress
            // 
            lblIPAddress.AutoSize = true;
            lblIPAddress.Location = new Point(20, 29);
            lblIPAddress.Name = "lblIPAddress";
            lblIPAddress.Size = new Size(48, 15);
            lblIPAddress.TabIndex = 1;
            lblIPAddress.Text = "IP 주소:";
            // 
            // btnConnect
            // 
            btnConnect.Location = new Point(350, 23);
            btnConnect.Margin = new Padding(3, 4, 3, 4);
            btnConnect.Name = "btnConnect";
            btnConnect.Size = new Size(100, 31);
            btnConnect.TabIndex = 0;
            btnConnect.Text = "연결";
            btnConnect.UseVisualStyleBackColor = true;
            btnConnect.Click += btnConnect_Click;
            // 
            // pnlWeight
            // 
            pnlWeight.BackColor = Color.FromArgb(64, 64, 64);
            pnlWeight.BorderStyle = BorderStyle.FixedSingle;
            pnlWeight.Controls.Add(lblStability);
            pnlWeight.Controls.Add(lblUnit);
            pnlWeight.Controls.Add(lblWeight);
            pnlWeight.Location = new Point(12, 88);
            pnlWeight.Margin = new Padding(3, 4, 3, 4);
            pnlWeight.Name = "pnlWeight";
            pnlWeight.Size = new Size(460, 150);
            pnlWeight.TabIndex = 1;
            // 
            // lblStability
            // 
            lblStability.AutoSize = true;
            lblStability.FlatStyle = FlatStyle.System;
            lblStability.Font = new Font("맑은 고딕", 16F, FontStyle.Bold, GraphicsUnit.Point, 129);
            lblStability.ForeColor = Color.Yellow;
            lblStability.Location = new Point(199, 9);
            lblStability.Name = "lblStability";
            lblStability.Size = new Size(57, 30);
            lblStability.TabIndex = 2;
            lblStability.Text = "안정";
            // 
            // lblUnit
            // 
            lblUnit.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            lblUnit.AutoSize = true;
            lblUnit.Font = new Font("맑은 고딕", 24F, FontStyle.Bold);
            lblUnit.ForeColor = Color.White;
            lblUnit.Location = new Point(407, 83);
            lblUnit.Name = "lblUnit";
            lblUnit.Size = new Size(40, 45);
            lblUnit.TabIndex = 1;
            lblUnit.Text = "g";
            lblUnit.TextAlign = ContentAlignment.MiddleCenter;
            lblUnit.Click += lblUnit_Click;
            // 
            // lblWeight
            // 
            lblWeight.Font = new Font("맑은 고딕", 36F, FontStyle.Bold);
            lblWeight.ForeColor = Color.Lime;
            lblWeight.Location = new Point(10, 50);
            lblWeight.Name = "lblWeight";
            lblWeight.Size = new Size(391, 86);
            lblWeight.TabIndex = 0;
            lblWeight.Text = "0.0000000";
            lblWeight.TextAlign = ContentAlignment.MiddleRight;
            // 
            // pnlControl
            // 
            pnlControl.BorderStyle = BorderStyle.FixedSingle;
            pnlControl.Controls.Add(cboCategory);
            pnlControl.Controls.Add(lblCategory);
            pnlControl.Controls.Add(cboMeasurementType);
            pnlControl.Controls.Add(lblMeasurementType);
            pnlControl.Controls.Add(btnStartMeasurement);
            pnlControl.Controls.Add(lblDoorStatus);
            pnlControl.Controls.Add(nudStabilityTime);
            pnlControl.Controls.Add(lblStabilityTime);
            pnlControl.Controls.Add(btnTare);
            pnlControl.Controls.Add(btnZero);
            pnlControl.Controls.Add(btnSave);
            pnlControl.Controls.Add(txtSampleNumber);
            pnlControl.Controls.Add(lblSampleNumber);
            pnlControl.Location = new Point(12, 250);
            pnlControl.Margin = new Padding(3, 4, 3, 4);
            pnlControl.Name = "pnlControl";
            pnlControl.Size = new Size(460, 250);
            pnlControl.TabIndex = 2;
            // 
            // cboCategory
            // 
            cboCategory.DropDownStyle = ComboBoxStyle.DropDownList;
            cboCategory.Font = new Font("맑은 고딕", 10F);
            cboCategory.Location = new Point(90, 2);
            cboCategory.Name = "cboCategory";
            cboCategory.Size = new Size(200, 25);
            cboCategory.TabIndex = 11;
            cboCategory.SelectedIndexChanged += cboCategory_SelectedIndexChanged;
            // 
            // lblCategory
            // 
            lblCategory.AutoSize = true;
            lblCategory.Font = new Font("맑은 고딕", 10F);
            lblCategory.Location = new Point(10, 5);
            lblCategory.Name = "lblCategory";
            lblCategory.Size = new Size(40, 19);
            lblCategory.TabIndex = 10;
            lblCategory.Text = "구분:";
            // 
            // cboMeasurementType
            // 
            cboMeasurementType.DropDownStyle = ComboBoxStyle.DropDownList;
            cboMeasurementType.Font = new Font("맑은 고딕", 10F);
            cboMeasurementType.Location = new Point(90, 32);
            cboMeasurementType.Name = "cboMeasurementType";
            cboMeasurementType.Size = new Size(200, 25);
            cboMeasurementType.TabIndex = 9;
            cboMeasurementType.SelectedIndexChanged += cboMeasurementType_SelectedIndexChanged;
            // 
            // lblMeasurementType
            // 
            lblMeasurementType.AutoSize = true;
            lblMeasurementType.Font = new Font("맑은 고딕", 10F);
            lblMeasurementType.Location = new Point(10, 35);
            lblMeasurementType.Name = "lblMeasurementType";
            lblMeasurementType.Size = new Size(73, 19);
            lblMeasurementType.TabIndex = 8;
            lblMeasurementType.Text = "측정 항목:";
            // 
            // btnStartMeasurement
            // 
            btnStartMeasurement.BackColor = Color.FromArgb(0, 192, 0);
            btnStartMeasurement.Font = new Font("맑은 고딕", 12F, FontStyle.Bold);
            btnStartMeasurement.ForeColor = Color.White;
            btnStartMeasurement.Location = new Point(300, 2);
            btnStartMeasurement.Name = "btnStartMeasurement";
            btnStartMeasurement.Size = new Size(147, 55);
            btnStartMeasurement.TabIndex = 14;
            btnStartMeasurement.Text = "측정 시작";
            btnStartMeasurement.UseVisualStyleBackColor = false;
            btnStartMeasurement.Click += btnStartMeasurement_Click;
            // 
            // lblDoorStatus
            // 
            lblDoorStatus.AutoSize = true;
            lblDoorStatus.Font = new Font("맑은 고딕", 12F, FontStyle.Bold);
            lblDoorStatus.ForeColor = Color.Blue;
            lblDoorStatus.Location = new Point(10, 185);
            lblDoorStatus.Name = "lblDoorStatus";
            lblDoorStatus.Size = new Size(84, 21);
            lblDoorStatus.TabIndex = 15;
            lblDoorStatus.Text = "도어: 닫힘";
            // 
            // nudStabilityTime
            // 
            nudStabilityTime.Location = new Point(110, 213);
            nudStabilityTime.Maximum = new decimal(new int[] { 10, 0, 0, 0 });
            nudStabilityTime.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            nudStabilityTime.Name = "nudStabilityTime";
            nudStabilityTime.Size = new Size(50, 23);
            nudStabilityTime.TabIndex = 17;
            nudStabilityTime.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // lblStabilityTime
            // 
            lblStabilityTime.AutoSize = true;
            lblStabilityTime.Location = new Point(10, 215);
            lblStabilityTime.Name = "lblStabilityTime";
            lblStabilityTime.Size = new Size(87, 15);
            lblStabilityTime.TabIndex = 16;
            lblStabilityTime.Text = "안정화 시간(s):";
            // 
            // btnTare
            // 
            btnTare.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            btnTare.Location = new Point(210, 135);
            btnTare.Margin = new Padding(3, 4, 3, 4);
            btnTare.Name = "btnTare";
            btnTare.Size = new Size(80, 37);
            btnTare.TabIndex = 4;
            btnTare.Text = "용기무게";
            btnTare.UseVisualStyleBackColor = true;
            btnTare.Click += btnTare_Click;
            // 
            // btnZero
            // 
            btnZero.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            btnZero.Location = new Point(120, 135);
            btnZero.Margin = new Padding(3, 4, 3, 4);
            btnZero.Name = "btnZero";
            btnZero.Size = new Size(80, 37);
            btnZero.TabIndex = 3;
            btnZero.Text = "영점";
            btnZero.UseVisualStyleBackColor = true;
            btnZero.Click += btnZero_Click;
            // 
            // btnSave
            // 
            btnSave.BackColor = Color.FromArgb(0, 122, 204);
            btnSave.Font = new Font("맑은 고딕", 12F, FontStyle.Bold);
            btnSave.ForeColor = Color.White;
            btnSave.Location = new Point(300, 72);
            btnSave.Margin = new Padding(3, 4, 3, 4);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(147, 100);
            btnSave.TabIndex = 2;
            btnSave.Text = "수동 저장";
            btnSave.UseVisualStyleBackColor = false;
            btnSave.Click += btnSave_Click;
            // 
            // txtSampleNumber
            // 
            txtSampleNumber.Font = new Font("맑은 고딕", 12F);
            txtSampleNumber.Location = new Point(120, 85);
            txtSampleNumber.Margin = new Padding(3, 4, 3, 4);
            txtSampleNumber.Name = "txtSampleNumber";
            txtSampleNumber.Size = new Size(170, 29);
            txtSampleNumber.TabIndex = 1;
            // 
            // lblSampleNumber
            // 
            lblSampleNumber.AutoSize = true;
            lblSampleNumber.Font = new Font("맑은 고딕", 10F);
            lblSampleNumber.Location = new Point(10, 91);
            lblSampleNumber.Name = "lblSampleNumber";
            lblSampleNumber.Size = new Size(73, 19);
            lblSampleNumber.TabIndex = 0;
            lblSampleNumber.Text = "샘플 번호:";
            // 
            // pnlData
            // 
            pnlData.BorderStyle = BorderStyle.FixedSingle;
            pnlData.Controls.Add(dgvMeasurements);
            pnlData.Location = new Point(500, 88);
            pnlData.Margin = new Padding(3, 4, 3, 4);
            pnlData.Name = "pnlData";
            pnlData.Size = new Size(888, 700);
            pnlData.TabIndex = 3;
            // 
            // dgvMeasurements
            // 
            dgvMeasurements.AllowUserToAddRows = false;
            dgvMeasurements.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvMeasurements.Columns.AddRange(new DataGridViewColumn[] { colIndex, colSampleNumber, colWeight, colUnit, colDateTime, colEdit, colDelete });
            dgvMeasurements.Dock = DockStyle.Fill;
            dgvMeasurements.Location = new Point(0, 0);
            dgvMeasurements.Margin = new Padding(3, 4, 3, 4);
            dgvMeasurements.Name = "dgvMeasurements";
            dgvMeasurements.RowHeadersWidth = 62;
            dgvMeasurements.RowTemplate.Height = 23;
            dgvMeasurements.Size = new Size(886, 698);
            dgvMeasurements.TabIndex = 0;
            dgvMeasurements.CellClick += dgvMeasurements_CellClick;
            dgvMeasurements.CellDoubleClick += dgvMeasurements_CellDoubleClick;
            dgvMeasurements.CellEndEdit += dgvMeasurements_CellEndEdit;
            // 
            // colIndex
            // 
            colIndex.HeaderText = "순번";
            colIndex.MinimumWidth = 8;
            colIndex.Name = "colIndex";
            colIndex.ReadOnly = true;
            colIndex.Width = 60;
            // 
            // colSampleNumber
            // 
            colSampleNumber.HeaderText = "샘플 번호";
            colSampleNumber.MinimumWidth = 8;
            colSampleNumber.Name = "colSampleNumber";
            colSampleNumber.Width = 150;
            // 
            // colWeight
            // 
            colWeight.HeaderText = "무게";
            colWeight.MinimumWidth = 8;
            colWeight.Name = "colWeight";
            colWeight.Width = 180;
            // 
            // colUnit
            // 
            colUnit.HeaderText = "단위";
            colUnit.MinimumWidth = 8;
            colUnit.Name = "colUnit";
            colUnit.ReadOnly = true;
            colUnit.Width = 60;
            // 
            // colDateTime
            // 
            colDateTime.HeaderText = "측정시간";
            colDateTime.MinimumWidth = 8;
            colDateTime.Name = "colDateTime";
            colDateTime.ReadOnly = true;
            colDateTime.Width = 180;
            // 
            // colEdit
            // 
            colEdit.HeaderText = "수정";
            colEdit.MinimumWidth = 8;
            colEdit.Name = "colEdit";
            colEdit.Text = "수정";
            colEdit.UseColumnTextForButtonValue = true;
            colEdit.Width = 60;
            // 
            // colDelete
            // 
            colDelete.HeaderText = "삭제";
            colDelete.MinimumWidth = 8;
            colDelete.Name = "colDelete";
            colDelete.Text = "삭제";
            colDelete.UseColumnTextForButtonValue = true;
            colDelete.Width = 60;
            // 
            // pnlDataControl
            // 
            pnlDataControl.BorderStyle = BorderStyle.FixedSingle;
            pnlDataControl.Controls.Add(btnImportExcel);
            pnlDataControl.Controls.Add(btnExportExcel);
            pnlDataControl.Controls.Add(btnClearAll);
            pnlDataControl.Location = new Point(12, 508);
            pnlDataControl.Margin = new Padding(3, 4, 3, 4);
            pnlDataControl.Name = "pnlDataControl";
            pnlDataControl.Size = new Size(460, 75);
            pnlDataControl.TabIndex = 4;
            // 
            // btnImportExcel
            // 
            btnImportExcel.Location = new Point(327, 19);
            btnImportExcel.Margin = new Padding(3, 4, 3, 4);
            btnImportExcel.Name = "btnImportExcel";
            btnImportExcel.Size = new Size(120, 37);
            btnImportExcel.TabIndex = 2;
            btnImportExcel.Text = "엑셀 불러오기";
            btnImportExcel.UseVisualStyleBackColor = true;
            btnImportExcel.Click += btnImportExcel_Click;
            // 
            // btnExportExcel
            // 
            btnExportExcel.Location = new Point(179, 19);
            btnExportExcel.Margin = new Padding(3, 4, 3, 4);
            btnExportExcel.Name = "btnExportExcel";
            btnExportExcel.Size = new Size(120, 37);
            btnExportExcel.TabIndex = 1;
            btnExportExcel.Text = "엑셀로 저장";
            btnExportExcel.UseVisualStyleBackColor = true;
            btnExportExcel.Click += btnExportExcel_Click;
            // 
            // btnClearAll
            // 
            btnClearAll.Location = new Point(20, 19);
            btnClearAll.Margin = new Padding(3, 4, 3, 4);
            btnClearAll.Name = "btnClearAll";
            btnClearAll.Size = new Size(120, 37);
            btnClearAll.TabIndex = 0;
            btnClearAll.Text = "전체 삭제";
            btnClearAll.UseVisualStyleBackColor = true;
            btnClearAll.Click += btnClearAll_Click;
            // 
            // statusStrip
            // 
            statusStrip.ImageScalingSize = new Size(24, 24);
            statusStrip.Items.AddRange(new ToolStripItem[] { toolStripStatusLabel });
            statusStrip.Location = new Point(0, 615);
            statusStrip.Name = "statusStrip";
            statusStrip.Size = new Size(1400, 22);
            statusStrip.TabIndex = 5;
            statusStrip.Text = "statusStrip1";
            // 
            // toolStripStatusLabel
            // 
            toolStripStatusLabel.Name = "toolStripStatusLabel";
            toolStripStatusLabel.Size = new Size(56, 17);
            toolStripStatusLabel.Text = "준비 중...";
            // 
            // timerWeight
            // 
            timerWeight.Interval = 500;
            timerWeight.Tick += timerWeight_Tick;
            // 
            // timerAutoSave
            // 
            timerAutoSave.Interval = 1000;
            timerAutoSave.Tick += timerAutoSave_Tick;
            // 
            // timerDoorCheck
            // 
            timerDoorCheck.Interval = 500;
            timerDoorCheck.Tick += timerDoorCheck_Tick;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1400, 637);
            Controls.Add(statusStrip);
            Controls.Add(pnlDataControl);
            Controls.Add(pnlData);
            Controls.Add(pnlControl);
            Controls.Add(pnlWeight);
            Controls.Add(pnlConnection);
            Margin = new Padding(3, 4, 3, 4);
            Name = "MainForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "XPR 저울 데이터 로거";
            FormClosing += MainForm_FormClosing;
            Load += MainForm_Load;
            pnlConnection.ResumeLayout(false);
            pnlConnection.PerformLayout();
            pnlWeight.ResumeLayout(false);
            pnlWeight.PerformLayout();
            pnlControl.ResumeLayout(false);
            pnlControl.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)nudStabilityTime).EndInit();
            pnlData.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvMeasurements).EndInit();
            pnlDataControl.ResumeLayout(false);
            statusStrip.ResumeLayout(false);
            statusStrip.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Panel pnlConnection;
        private System.Windows.Forms.Label lblConnectionStatus;
        private System.Windows.Forms.TextBox txtPort;
        private System.Windows.Forms.Label lblPort;
        private System.Windows.Forms.TextBox txtIPAddress;
        private System.Windows.Forms.Label lblIPAddress;
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.Panel pnlWeight;
        private System.Windows.Forms.Label lblStability;
        private System.Windows.Forms.Label lblUnit;
        private System.Windows.Forms.Label lblWeight;
        private System.Windows.Forms.Panel pnlControl;
        private System.Windows.Forms.ComboBox cboCategory;
        private System.Windows.Forms.Label lblCategory;
        private System.Windows.Forms.ComboBox cboMeasurementType;
        private System.Windows.Forms.Label lblMeasurementType;
        private System.Windows.Forms.Button btnStartMeasurement;
        private System.Windows.Forms.Label lblDoorStatus;
        private System.Windows.Forms.NumericUpDown nudStabilityTime;
        private System.Windows.Forms.Label lblStabilityTime;
        private System.Windows.Forms.Button btnTare;
        private System.Windows.Forms.Button btnZero;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtSampleNumber;
        private System.Windows.Forms.Label lblSampleNumber;
        private System.Windows.Forms.Panel pnlData;
        private System.Windows.Forms.DataGridView dgvMeasurements;
        private System.Windows.Forms.DataGridViewTextBoxColumn colIndex;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSampleNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn colWeight;
        private System.Windows.Forms.DataGridViewTextBoxColumn colUnit;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDateTime;
        private System.Windows.Forms.DataGridViewButtonColumn colEdit;
        private System.Windows.Forms.DataGridViewButtonColumn colDelete;
        private System.Windows.Forms.Panel pnlDataControl;
        private System.Windows.Forms.Button btnImportExcel;
        private System.Windows.Forms.Button btnExportExcel;
        private System.Windows.Forms.Button btnClearAll;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.Timer timerWeight;
        private System.Windows.Forms.Timer timerAutoSave;
        private System.Windows.Forms.Timer timerDoorCheck;
    }
}