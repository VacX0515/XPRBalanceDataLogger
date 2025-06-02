using System;
using System.Windows.Forms;

namespace XPRBalanceDataLogger
{
    public partial class EditDataForm : Form
    {
        public string SampleNumber { get; private set; }
        public string Weight { get; private set; }

        public EditDataForm(string sampleNumber, string weight)
        {
            InitializeComponent();
            txtSampleNumber.Text = sampleNumber;
            txtWeight.Text = weight;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSampleNumber.Text))
            {
                MessageBox.Show("샘플 번호를 입력하세요.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSampleNumber.Focus();
                return;
            }

            if (!IsValidWeight(txtWeight.Text))
            {
                MessageBox.Show("올바른 무게 값을 입력하세요.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtWeight.Focus();
                return;
            }

            SampleNumber = txtSampleNumber.Text.Trim();
            Weight = txtWeight.Text.Trim();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private bool IsValidWeight(string weight)
        {
            if (string.IsNullOrWhiteSpace(weight))
                return false;

            return double.TryParse(weight, out double result) && result >= 0;
        }

        private void EditDataForm_Load(object sender, EventArgs e)
        {
            txtSampleNumber.SelectAll();
            txtSampleNumber.Focus();
        }
    }
}