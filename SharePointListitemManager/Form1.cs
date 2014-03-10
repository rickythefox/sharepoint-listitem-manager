using System;
using System.Linq;
using System.Windows.Forms;

namespace SharePointListitemManager
{
    public partial class Form1 : Form
    {
        private const string Version = "0.1";
        private readonly ISharepointService _sharepointService;
        private readonly IExcelService _excelService;

        public Form1(ISharepointService sharepointService, IExcelService excelService)
        {
            _sharepointService = sharepointService;
            _excelService = excelService;

            InitializeComponent();

            txtSite.TextChanged += ChangeGuiState;
            cbList.SelectedIndexChanged += ChangeGuiState;
            Shown += ChangeGuiState;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.Text += Version;
            var siteUrls = _sharepointService.GetAllSiteUrls().ToArray();
            var aSiteNames = new AutoCompleteStringCollection();
            aSiteNames.AddRange(siteUrls);

            txtSite.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txtSite.AutoCompleteCustomSource = aSiteNames;
            txtSite.AutoCompleteMode = AutoCompleteMode.Suggest;
        }

        private void txtSite_Leave(object sender, EventArgs e)
        {
            if (txtSite.TextLength <= 0) return;

            var lists = _sharepointService.GetAllListsForSite(txtSite.Text);
            foreach (var list in lists)
            {
                cbList.Items.Add(list);
            }
            if (lists.Count > 0)
            {
                cbList.SelectedIndex = 0;
                cbList.Enabled = true;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            var list = _sharepointService.ExportListToObjects(txtSite.Text, cbList.Text);
            _excelService.WriteListToExcel(saveFileDialog1.FileName, list);

            MessageBox.Show(this, "Export complete", this.Text, MessageBoxButtons.OK);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;

            if (chkCleanup.Checked)
            {
                _sharepointService.DeleteAllListItems(txtSite.Text, cbList.Text);
            }

            var list = _excelService.ReadListFromExcel(openFileDialog1.FileName);
            _sharepointService.ImportObjectsToList(txtSite.Text, cbList.Text, list);

            MessageBox.Show(this, "Import complete", this.Text, MessageBoxButtons.OK);

            chkCleanup.Checked = false;
        }

        private void ChangeGuiState(object sender, EventArgs eventArgs)
        {
            if (string.IsNullOrEmpty(txtSite.Text) || string.IsNullOrEmpty(cbList.Text))
            {
                btnExport.Enabled = false;
                btnImport.Enabled = false;
                chkCleanup.Enabled = false;
            }
            else
            {
                btnExport.Enabled = true;
                btnImport.Enabled = true;
                chkCleanup.Enabled = true;
            }
        }
    }
}
