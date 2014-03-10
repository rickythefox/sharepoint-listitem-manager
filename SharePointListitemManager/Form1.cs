using System;
using System.Linq;
using System.Windows.Forms;

namespace SharePointListitemManager
{
    public partial class Form1 : Form
    {
        private readonly ISharepointService _sharepointService;
        private readonly IExcelService _excelService;

        public Form1(ISharepointService sharepointService, IExcelService excelService)
        {
            _sharepointService = sharepointService;
            _excelService = excelService;

            InitializeComponent();

            txtSite.TextChanged += ChangeButtonState;
            cbList.SelectedIndexChanged += ChangeButtonState;
            Shown += ChangeButtonState;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
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
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            var list = _excelService.ReadListFromExcel("");
            _sharepointService.ImportObjectsToList(txtSite.Text, cbList.Text, list);
        }

        private void ChangeButtonState(object sender, EventArgs eventArgs)
        {
            if (string.IsNullOrEmpty(txtSite.Text))
            {
                cbList.Enabled = false;
            }

            if (string.IsNullOrEmpty(txtSite.Text) || string.IsNullOrEmpty(cbList.Text))
            {
                btnExport.Enabled = false;
                btnImport.Enabled = false;
            }
            else
            {
                btnExport.Enabled = true;
                btnImport.Enabled = true;
            }
        }
    }
}
