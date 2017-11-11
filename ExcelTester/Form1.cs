using ExcelEditor;
using System;
using System.Windows.Forms;

namespace ExcelTester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            rb_logReports.Checked = true;
        }

        private void btn_Browse_Click(object sender, EventArgs e)
        {
            var fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel file (*.xls)|*.xlsx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                tb_inputFile.Text = fileDialog.FileName;
            }
            //var folderDialog = new FolderBrowserDialog();
            //if (folderDialog.ShowDialog() == DialogResult.OK)
            //{
            //    txtBox_inputFolder.Text = folderDialog.SelectedPath;
            //}
        }

        private void btn_doWork_Click(object sender, EventArgs e)
        {

            if (cb_removeFormulas.Checked == true)
            {
                var result2 = ExcelEditorServices.ReplaceFormulasWithValues(tb_inputFile.Text);
            }

            MessageBox.Show("Job Done");
        }
    }
}
