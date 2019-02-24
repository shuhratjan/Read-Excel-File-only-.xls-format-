using Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_To_Read_From_file
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataSet dataSetResult;
        private void btnOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xls", ValidateNames = true })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileStream fileStream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(fileStream);
                    reader.IsFirstRowAsColumnNames = true;
                    dataSetResult = reader.AsDataSet();
                    cboSheet.Items.Clear();
                    foreach(DataTable dt in dataSetResult.Tables)
                    {
                        cboSheet.Items.Add(dt.TableName);
                    }
                    reader.Close();
                }
            }
        }

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView.DataSource = dataSetResult.Tables[cboSheet.SelectedIndex];
        }
    }
}
