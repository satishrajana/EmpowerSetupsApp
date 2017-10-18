using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmpowerSetupsList
{
    public partial class Setup : Form
    {
        public Setup()
        {
            InitializeComponent();
            btnSearch.Enabled = false;
            btnExportToExcel.Enabled = false;
            txtSearch.Text = string.Empty;
            btnClearSrch.Enabled = false;
        }

        private void btnQuickCode_Click(object sender, EventArgs e)
        {
            if (txtFolder.Text == string.Empty)
            {
                lblStatus.Text = "Enter a valid plug folder path : /Empower.NET/Setups/Empower.Setups.Plugins";
                return;
            };

            SeedTheGrid();
           
        }

        public void SeedTheGrid()
        {
            DataTable QcTable = new DataTable();
            DataTable dt = new DataTable();
            txtSearch.Text = string.Empty;


            try
            {
                //Get List of Quickcode and Description
                DataTable Codes = GetQuickCodes(QcTable);
                //Get List of Description, Master and Child Table
                DataTable MCTables = GetMasterChildTables(dt);

                //Clone table A to add columns from A to target tabe
                DataTable targetTable = QcTable.Clone();
                //select columns from B that need to be added to target table
                var dt2Columns = dt.Columns.OfType<DataColumn>().Select(dc =>
                new DataColumn(dc.ColumnName, dc.DataType, dc.Expression, dc.ColumnMapping));
                //remove from target table, column that is common between A and B
                var dt2FinalColumns = from dc in dt2Columns.AsEnumerable()
                                      where targetTable.Columns.Contains(dc.ColumnName) == false
                                      select dc;
                //Add columns from B to target table					  
                targetTable.Columns.AddRange(dt2FinalColumns.ToArray());
                //retrieve rows from A and B - minus the Description column that is common
                var rowData = from row1 in QcTable.AsEnumerable()
                              join row2 in dt.AsEnumerable()
                              on row1.Field<string>("Description") equals row2.Field<string>("Description")
                              orderby row1.Field<string>("Description")
                              select row1.ItemArray.Concat(row2.ItemArray.Where(r2 => row1.ItemArray.Contains(r2) == false)).ToArray();
                //add rows to target table			  
                foreach (object[] values in rowData)
                    targetTable.Rows.Add(values);
                targetTable.DefaultView.Sort = "QuickCode ASC";

                dataGridView1.DataSource = targetTable;
                dataGridView1.AutoResizeColumns();
                dataGridView1.ScrollBars = ScrollBars.Vertical;

                foreach (DataGridViewBand band in dataGridView1.Columns)
                {
                    band.ReadOnly = true;
                }
                lblStatus.Text = "Total Number of Rows :" + (dataGridView1.RowCount - 1).ToString();

                //textBox1.Text = @"C:\Empower\BASE_7_0_1_SP1\Empower.NET\Setups\Empower.Setups.Plugins";
                btnExportToExcel.Enabled = true;
                btnSearch.Enabled = true;
                

            }
            catch (DirectoryNotFoundException DirNotFound)
            {
                lblStatus.Text = DirNotFound.Message;
            }
            catch (UnauthorizedAccessException UnAuthDir)
            {
                lblStatus.Text = UnAuthDir.Message;
            }
            catch (PathTooLongException LongPath)
            {
                lblStatus.Text = LongPath.Message;
            }
            catch (System.ArgumentException ane)
            {
                lblStatus.Text = "Select a valid folder";
                lblStatus.ForeColor = System.Drawing.Color.Red;
            }
        }
        public DataTable GetQuickCodes(DataTable QcTable)
        {

            QcTable.Clear();
            QcTable.Columns.Add("QuickCode", typeof(string));
            QcTable.Columns.Add("Description", typeof(string));

            string foldername = txtFolder.Text;
            var files = from file in Directory.EnumerateFiles(foldername, "*.cs", SearchOption.AllDirectories)
                        from line in File.ReadLines(file)
                        where line.Contains("Plugin(")
                        select new
                        { File = file, Line = line };


            foreach (var f in files)
            {
                String l = f.Line;
                string fl = Path.GetFileName(f.File).Split('.')[0];
                int startidx = l.IndexOf("Plugin(\"", 0, StringComparison.CurrentCulture);

                if (startidx > 0)
                {
                    int endidx = l.Length;
                    string result = l.Substring(startidx, endidx - startidx);
                    Regex regex = new Regex(".*?\\(.*?,.*?, .*?, \"(?<WhatEverPP01Is>.*?)\",.*");

                    var match = regex.Match(result);

                    if (match.Success)
                    {
                        var groupValue = fl;
                        var whateverPP01Is = match.Groups["WhatEverPP01Is"].Value;
                        String str = whateverPP01Is + " - " + groupValue;
                        QcTable.Rows.Add(whateverPP01Is, groupValue);
                    }
                }
            }
            return QcTable;
        }

        public DataTable GetMasterChildTables(DataTable dt)

        {
            string foldername = txtFolder.Text;
            var files = from file in Directory.EnumerateFiles(foldername, "*.designer.cs", SearchOption.AllDirectories)
                        from line in File.ReadLines(file)
                        where line.Contains("TableName")
                        orderby Path.GetFileName(file).Split('.')[0]
                        select new
                        { File = file, Line = line };

            string cn, pcn, ftn, tn, ptn, pctn;
            cn = pcn = ftn = tn = ptn = pctn = string.Empty;

            dt.Clear();
            dt.Columns.Add("Description");
            dt.Columns.Add("MasterTable");
            dt.Columns.Add("ChildTable");

            foreach (var f in files)
            {
                //Displays the line that has the TABLENAME
                String l = f.Line;
                //Displays the file name
                string fl = f.File;
                cn = Path.GetFileName(fl).Split('.')[0];

                Regex reg = new Regex("\"(?<Group>.*?)\"");
                var match = reg.Match(l);
                tn = match.Groups["Group"].Value;

                if (pcn == "" && ptn == "" && pctn == "")
                {
                    pcn = cn; ptn = tn;
                }
                if (cn == pcn && ptn != "" && pcn != "")
                {
                    pctn = tn;
                }
                if (cn != pcn && pcn != "")
                {
                    if (ptn != "" || pctn != "")
                    {
                        ftn = ptn + "			" + pctn;
                        Console.WriteLine(pcn + "				" + ftn);
                        dt.Rows.Add(pcn, ptn, pctn);
                    }
                    pcn = cn; ptn = tn; pctn = "";
                }
            }
            return dt;
        }
        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = "";
            FolderBrowserDialog fb = new FolderBrowserDialog();
            fb.ShowNewFolderButton = false;
            fb.SelectedPath = @"C:";
            
            if (fb.ShowDialog() == DialogResult.OK)
            {
                txtFolder.Text = fb.SelectedPath;
            }
            lblStatus.Text = string.Empty;
        }

        private void Setup_Load(object sender, EventArgs e)
        { }
        private void btnExit_Click(object sender, EventArgs e)
        {
            Setup.ActiveForm.Close();
        }
        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            Excel._Application app = new Excel.Application();
            Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Excel._Worksheet worksheet = null;
            app.Visible = true;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;

            // changing the name of active sheet
            worksheet.Name = "Setups-QuickCode";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }

            workbook.SaveAs("c:\\output.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            
            if (txtSearch.Text == string.Empty)
            {
                lblStatus.Text = "Please enter a valid Description to Search";
            }
            else
            {
                int sortN = cmbSelect.SelectedIndex < 0 ? 1 : cmbSelect.SelectedIndex;
                
                String srchVale = Convert.ToString(txtSearch.Text);
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns[sortN].HeaderText.ToString() + " Like '%" + srchVale + "%'";
                int rowcount = dataGridView1.RowCount;
                lblStatus.Text = "Rows Returned :  " + (rowcount - 1).ToString();
                dataGridView1.DataSource = bs;
                btnClearSrch.Enabled = true;
            }


        }

        private void btnClearSrch_Click(object sender, EventArgs e)
        {
            btnSearch.Enabled = false;
            btnExportToExcel.Enabled = false;
            txtSearch.Text = string.Empty;
            dataGridView1.DataSource = null;
            SeedTheGrid();
        }
    }

}

