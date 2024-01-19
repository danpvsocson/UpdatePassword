using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using OfficeOpenXml;
using ExcelDataReader;
using Z.Dapper.Plus;
using System.IO;
using System.Data.SqlTypes;

namespace UpdatePassword
{
    public partial class Form1 : Form
    {
        private DataTableCollection tables;
        private DataTable excelDataTable;
        private string connectionString = "Server=CAD001\\WEB;Database=THIETBI;User=sa;Password=abc123";
        public Form1()
        {
            InitializeComponent();
        }

        private void find_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtPath.Text = ofd.FileName;
                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                            tables = result.Tables;
                            comboBox1.Items.Clear();
                            foreach (DataTable table in tables)
                                comboBox1.Items.Add(table.TableName);
                        }
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                DataTable dt = tables[comboBox1.SelectedItem.ToString()];

                // Remove empty columns
                List<DataColumn> columnsToRemove = new List<DataColumn>();
                foreach (DataColumn column in dt.Columns)
                {
                    bool isColumnEmpty = true;
                    foreach (DataRow row in dt.Rows)
                    {
                        if (!string.IsNullOrWhiteSpace(row[column.ColumnName].ToString()))
                        {
                            isColumnEmpty = false;
                            break;
                        }
                    }
                    if (isColumnEmpty)
                    {
                        columnsToRemove.Add(column);
                    }
                }

                foreach (DataColumn columnToRemove in columnsToRemove)
                {
                    dt.Columns.Remove(columnToRemove);
                }

                dataGridView1.DataSource = dt;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                DataTable dt = tables[comboBox1.SelectedItem.ToString()];
                btnImport.Enabled = false;
                // Đảm bảo các cột bắt buộc tồn tại trong DataTable
                //if (!dt.Columns.Contains("tk") || !dt.Columns.Contains("mk"))
                //{
                //    MessageBox.Show("Cột 'Tên tài khoản' và 'Mật khẩu' là bắt buộc.");
                //    return;
                //}

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    progressBar.Minimum = 0;
                    progressBar.Maximum = dt.Rows.Count;
                    progressBar.Step = 1;
                    progressBar.Value = 0;

                    foreach (DataRow row in dt.Rows)
                    {
                        string selectQuery = "SELECT TOP 1 * FROM nhatky ORDER BY STT DESC";
                        using (SqlCommand command1 = new SqlCommand(selectQuery, connection))
                        {
                            object result = command1.ExecuteScalar();
                            if (result != DBNull.Value && int.TryParse(result.ToString(), out int maxStt))
                            {
                                int stt = maxStt + 1;

                                //string stt = row["STT"].ToString();
                                string tb = row["TB"].ToString();
                                string mk = row["New Password"].ToString();
                                string tk = row["Account 4"].ToString();
                                string ngay = DateTime.Today.ToString("yyyy-MM-dd");
                                string error = "Thay đổi mật khẩu định kỳ";
                                string action = "Thay đổi mật khẩu định kỳ";
                                TimeSpan start_t = new TimeSpan(8, 0, 0);
                                TimeSpan end_t = new TimeSpan(8, 30, 0);
                                string manv = "4247";
                                string xacnhan = row["Name 2"].ToString();
                                string checkid = SqlBoolean.True.ToString();
                                string nvnhap = manv;
                                string ngaynhap = ngay;
                                //string tb = Environment.MachineName; //tên tb

                                // Thêm dữ liệu vào bảng SQL
                                string insertQuery = "INSERT INTO Nhatky (STT, TenTB, Ngay, Error, Action, Start_t, End_t, MaNV, XacNhan, Check_ID, NVnhap, Ngaynhap) VALUES" +
                                    " (@stt, @tb, @ngay, @error, @action, @start_t, @end_t, @manv, @xacnhan, @checkid, @nvnhap, @ngaynhap)";

                                using (SqlCommand command = new SqlCommand(insertQuery, connection))
                                {
                                    command.Parameters.AddWithValue("@stt", stt);
                                    command.Parameters.AddWithValue("@tb", tb);
                                    command.Parameters.AddWithValue("@ngay", ngay);
                                    command.Parameters.AddWithValue("@error", error);
                                    command.Parameters.AddWithValue("@action", action);
                                    command.Parameters.AddWithValue("@start_t", start_t);
                                    command.Parameters.AddWithValue("@end_t", end_t);
                                    command.Parameters.AddWithValue("@manv", manv);
                                    command.Parameters.AddWithValue("@xacnhan", xacnhan);
                                    command.Parameters.AddWithValue("@checkid", checkid);
                                    command.Parameters.AddWithValue("@nvnhap", nvnhap);
                                    command.Parameters.AddWithValue("@ngaynhap", ngaynhap);
                                    //command.Parameters.AddWithValue("@tb", tb);
                                    command.ExecuteNonQuery();
                                }

                                // Update mật khẩu vào cơ sở dữ liệu
                                string updateQuery = "UPDATE Thietbi SET userpass = @password WHERE TenTB = @tb";

                                using (SqlCommand command = new SqlCommand(updateQuery, connection))
                                {
                                    command.Parameters.AddWithValue("@password", mk);
                                    command.Parameters.AddWithValue("@tb", tb);
                                    command.ExecuteNonQuery();
                                }
                                progressBar.PerformStep();
                                int percent = (progressBar.Value * 100) / progressBar.Maximum;
                                progressBar.CreateGraphics().DrawString(percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar.Width / 2 - 10, progressBar.Height / 2 - 7));
                            }
                        }
                    }
                    MessageBox.Show("Update dữ liệu thành công!", "Trạng Thái");

                }
                Application.Exit();
            }
        }
    }
}
