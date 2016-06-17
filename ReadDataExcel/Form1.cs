using Infragistics.Excel;
using System;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ReadDataExcel
{
    public partial class Form1: Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void simpleButtonSelectFile_Click(object sender,EventArgs e)
        {
            try
            {
                OpenFileDialog of = new OpenFileDialog();
                of.ShowDialog();
                string fileName = of.FileName;
                if (string.IsNullOrEmpty(fileName))
                {
                    return;
                }
                textEditFileName.Text = fileName;
                string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=Excel 8.0";
                // Tạo đối tượng kết nối
                OleDbConnection oledbConn = new OleDbConnection(connString);
                // Mở kết nối
                oledbConn.Open();
                DataTable dtexcel = new DataTable();
                DataTable schemaTable = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,new object[] { null,null,null,"TABLE" });
                //Looping Total Sheet of Xl File
                /*foreach (DataRow schemaRow in schemaTable.Rows)
                {
                }*/
                //Looping a first Sheet of Xl File
                DataRow schemaRow = schemaTable.Rows[0];
                string sheet = schemaRow["TABLE_NAME"].ToString();
                if (!sheet.EndsWith("_"))
                {
                    string query = "SELECT * FROM [Sheet1$]";
                    OleDbDataAdapter daexcel = new OleDbDataAdapter(query,oledbConn);
                    dtexcel.Locale = CultureInfo.CurrentCulture;
                    daexcel.Fill(dtexcel);
                }
                gridControl1.DataSource = dtexcel;
                oledbConn.Close();
            }
            catch (Exception ex)
            {
                EasyDialog.ShowErrorDialog(ex.Message);
            }
        }

        private void simpleButtonExport_Click(object sender,EventArgs e)
        {
            try
            {
                DataTable dtexcel = (DataTable)gridControl1.DataSource;
                var result = (from a in dtexcel.AsEnumerable() select a[1]).Distinct().ToList();
                foreach (var item in result)
                {
                    if (File.Exists(Application.StartupPath + item + ".xls"))
                    {
                        try
                        {
                            File.Delete(Application.StartupPath + item + ".xls");
                        }
                        catch (Exception ex)
                        {
                            EasyDialog.ShowErrorDialog(ex.Message);
                        }
                    }
                    var iRowCount = 16;
                    var iStt = 1;
                    var wb = Workbook.Load(Application.StartupPath + "\\MAU01.xls");
                    var rs = (from a in dtexcel.AsEnumerable() where a[1].Equals(item.ToString()) select a[7]).FirstOrDefault();
                    var rs1 = (from a in dtexcel.AsEnumerable() where a[1].Equals(item.ToString()) select a[8]).FirstOrDefault();
                    wb.Worksheets[0].Rows[4].Cells[0].Value = rs1.ToString();
                    wb.Worksheets[0].Rows[6].Cells[0].Value = "Lớp : " + rs.ToString();
                    wb.Worksheets[0].Rows[6].Cells[8].Value = "Năm nhập học : " + textEditNamNhapHoc.EditValue.ToString();
                    wb.Worksheets[0].Rows[7].Cells[8].Value = "Ngày thi : " + textEditNgayThi.EditValue.ToString();
                    wb.Worksheets[0].Rows[8].Cells[8].Value = "Lần thi thứ : " + textEditLanThi.EditValue.ToString();
                    wb.Worksheets[0].Rows[8].Cells[0].Value = "Ngành : " + textEditNganh.EditValue.ToString();
                    wb.Worksheets[0].Rows[9].Cells[0].Value = "Địa điểm thi : " + textEditDiaDiemThi.EditValue.ToString();
                    var dr = (from a in dtexcel.AsEnumerable() where a[1].Equals(item.ToString()) select a).CopyToDataTable();
                    for (var i = 0; i < dr.Rows.Count; i++)
                    {
                        wb.Worksheets[0].Rows[iRowCount].Cells[0].Value = iStt;
                        wb.Worksheets[0].Rows[iRowCount].Cells[1].Value = dr.Rows[i]["MASV"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[2].Value = dr.Rows[i]["HODEM"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[3].Value = dr.Rows[i]["TEN"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[4].Value = dr.Rows[i]["NGAYSINH"];
                        wb.Worksheets[0].Rows[iRowCount].Cells[5].Value = dr.Rows[i]["GIOITINH"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[6].Value = dr.Rows[i]["DIEMDANHGIA"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[7].Value = dr.Rows[i]["BTVN1"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[8].Value = dr.Rows[i]["BTVN2"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[9].Value = dr.Rows[i]["GIUAKY"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[12].Value = dr.Rows[i]["DIEUKIENTHI"].ToString();
                        iRowCount++;
                        wb.Worksheets[0].Rows[7].Cells[0].Value = "Tổng số học viên : " + iStt;
                        iStt++;
                    }
                    wb.Save(item.ToString() + ".xls");
                }
            }
            catch (Exception ex)
            {
                EasyDialog.ShowErrorDialog(ex.Message);
            }
        }
    }
}