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
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textEditSavaFilePath.Text = folderBrowserDialog1.SelectedPath;
            }

            try
            {
                DataTable dtexcel = (DataTable)gridControl1.DataSource;
                var result = (from a in dtexcel.AsEnumerable() select a[1]).Distinct().ToList();
                foreach (var item in result)
                {
                    if (File.Exists(textEditSavaFilePath.EditValue.ToString() + "\\" + item.ToString() + ".xls"))
                    {
                        try
                        {
                            File.Delete(textEditSavaFilePath.EditValue.ToString() + "\\" + item.ToString() + ".xls");
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
                        // Format collum STT
                        wb.Worksheets[0].Rows[iRowCount].Cells[0].Value = iStt;
                        wb.Worksheets[0].Rows[iRowCount].Cells[0].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[0].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[0].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[0].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[0].CellFormat.ShrinkToFit = ExcelDefaultableBoolean.True;
                        // Format collum MASV
                        wb.Worksheets[0].Rows[iRowCount].Cells[1].Value = dr.Rows[i]["MASV"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[1].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[1].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[1].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[1].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[1].CellFormat.ShrinkToFit = ExcelDefaultableBoolean.True;
                        // Format collum HODEM
                        wb.Worksheets[0].Rows[iRowCount].Cells[2].Value = dr.Rows[i]["HODEM"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[2].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[2].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[2].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[2].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[2].CellFormat.ShrinkToFit = ExcelDefaultableBoolean.True;
                        // Format collum TEN
                        wb.Worksheets[0].Rows[iRowCount].Cells[3].Value = dr.Rows[i]["TEN"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[3].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[3].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[3].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[3].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[3].CellFormat.ShrinkToFit = ExcelDefaultableBoolean.True;
                        // Format collum NGAYSINH
                        wb.Worksheets[0].Rows[iRowCount].Cells[4].Value = dr.Rows[i]["NGAYSINH"];
                        wb.Worksheets[0].Rows[iRowCount].Cells[4].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[4].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[4].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[4].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[4].CellFormat.ShrinkToFit = ExcelDefaultableBoolean.True;
                        // Format collum GIOITINH
                        wb.Worksheets[0].Rows[iRowCount].Cells[5].Value = dr.Rows[i]["GIOITINH"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[5].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[5].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[5].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[5].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[5].CellFormat.ShrinkToFit = ExcelDefaultableBoolean.True;
                        // Format collum DIEMDANHGIA
                        wb.Worksheets[0].Rows[iRowCount].Cells[6].Value = dr.Rows[i]["DIEMDANHGIA"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[6].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[6].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[6].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[6].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;

                        // Format collum BTVN1
                        wb.Worksheets[0].Rows[iRowCount].Cells[7].Value = dr.Rows[i]["BTVN1"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[7].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[7].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[7].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[7].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;

                        // Format collum BTVN2
                        wb.Worksheets[0].Rows[iRowCount].Cells[8].Value = dr.Rows[i]["BTVN2"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[8].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[8].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[8].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[8].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;

                        // Format collum GIUAKY
                        wb.Worksheets[0].Rows[iRowCount].Cells[9].Value = dr.Rows[i]["GIUAKY"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[9].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[9].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[9].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[9].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;

                        // Format collum DIEMTHI
                        wb.Worksheets[0].Rows[iRowCount].Cells[10].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[10].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[10].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[10].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;

                        // Format collum DIEMTONGKET
                        wb.Worksheets[0].Rows[iRowCount].Cells[11].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[11].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[11].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[11].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;
                        // Format collum DIEUKIENTHI
                        wb.Worksheets[0].Rows[iRowCount].Cells[12].Value = dr.Rows[i]["DIEUKIENTHI"].ToString();
                        wb.Worksheets[0].Rows[iRowCount].Cells[12].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[12].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[12].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[12].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[12].CellFormat.ShrinkToFit = ExcelDefaultableBoolean.True;
                        // Format collum GHICHU
                        wb.Worksheets[0].Rows[iRowCount].Cells[13].CellFormat.BottomBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[13].CellFormat.TopBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[13].CellFormat.LeftBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[13].CellFormat.RightBorderStyle = CellBorderLineStyle.Thin;
                        wb.Worksheets[0].Rows[iRowCount].Cells[13].CellFormat.ShrinkToFit = ExcelDefaultableBoolean.True;
                        iRowCount++;
                        wb.Worksheets[0].Rows[7].Cells[0].Value = "Tổng số học viên : " + iStt;
                        iStt++;
                    }
                    // Insert new empty row
                    wb.Worksheets[0].Rows[iRowCount].Cells[0].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[1].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[2].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[3].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[4].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[5].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[6].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[7].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[8].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[9].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[10].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[11].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[12].Value = string.Empty;
                    wb.Worksheets[0].Rows[iRowCount].Cells[13].Value = string.Empty;
                    //Merged cell for footer
                    WorksheetMergedCellsRegion mergedRergion = wb.Worksheets[0].MergedCellsRegions.Add(iRowCount + 1,0,iRowCount + 1,13);
                    WorksheetMergedCellsRegion mergedRergion1 = wb.Worksheets[0].MergedCellsRegions.Add(iRowCount + 2,0,iRowCount + 2,5);
                    WorksheetMergedCellsRegion mergedRergion2 = wb.Worksheets[0].MergedCellsRegions.Add(iRowCount + 2,6,iRowCount + 2,13);
                    WorksheetMergedCellsRegion mergedRergion3 = wb.Worksheets[0].MergedCellsRegions.Add(iRowCount + 3,6,iRowCount + 3,13);
                    WorksheetMergedCellsRegion mergedRergion4 = wb.Worksheets[0].MergedCellsRegions.Add(iRowCount + 3,0,iRowCount + 3,5);
                    // Footer 1
                    mergedRergion.Value = "Công thức tính điểm : Tùy theo từng môn học : D = A*0.1 + B*0.3 (Hoặc 0.2 tùy từng môn) + T*0.6";
                    mergedRergion.CellFormat.VerticalAlignment = VerticalCellAlignment.Center;
                    mergedRergion.CellFormat.Alignment = HorizontalCellAlignment.Left;
                    mergedRergion.CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
                    // Footer 2
                    mergedRergion1.Value = "Cán bộ vào điểm";
                    mergedRergion1.CellFormat.VerticalAlignment = VerticalCellAlignment.Center;
                    mergedRergion1.CellFormat.Alignment = HorizontalCellAlignment.Center;
                    mergedRergion1.CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
                    // Footer 3
                    mergedRergion2.Value = "Thái Nguyên, Ngày ... tháng .. năm 20..";
                    mergedRergion2.CellFormat.VerticalAlignment = VerticalCellAlignment.Center;
                    mergedRergion2.CellFormat.Alignment = HorizontalCellAlignment.Center;
                    mergedRergion2.CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
                    // Footer 4
                    mergedRergion3.Value = "TRUNG TÂM ĐTTX";
                    mergedRergion3.CellFormat.VerticalAlignment = VerticalCellAlignment.Center;
                    mergedRergion3.CellFormat.Alignment = HorizontalCellAlignment.Center;
                    mergedRergion3.CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
                    //Save file
                    if (string.IsNullOrEmpty(textEditSavaFilePath.EditValue.ToString()))
                    {
                        wb.Save(item.ToString() + ".xls");
                    }
                    else
                    {
                        wb.Save(textEditSavaFilePath.EditValue.ToString() + "\\" + item.ToString() + ".xls");
                    }
                }
            }
            catch (Exception ex)
            {
                EasyDialog.ShowErrorDialog(ex.Message);
            }
        }

        private void Form1_Load(object sender,EventArgs e)
        {
            textEditSavaFilePath.Enabled = false;
        }
    }
}