using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Data;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Windows.Forms;

namespace QuanLiSinhViennn
{
    public class BAL
    {
        private string MyConnection = "SERVER=localhost;" +
            "DATABASE=qlisinhvien;" +
            "UID=sa;" +
            "PASSWORD=Dungboro1;";



        public bool ThemSinhVien(string mssv, string name, string tuoiText, bool gioiTinh, string path)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(MyConnection))
                {
                    connection.Open();

                    string query = "INSERT INTO sinhvien (mssv, name, ngaysinh, gioitinh, hinhanh) VALUES (@mssv, @name, @ngaysinh, @gioitinh, @hinhanh)";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@name", name);
                        command.Parameters.AddWithValue("@ngaysinh", Convert.ToDateTime(tuoiText));
                        command.Parameters.AddWithValue("@gioitinh", gioiTinh);
                        command.Parameters.AddWithValue("@mssv", mssv);
                        command.Parameters.AddWithValue("@hinhanh", path);

                        int result = command.ExecuteNonQuery();
                        return result > 0;
                    }
                }
            }
            catch (Exception)
            {
                // Xử lý exception hoặc logging (tùy vào yêu cầu)
                return false;
            }
        }

        public bool IsMssvExists(string mssv)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(MyConnection))
                {
                    connection.Open();

                    string query = "SELECT COUNT(*) FROM sinhvien WHERE mssv = @mssv";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@mssv", mssv);

                        int count = (int)command.ExecuteScalar();

                        return count > 0;
                    }
                }
            }
            catch (Exception)
            {
                // Xử lý exception hoặc logging (tùy vào yêu cầu)
                return false;
            }
        }

        public bool XoaSinhVien(string mssvToDelete)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(MyConnection))
                {
                    connection.Open();

                    string query = "DELETE FROM sinhvien WHERE mssv = @mssv";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@mssv", mssvToDelete);
                        int affectedRows = command.ExecuteNonQuery();

                        return affectedRows > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi xóa sinh viên: {ex.Message}");
            }
        }

        public bool SuaSinhVien(string mssv, string name, string ngaysinh, bool gioitinh, string hinhanhPath, string mssv2)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(MyConnection))
                {
                    connection.Open();

                    string query = "UPDATE sinhvien SET mssv = @mssv, name = @name, ngaysinh = @ngaysinh, gioitinh = @gioitinh, hinhanh = @hinhanh WHERE mssv = @mssv2";

                    string hinhanhBytes = hinhanhPath; 

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@name", name);
                        command.Parameters.AddWithValue("@ngaysinh", Convert.ToDateTime(ngaysinh));
                        command.Parameters.AddWithValue("@gioitinh", gioitinh);
                        command.Parameters.AddWithValue("@mssv", mssv);
                        command.Parameters.AddWithValue("@hinhanh", hinhanhBytes);
                        command.Parameters.AddWithValue("@mssv2", mssv2); // MSSV cũ

                        int result = command.ExecuteNonQuery();

                        return result > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi sửa sinh viên: {ex.Message}");
            }
        }

        public DataTable TimKiemSinhVien(string keyword)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(MyConnection))
                {
                    connection.Open();

                    string query = "SELECT mssv, name, ngaysinh, gioitinh, hinhanh FROM sinhvien WHERE mssv LIKE @keyword OR name LIKE @keyword";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@keyword", "%" + keyword + "%");

                        SqlDataAdapter adapter = new SqlDataAdapter(command);
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        // Thêm cột mới để lưu dữ liệu hình ảnh dưới dạng byte[]
                        dt.Columns.Add("Pic", Type.GetType("System.Byte[]"));
                        foreach (DataRow row in dt.Rows)
                        {
                            // Đọc dữ liệu hình ảnh từ đường dẫn và gán vào cột "Pic"
                            string imagePath = row["hinhanh"].ToString();

                            // Add a check to ensure the file path is not empty or null
                            if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                            {
                                row["Pic"] = File.ReadAllBytes(imagePath);
                            }
                            else
                            {
                                row["Pic"] = DBNull.Value; // or handle it according to your needs
                            }
                        }

                        return dt;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi tìm kiếm sinh viên: {ex.Message}");
            }
        }


        public void ExportToPdf(DataGridView dataGridView, string filePath)
        {
            try
            {
                Document document = new Document();
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
                document.Open();

                PdfPTable pdfTable = new PdfPTable(dataGridView.Columns.Count);
                pdfTable.DefaultCell.Padding = 3;
                pdfTable.WidthPercentage = 100;
                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

                // Tiêu đề của các cột
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                    cell.BackgroundColor = new BaseColor(240, 240, 240);
                    pdfTable.AddCell(cell);
                }

                // Dữ liệu từ DataGridView
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        DataGridViewCell cell = row.Cells[i];

                        if (i != 4) // Kiểm tra nếu đây không phải là cột thêm vào từ hàm lammoi()
                        {
                            pdfTable.AddCell(cell.Value?.ToString() ?? string.Empty);
                        }
                        else
                        {
                            if (cell.Value != null)
                            {
                                string imagePath = cell.Value.ToString();
                                if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                                {
                                    iTextSharp.text.Image pdfImage = iTextSharp.text.Image.GetInstance(imagePath);
                                    pdfTable.AddCell(pdfImage);
                                }
                            }
                        }
                    }
                }

                // Thêm bảng vào tệp PDF
                document.Add(pdfTable);

                // Thêm hình ảnh từ PictureBox (nếu có)
                document.Close();
                writer.Close();

                MessageBox.Show("Xuất tệp PDF thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void ExportToExcel(DataGridView dataGridView, string filePath)
        {
            try
            {
                XSSFWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Sheet1");

                IRow headerRow = sheet.CreateRow(0);

                // Thêm cột header
                for (int j = 0; j < dataGridView.Columns.Count; j++)
                {
                    ICell cell = headerRow.CreateCell(j);
                    cell.SetCellValue(dataGridView.Columns[j].HeaderText);
                }

                // Duyệt qua từng dòng trong DataGridView và thêm dữ liệu vào sheet
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    IRow row = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        ICell cell = row.CreateCell(j);

                        object cellValue = dataGridView.Rows[i].Cells[j].Value;
                        if (cellValue != null)
                        {
                            if (cellValue is byte[] imageData) // Kiểm tra xem dữ liệu có phải là hình ảnh không
                            {
                                var pictureIdx = workbook.AddPicture(imageData, PictureType.JPEG);
                                var drawing = sheet.CreateDrawingPatriarch();
                                var anchor = new XSSFClientAnchor(0, 0, 0, 0, j, i + 1, j + 1, i + 2);
                                var picture = drawing.CreatePicture(anchor, pictureIdx);
                            }
                            else
                            {
                                cell.SetCellValue(cellValue.ToString());
                            }
                        }
                        else
                        {
                            cell.SetCellValue(string.Empty);
                        }
                    }
                }

                // Điều chỉnh kích thước cột cho phù hợp với nội dung
                for (int columnIndex = 0; columnIndex < dataGridView.Columns.Count; columnIndex++)
                {
                    sheet.AutoSizeColumn(columnIndex);
                }

                // Lưu tệp Excel tại vị trí đã chọn
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fileStream);
                }

                // Đóng workbook sau khi hoàn thành
                workbook.Close();

                // Thông báo thành công
                MessageBox.Show("Xuất dữ liệu thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Thông báo lỗi nếu có lỗi xảy ra
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}
