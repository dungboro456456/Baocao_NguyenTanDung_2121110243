using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;

namespace QuanLiSinhViennn
{
    public class DAL
    {
        private string MyConnection = "SERVER=localhost;" +
            "DATABASE=qlisinhvien;" +
            "UID=sa;" +
            "PASSWORD=Dungboro1;";

        public DataTable LayDanhSachSinhVien()
        {
            using (SqlConnection connection = new SqlConnection(MyConnection))
            {
                try
                {
                    connection.Open();

                    string query = "SELECT mssv, name, ngaysinh, gioitinh, hinhanh FROM sinhvien";

                    SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    // Thêm cột mới để lưu dữ liệu hình ảnh dưới dạng byte[]
                    dt.Columns.Add("Pic", Type.GetType("System.Byte[]"));
                    foreach (DataRow row in dt.Rows)
                    {
                        // Đọc dữ liệu hình ảnh từ đường dẫn và gán vào cột "Pic"
                        string imagePath = row["hinhanh"].ToString();

                        // Add a check to ensure the file path is not empty or null
                        if (!string.IsNullOrEmpty(imagePath) && System.IO.File.Exists(imagePath))
                        {
                            row["Pic"] = System.IO.File.ReadAllBytes(imagePath);
                        }
                        else
                        {
                            row["Pic"] = DBNull.Value; // or handle it according to your needs
                        }
                    }

                    return dt;
                }
                catch (Exception)
                {
                    // Xử lý exception hoặc logging (tùy vào yêu cầu)
                    return null;
                }
            }
        }


    }
}
