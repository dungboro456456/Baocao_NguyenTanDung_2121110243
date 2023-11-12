using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace QuanLiSinhViennn
{
    public partial class Form1 : Form
    {
        private BAL bal;
        private BEL bel;
        private DAL dal;
        private string MyConnection =   "SERVER=localhost;" +
                                        "DATABASE=qlisinhvien;" +
                                        "UID=sa;" +
                                        "PASSWORD=Dungboro1;";
        public Form1()
        {
            InitializeComponent();
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            bal = new BAL();
            bel = new BEL();
            dal = new DAL();
        }
        private string path;
        private int index;

        private void Form1_Load(object sender, EventArgs e)
        {
            HienThiDanhSachSinhVien();
        }

        private void btThem_Click(object sender, EventArgs e)
        {
            try
            {
                string mssv = textboxMasv.Text;
                string name = textBoxName.Text;
                string tuoiText = textBoxTuoi.Text;
                bool gioiTinh = radiobtnam.Checked;

                if (string.IsNullOrWhiteSpace(mssv) || string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(tuoiText))
                {
                    MessageBox.Show("Vui lòng điền đầy đủ thông tin.");
                    return;
                }

                if (bal.IsMssvExists(mssv))
                {
                    MessageBox.Show("Mã số sinh viên đã tồn tại. Vui lòng chọn một mã số khác.");
                    return;
                }

                if (!radiobtnam.Checked && !radiobtnu.Checked)
                {
                    MessageBox.Show("Bạn chưa chọn giới tính.");
                    return;
                }

                bel.MSSV = mssv;
                bel.Name = name;
                bel.NgaySinh = tuoiText;
                bel.GioiTinh = gioiTinh;
                bel.HinhAnhPath = path;


                if (bal.ThemSinhVien(bel.MSSV, bel.Name, bel.NgaySinh, bel.GioiTinh, bel.HinhAnhPath))
                {
                    MessageBox.Show("Thêm sinh viên " + mssv + " thành công.");
                    HienThiDanhSachSinhVien();
                }
                else
                {
                    MessageBox.Show("Thêm sinh viên thất bại.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void HienThiDanhSachSinhVien()
        {
            DataTable dt = dal.LayDanhSachSinhVien();
            if (dt != null)
            {
                dataGridView1.DataSource = dt;
                dataGridView1.ClearSelection();
                if (index == dataGridView1.Rows.Count - 1)
                {
                    index = dataGridView1.Rows.Count - 2;
                    indamdong();
                }
                else
                {
                    indamdong();
                }
            }
            else
            {
                MessageBox.Show("Không thể lấy dữ liệu từ cơ sở dữ liệu.");
            }
        }
        public void clickchuot(DataGridViewCellEventArgs e)
        {
            try
            {
                index = e.RowIndex;
                DataGridViewRow row = dataGridView1.Rows[index];
                //textboxMasv.Enabled = false;
                textboxMasv.Text = row.Cells[0].Value.ToString().Trim();
                textBoxName.Text = row.Cells[1].Value.ToString().Trim();
                textBoxTuoi.Text = row.Cells[2].Value.ToString().Trim();

                // Kiểm tra giới tính
                bool checkgioitinh = (bool)row.Cells[3].Value;
                if (checkgioitinh)
                {
                    radiobtnam.Checked = true;
                }
                else
                {
                    radiobtnu.Checked = true;
                }

                // Hiển thị hình ảnh
                byte[] imgData = (byte[])row.Cells["Pic"].Value;
                if (imgData != null && imgData.Length > 0)
                {
                    MemoryStream ms = new MemoryStream(imgData);
                    pictureBox1.Image = System.Drawing.Image.FromStream(ms);
                }
                else
                {
                    pictureBox1.Image = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\ndưới đó chưa có dữ liệu đâu");
            }

        }
        public void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            index = e.RowIndex;
            if (index >= 0)
            {
                // Bạn muốn thay đổi màu của hàng, hãy đặt màu cho tất cả ô trong hàng
                indamdong();
            }
            clickchuot(e);
        }
        public void indamdong()
        {
            DataGridViewRow selectedRow = dataGridView1.Rows[index];
            selectedRow.Selected = true;
            selectedRow.DefaultCellStyle.SelectionBackColor = Color.LightBlue; // Màu nền khi chọn
            selectedRow.DefaultCellStyle.SelectionForeColor = Color.Black; // Màu chữ khi chọn
        }

        // Trong class Form1 hoặc class tương ứng của bạn


        private void btSua_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(textboxMasv.Text) || string.IsNullOrWhiteSpace(textBoxName.Text) || string.IsNullOrWhiteSpace(textBoxTuoi.Text) || pictureBox1.Image == null)
                {
                    MessageBox.Show("Vui lòng điền đầy đủ thông tin.");
                    return;
                }
                if (index < 0)
                {
                    MessageBox.Show("Vui lòng chọn một sinh viên để sửa.");
                    return;
                }
                if (!radiobtnam.Checked && !radiobtnu.Checked)
                {
                    MessageBox.Show("Bạn chưa chọn giới tính.");
                    return;
                }

                string mssv = textboxMasv.Text;
                string name = textBoxName.Text;
                string tuoiText = textBoxTuoi.Text;
                bool gioiTinh = radiobtnam.Checked;

                // Kiểm tra xem Mssv đã thay đổi so với giá trị ban đầu
                string mssv2 = dataGridView1.Rows[index].Cells[0].Value.ToString().Trim();
                if (string.IsNullOrEmpty(path))
                {
                    path = dataGridView1.Rows[index].Cells[4].Value?.ToString()?.Trim();
                }


                if (!mssv.Equals(mssv2))
                {
                    MessageBox.Show("Không được thay đổi MSSV");
                    return;
                }

                bool success = bal.SuaSinhVien(mssv, name, tuoiText, gioiTinh, path, mssv2);

                if (success)
                {
                    MessageBox.Show($"Sửa thông tin sinh viên {mssv} thành công.");
                    HienThiDanhSachSinhVien();
                }
                else
                {
                    MessageBox.Show($"Sửa thông tin sinh viên {mssv} thất bại.");
                    HienThiDanhSachSinhVien();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }
        private void btThoat_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát khỏi ứng dụng?", "Xác nhận thoát", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Người dùng đã xác nhận thoát, đóng ứng dụng
                Close();
            }
            else
            {
                // Người dùng đã hủy việc thoát, không làm gì cả
            }
        }
        private void importExcelButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = openFileDialog.FileName;

                        using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                        {
                            XSSFWorkbook workbook = new XSSFWorkbook(file);
                            ISheet sheet = workbook.GetSheetAt(0);

                            DataTable dt = new DataTable();

                            // Thêm cột vào DataTable dựa trên số cột hiện tại trong DataGridView
                            foreach (DataGridViewColumn column in dataGridView1.Columns)
                            {
                                // Add only the columns you need, excluding the 'Pic' column
                                if (column.Name != "Pic")
                                {
                                    dt.Columns.Add(column.HeaderText);
                                }
                            }

                            for (int i = 1; i <= sheet.LastRowNum; i++)
                            {
                                IRow row = sheet.GetRow(i);
                                if (row != null)
                                {
                                    DataRow newRow = dt.NewRow();
                                    for (int j = 0; j < dt.Columns.Count; j++)
                                    {
                                        if (j < row.LastCellNum)
                                        {
                                            ICell cell = row.GetCell(j);
                                            if (cell != null)
                                            {
                                                try
                                                {
                                                    if (cell.CellType == CellType.Blank)
                                                    {
                                                        newRow[j] = DBNull.Value;
                                                    }
                                                    else if (cell.CellType == CellType.String)
                                                    {
                                                        newRow[j] = cell.StringCellValue;
                                                    }
                                                    else if (cell.CellType == CellType.Numeric)
                                                    {
                                                        newRow[j] = cell.NumericCellValue;
                                                    }
                                                    else if (cell.CellType == CellType.Boolean)
                                                    {
                                                        newRow[j] = cell.BooleanCellValue;
                                                    }
                                                    else if (cell is IPicture)
                                                    {
                                                        // Đọc dữ liệu hình ảnh và lưu vào DataTable
                                                        byte[] pictureData = (cell as XSSFPicture).PictureData.Data;
                                                        newRow[j] = pictureData;
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    // Bắt lỗi và thực hiện xử lý tùy chọn
                                                    Console.WriteLine($"Error reading cell ({i + 1}, {j + 1}): {ex.Message}");
                                                    newRow[j] = DBNull.Value;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // Cột Excel không đủ để fill hết cột trong DataGridView
                                            newRow[j] = DBNull.Value;
                                        }
                                    }
                                    dt.Rows.Add(newRow);
                                }
                            }

                            dt.Columns.Add("Pic", Type.GetType("System.Byte[]"));

                            foreach (DataRow row in dt.Rows)
                            {
                                // Đọc dữ liệu hình ảnh từ đường dẫn và gán vào cột "Pic"
                                string imagePath = row[4].ToString(); // Assuming "hinhanh" is the correct column name

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

                            dataGridView1.DataSource = dt;
                            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error reading Excel file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        private void btexportpdf_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "PDF Files|*.pdf";
                saveFileDialog.Title = "Chọn vị trí và tên tệp PDF";
                saveFileDialog.FileName = "SinhVien.pdf";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string pdfFilePath = saveFileDialog.FileName;
                    BAL bal = new BAL();
                    bal.ExportToPdf(dataGridView1, pdfFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void btexportexcel_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.Title = "Chọn vị trí và tên tệp Excel";
                saveFileDialog.FileName = "SinhVien.xlsx"; // Tên mặc định

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string excelFilePath = saveFileDialog.FileName;
                    BAL bal = new BAL();
                    bal.ExportToExcel(dataGridView1, excelFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog opnfd = new OpenFileDialog();
                opnfd.Filter = "Image Files (*.jpg;*.jpeg;*.gif;*.png)|*.jpg;*.jpeg;*.gif;*.png";

                if (opnfd.ShowDialog() == DialogResult.OK)
                {
                    // Kiểm tra kích thước tối đa cho hình ảnh (ví dụ: giới hạn kích thước 300x300)
                    int maxWidth = 500;
                    int maxHeight = 500;

                    System.Drawing.Image selectedImage = System.Drawing.Image.FromFile(opnfd.FileName);

                    // Kiểm tra kích thước của hình ảnh
                    if (selectedImage.Width <= maxWidth && selectedImage.Height <= maxHeight)
                    {
                        pictureBox1.Image = new Bitmap(selectedImage);
                        path = opnfd.FileName;
                    }
                    else
                    {
                        MessageBox.Show("Hình ảnh vượt quá kích thước tối đa (" + maxWidth + "x" + maxHeight + "). Vui lòng chọn hình ảnh khác.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "Có vẻ như thiếu hình rồi");
            }
        }

        private void button1_TimKiem_Click(object sender, EventArgs e)
        {
            try
            {
                string keyword = tbTimkiem.Text;

                DataTable dt = bal.TimKiemSinhVien(keyword);

                if (dt != null && dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    dataGridView1.ClearSelection();

                }
                else
                {
                    MessageBox.Show("Không tìm thấy sinh viên nào phù hợp với từ khóa.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }
        private void btLamMoi_Click(object sender, EventArgs e)
        {
            HienThiDanhSachSinhVien();
            textboxMasv.Text = String.Empty;
            textboxMasv.Enabled = true;
            textBoxName.Text = String.Empty;
            textBoxTuoi.Text = String.Empty;
            tbTimkiem.Text = String.Empty;
            pictureBox1.Image = null;
        }

        private void btXoa_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0 && index >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[index];
                    string mssvToDelete = row.Cells[0].Value.ToString().Trim();

                    DialogResult result = MessageBox.Show($"Bạn có chắc chắn muốn xóa sinh viên {mssvToDelete} không?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        bool success = bal.XoaSinhVien(mssvToDelete);

                        if (success)
                        {
                            MessageBox.Show($"Xóa thông tin sinh viên {mssvToDelete} thành công.");
                            HienThiDanhSachSinhVien();
                            if (dataGridView1.Rows.Count == 0)
                            {
                                index = -1;
                            }
                            // Có thể tăng giá trị của index nếu bạn muốn chuyển tới dòng tiếp theo
                        }
                        else
                        {
                            MessageBox.Show($"Xóa thông tin sinh viên {mssvToDelete} thất bại.");
                            HienThiDanhSachSinhVien();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Không có dữ liệu hoặc bạn chưa chọn một sinh viên để xóa.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


    // Các phương thức và sự kiện khác ở đây...
}
}
