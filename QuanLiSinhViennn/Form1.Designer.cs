namespace QuanLiSinhViennn
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            radiobtnam = new RadioButton();
            dataGridView1 = new DataGridView();
            btThoat = new Button();
            btSua = new Button();
            btXoa = new Button();
            btThem = new Button();
            label7 = new Label();
            label6 = new Label();
            label3 = new Label();
            textBoxName = new TextBox();
            textBox2 = new TextBox();
            label5 = new Label();
            label2 = new Label();
            textboxMasv = new TextBox();
            label4 = new Label();
            textBox1 = new TextBox();
            label1 = new Label();
            btLammoi = new Button();
            tbTimkiem = new TextBox();
            lbTimkiem = new Label();
            btTimkiem = new Button();
            textBoxTuoi = new TextBox();
            pictureBox1 = new PictureBox();
            radiobtnu = new RadioButton();
            button1 = new Button();
            btexportexcel = new Button();
            btimportexcel = new Button();
            btexportpdf = new Button();
            Column1 = new DataGridViewTextBoxColumn();
            Column2 = new DataGridViewTextBoxColumn();
            Column3 = new DataGridViewTextBoxColumn();
            Column4 = new DataGridViewCheckBoxColumn();
            Column5 = new DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            SuspendLayout();
            // 
            //radiobtnam
            // 
            radiobtnam.AutoSize = true;
            radiobtnam.Location = new Point(135, 461);
            radiobtnam.Name = "radiobtnam";
            radiobtnam.Size = new Size(51, 19);
            radiobtnam.TabIndex = 39;
            radiobtnam.TabStop = true;
            radiobtnam.Text = "Nam";
            radiobtnam.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Columns.AddRange(new DataGridViewColumn[] { Column1, Column2, Column3, Column4, Column5 });
            dataGridView1.Location = new Point(40, 68);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(657, 239);
            dataGridView1.TabIndex = 37;
            dataGridView1.CellClick += dataGridView1_CellClick;
            // 
            // btThoat
            // 
            btThoat.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btThoat.Image = (Image)resources.GetObject("btThoat.Image");
            btThoat.ImageAlign = ContentAlignment.MiddleLeft;
            btThoat.Location = new Point(819, 328);
            btThoat.Name = "btThoat";
            btThoat.Size = new Size(112, 51);
            btThoat.TabIndex = 34;
            btThoat.Text = "Thoát";
            btThoat.TextAlign = ContentAlignment.MiddleRight;
            btThoat.UseVisualStyleBackColor = true;
            btThoat.Click += btThoat_Click;
            // 
            // btSua
            // 
            btSua.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btSua.Image = (Image)resources.GetObject("btSua.Image");
            btSua.ImageAlign = ContentAlignment.MiddleLeft;
            btSua.Location = new Point(824, 439);
            btSua.Name = "btSua";
            btSua.Size = new Size(112, 55);
            btSua.TabIndex = 33;
            btSua.Text = "Sửa";
            btSua.TextAlign = ContentAlignment.MiddleRight;
            btSua.UseVisualStyleBackColor = true;
            btSua.Click += btSua_Click;
            // 
            // btXoa
            // 
            btXoa.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btXoa.Image = (Image)resources.GetObject("btXoa.Image");
            btXoa.ImageAlign = ContentAlignment.MiddleLeft;
            btXoa.Location = new Point(701, 438);
            btXoa.Name = "btXoa";
            btXoa.Size = new Size(112, 56);
            btXoa.TabIndex = 35;
            btXoa.Text = "Xóa";
            btXoa.TextAlign = ContentAlignment.MiddleRight;
            btXoa.UseVisualStyleBackColor = true;
            btXoa.Click += btXoa_Click;
            // 
            // btThem
            // 
            btThem.BackColor = SystemColors.ControlLightLight;
            btThem.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btThem.Image = (Image)resources.GetObject("btThem.Image");
            btThem.ImageAlign = ContentAlignment.MiddleLeft;
            btThem.Location = new Point(577, 440);
            btThem.Name = "btThem";
            btThem.Size = new Size(112, 55);
            btThem.TabIndex = 36;
            btThem.Text = "Thêm";
            btThem.TextAlign = ContentAlignment.MiddleRight;
            btThem.UseVisualStyleBackColor = false;
            btThem.Click += btThem_Click;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(40, 461);
            label7.Name = "label7";
            label7.Size = new Size(89, 15);
            label7.TabIndex = 20;
            label7.Text = "Giới tính (Nam)";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(40, 421);
            label6.Name = "label6";
            label6.Size = new Size(60, 15);
            label6.TabIndex = 22;
            label6.Text = "Ngày sinh";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(40, 421);
            label3.Name = "label3";
            label3.Size = new Size(30, 15);
            label3.TabIndex = 21;
            label3.Text = "Tuổi";
            // 
            // textBoxName
            // 
            textBoxName.Location = new Point(107, 371);
            textBoxName.Name = "textBoxName";
            textBoxName.Size = new Size(137, 23);
            textBoxName.TabIndex = 29;
            // 
            // textBox2
            // 
            textBox2.Location = new Point(107, 371);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(137, 23);
            textBox2.TabIndex = 30;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(40, 379);
            label5.Name = "label5";
            label5.Size = new Size(25, 15);
            label5.TabIndex = 24;
            label5.Text = "Tên";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(40, 379);
            label2.Name = "label2";
            label2.Size = new Size(25, 15);
            label2.TabIndex = 23;
            label2.Text = "Tên";
            // 
            // textboxMasv
            // 
            textboxMasv.Location = new Point(107, 328);
            textboxMasv.Name = "textboxMasv";
            textboxMasv.Size = new Size(137, 23);
            textboxMasv.TabIndex = 32;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(40, 336);
            label4.Name = "label4";
            label4.Size = new Size(24, 15);
            label4.TabIndex = 25;
            label4.Text = "Mã";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(107, 328);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(137, 23);
            textBox1.TabIndex = 31;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(40, 336);
            label1.Name = "label1";
            label1.Size = new Size(24, 15);
            label1.TabIndex = 26;
            label1.Text = "Mã";
            // 
            // btLammoi
            // 
            btLammoi.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btLammoi.Image = (Image)resources.GetObject("btLammoi.Image");
            btLammoi.ImageAlign = ContentAlignment.MiddleLeft;
            btLammoi.Location = new Point(660, 329);
            btLammoi.Name = "btLammoi";
            btLammoi.Size = new Size(128, 50);
            btLammoi.TabIndex = 40;
            btLammoi.Text = "Làm mới";
            btLammoi.TextAlign = ContentAlignment.MiddleRight;
            btLammoi.UseVisualStyleBackColor = true;
            btLammoi.Click += btLamMoi_Click;
            // 
            // tbTimkiem
            // 
            tbTimkiem.Location = new Point(623, 23);
            tbTimkiem.Name = "tbTimkiem";
            tbTimkiem.PlaceholderText = "Nhập thông tin tìm kiếm";
            tbTimkiem.Size = new Size(165, 23);
            tbTimkiem.TabIndex = 41;
            // 
            // lbTimkiem
            // 
            lbTimkiem.AutoSize = true;
            lbTimkiem.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            lbTimkiem.Location = new Point(552, 29);
            lbTimkiem.Name = "lbTimkiem";
            lbTimkiem.Size = new Size(65, 15);
            lbTimkiem.TabIndex = 42;
            lbTimkiem.Text = "Tìm kiếm :";
            // 
            // btTimkiem
            // 
            btTimkiem.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            btTimkiem.Location = new Point(794, 23);
            btTimkiem.Name = "btTimkiem";
            btTimkiem.Size = new Size(75, 23);
            btTimkiem.TabIndex = 43;
            btTimkiem.Text = "Tìm";
            btTimkiem.UseVisualStyleBackColor = true;
            btTimkiem.Click += button1_TimKiem_Click;
            // 
            // textBoxTuoi
            // 
            textBoxTuoi.Location = new Point(107, 413);
            textBoxTuoi.Name = "textBoxTuoi";
            textBoxTuoi.Size = new Size(137, 23);
            textBoxTuoi.TabIndex = 29;
            // 
            // pictureBox1
            // 
            pictureBox1.Location = new Point(341, 313);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(230, 216);
            pictureBox1.TabIndex = 44;
            pictureBox1.TabStop = false;
            // 
            // radiobtnu
            // 
            radiobtnu.AutoSize = true;
            radiobtnu.Location = new Point(193, 461);
            radiobtnu.Name = "radiobtnu";
            radiobtnu.Size = new Size(41, 19);
            radiobtnu.TabIndex = 45;
            radiobtnu.TabStop = true;
            radiobtnu.Text = "Nữ";
            radiobtnu.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            button1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            button1.Location = new Point(260, 472);
            button1.Name = "button1";
            button1.Size = new Size(75, 44);
            button1.TabIndex = 46;
            button1.Text = "Browser";
            button1.UseVisualStyleBackColor = true;
            button1.Click += pictureBox1_Click;
            // 
            // btexportexcel
            // 
            btexportexcel.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btexportexcel.Image = (Image)resources.GetObject("btexportexcel.Image");
            btexportexcel.ImageAlign = ContentAlignment.MiddleLeft;
            btexportexcel.Location = new Point(761, 173);
            btexportexcel.Name = "btexportexcel";
            btexportexcel.Size = new Size(128, 50);
            btexportexcel.TabIndex = 47;
            btexportexcel.Text = "Export excel";
            btexportexcel.TextAlign = ContentAlignment.MiddleRight;
            btexportexcel.UseVisualStyleBackColor = true;
            btexportexcel.Click += btexportexcel_Click;
            // 
            // btimportexcel
            // 
            btimportexcel.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btimportexcel.Image = (Image)resources.GetObject("btimportexcel.Image");
            btimportexcel.ImageAlign = ContentAlignment.MiddleLeft;
            btimportexcel.Location = new Point(761, 257);
            btimportexcel.Name = "btimportexcel";
            btimportexcel.Size = new Size(128, 50);
            btimportexcel.TabIndex = 49;
            btimportexcel.Text = "ImportExcel";
            btimportexcel.TextAlign = ContentAlignment.MiddleRight;
            btimportexcel.UseVisualStyleBackColor = true;
            btimportexcel.Click += importExcelButton_Click;
            // 
            // btexportpdf
            // 
            btexportpdf.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            btexportpdf.Image = (Image)resources.GetObject("btexportpdf.Image");
            btexportpdf.ImageAlign = ContentAlignment.MiddleLeft;
            btexportpdf.Location = new Point(761, 98);
            btexportpdf.Name = "btexportpdf";
            btexportpdf.Size = new Size(128, 50);
            btexportpdf.TabIndex = 50;
            btexportpdf.Text = "Export Pdf";
            btexportpdf.TextAlign = ContentAlignment.MiddleRight;
            btexportpdf.UseVisualStyleBackColor = true;
            btexportpdf.Click += btexportpdf_Click;
            // 
            // Column1
            // 
            Column1.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Column1.DataPropertyName = "mssv";
            Column1.HeaderText = "Mã Sinh Viên";
            Column1.Name = "Column1";
            // 
            // Column2
            // 
            Column2.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Column2.DataPropertyName = "name";
            Column2.HeaderText = "Tên Sinh Viên";
            Column2.Name = "Column2";
            // 
            // Column3
            // 
            Column3.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Column3.DataPropertyName = "ngaysinh";
            Column3.HeaderText = "Ngày Sinh";
            Column3.Name = "Column3";
            // 
            // Column4
            // 
            Column4.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Column4.DataPropertyName = "gioitinh";
            Column4.HeaderText = "Giới tính (Nam)";
            Column4.Name = "Column4";
            Column4.Resizable = DataGridViewTriState.True;
            // 
            // Column5
            // 
            Column5.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Column5.DataPropertyName = "hinhanh";
            Column5.HeaderText = "Hinh anh";
            Column5.Name = "Column5";
            Column5.Resizable = DataGridViewTriState.True;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackgroundImage = (Image)resources.GetObject("$this.BackgroundImage");
            ClientSize = new Size(936, 528);
            Controls.Add(btexportpdf);
            Controls.Add(btimportexcel);
            Controls.Add(btexportexcel);
            Controls.Add(button1);
            Controls.Add(radiobtnu);
            Controls.Add(pictureBox1);
            Controls.Add(btTimkiem);
            Controls.Add(lbTimkiem);
            Controls.Add(tbTimkiem);
            Controls.Add(btLammoi);
            Controls.Add(radiobtnam);
            Controls.Add(dataGridView1);
            Controls.Add(btThoat);
            Controls.Add(btSua);
            Controls.Add(btXoa);
            Controls.Add(btThem);
            Controls.Add(label7);
            Controls.Add(label6);
            Controls.Add(label3);
            Controls.Add(textBoxTuoi);
            Controls.Add(textBoxName);
            Controls.Add(textBox2);
            Controls.Add(label5);
            Controls.Add(label2);
            Controls.Add(textboxMasv);
            Controls.Add(label4);
            Controls.Add(textBox1);
            Controls.Add(label1);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private RadioButton radiobtnam;
        private DataGridView dataGridView1;
        private Button btThoat;
        private Button btSua;
        private Button btXoa;
        private Button btThem;
        private Label label7;
        private Label label6;
        private Label label3;
        private TextBox textBoxName;
        private TextBox textBox2;
        private Label label5;
        private Label label2;
        private TextBox textboxMasv;
        private Label label4;
        private TextBox textBox1;
        private Label label1;
        private Button btLammoi;
        private TextBox tbTimkiem;
        private Label lbTimkiem;
        private Button btTimkiem;
        private TextBox textBoxTuoi;
        private PictureBox pictureBox1;
        private RadioButton radiobtnu;
        private Button button1;
        private Button btexportexcel;
        private Button btimportexcel;
        private Button btexportpdf;
        private DataGridViewTextBoxColumn Column1;
        private DataGridViewTextBoxColumn Column2;
        private DataGridViewTextBoxColumn Column3;
        private DataGridViewCheckBoxColumn Column4;
        private DataGridViewTextBoxColumn Column5;
    }
}