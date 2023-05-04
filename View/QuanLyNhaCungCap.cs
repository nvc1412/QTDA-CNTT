using QuanLyCuaHangGiaDung.Controller;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLyCuaHangGiaDung.View
{
    public partial class QuanLyNhaCungCap : Form
    {
        NhaCungCapController ncc = new NhaCungCapController();
        Regex isValidInput = new Regex(@"^\d{9,11}$");
        Regex number = new Regex(@"^[-+]?[0-9]*.?[0-9]+$");
        public QuanLyNhaCungCap()
        {
            InitializeComponent();
        }

        private void QuanLyNhaCungCap_Load(object sender, EventArgs e)
        {
            txtMancc.Focus();
            dgvNhacungcap.DataSource = ncc.getData();
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
        }

        public void setNull()
        {
            txtTimkiem.Text = null;
            txtMancc.Text = null;
            txtTenncc.Text = null;
            txtDiachi.Text = null;
            txtSdt.Text = null;
            cbSapxep.Text = null;
            dgvNhacungcap.DataSource = ncc.getData();
            btnThem.Enabled = true;
            btnXuatexcel.Enabled = true;
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            txtMancc.Enabled = true;
        }

        private void btnLammoi_Click(object sender, EventArgs e)
        {
            setNull();
            txtMancc.Focus();
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            if (txtMancc.Text == "" || txtTenncc.Text == "" || txtDiachi.Text == "" || txtSdt.Text == "")
            {
                MessageBox.Show("Không được để trống!");
            }
            else if (ncc.getMaNCC(txtMancc.Text) > 0)
            {
                MessageBox.Show("Mã nhà cung cấp đã tồn tại!");
            }
            else if (!isValidInput.IsMatch(txtSdt.Text))
            {
                MessageBox.Show("Số điện thoại không hợp lệ!");
            }
            else
            {
                string Query = $"INSERT INTO NhaCungCap(MaNCC, TenNCC, DiaChi, Sdt) Values(N'{txtMancc.Text}', N'{txtTenncc.Text}', N'{txtDiachi.Text}', N'{txtSdt.Text}')";
                int sl = ncc.ThemSuaXoaNCC(Query);
                if (sl > 0)
                {
                    MessageBox.Show("Thêm mới thành công!");
                    setNull();
                }
                else
                {
                    MessageBox.Show("Thêm mới thất bại!");
                }
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (txtMancc.Text == "" || txtTenncc.Text == "" || txtDiachi.Text == "" || txtSdt.Text == "")
            {
                MessageBox.Show("Không được để trống!");
            }
            else if (!isValidInput.IsMatch(txtSdt.Text))
            {
                MessageBox.Show("Số điện thoại không hợp lệ!");
            }
            else
            {
                string Query = $"UPDATE NhaCungCap SET TenNCC=N'{txtTenncc.Text}', DiaChi=N'{txtDiachi.Text}', Sdt=N'{txtSdt.Text}' WHERE MaNCC=N'{txtMancc.Text}'";
                int sl = ncc.ThemSuaXoaNCC(Query);
                if (sl > 0)
                {
                    MessageBox.Show("Sửa thành công!");
                    setNull();
                }
                else
                {
                    MessageBox.Show("Sửa thất bại!");
                }
            }
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string Query = $"DELETE FROM NhaCungCap WHERE MaNCC=N'{txtMancc.Text}'";
            int sl = ncc.ThemSuaXoaNCC(Query);
            if (sl > 0)
            {
                MessageBox.Show("Xóa thành công!");
                setNull();
            }
            else
            {
                MessageBox.Show("Xóa thất bại!");
            }
        }

        private void btnXuatexcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //gọi hàm ToExcel() với tham số là dgvNhanvien và filename từ SaveFileDialog
                ncc.ToExcel(dgvNhacungcap, saveFileDialog1.FileName);
            }
        }

        private void dgvNhacungcap_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            btnSua.Enabled = true;
            btnXoa.Enabled = true;
            btnThem.Enabled = false;
            btnXuatexcel.Enabled = false;
            txtMancc.Enabled = false;
            try
            {
                int row = e.RowIndex;
                int col = e.ColumnIndex;
                if (row >= 0)
                {
                    txtMancc.Text = (string)dgvNhacungcap.Rows[row].Cells["MaNCC"].Value;
                    txtTenncc.Text = (string)dgvNhacungcap.Rows[row].Cells["TenNCC"].Value;
                    txtDiachi.Text = (string)dgvNhacungcap.Rows[row].Cells["DiaChi"].Value;
                    txtSdt.Text = (string)dgvNhacungcap.Rows[row].Cells["Sdt"].Value;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi: " + ex.Message);
            }
        }

        private void btnTimkiem_Click(object sender, EventArgs e)
        {
            string Query = $"SELECT * FROM NhaCungCap WHERE MaNCC LIKE N'%{txtTimkiem.Text}%' OR TenNCC LIKE N'%{txtTimkiem.Text}%'" +
                $" OR DiaChi LIKE N'%{txtTimkiem.Text}%' OR Sdt LIKE N'%{txtTimkiem.Text}%'";
            dgvNhacungcap.DataSource = ncc.TimNCC(Query);
        }

        private void cbSapxep_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sapxep = "MaNCC";
            switch (cbSapxep.SelectedIndex)
            {
                case 0: sapxep = "MaNCC"; break;
                case 1: sapxep = "MaNCC DESC"; break;
                default: sapxep = "MaNCC"; break;
            }
            dgvNhacungcap.DataSource = ncc.SapXepNCC(sapxep);
        }
    }
}
