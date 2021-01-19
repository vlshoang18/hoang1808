using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Media;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Threading;


namespace ailatrieuphu_1
{
    public partial class fplay : Form
    {

        public fplay()
        {
            InitializeComponent();
        }
        private cauhoi cauHoiHienTai = null;
        List<cauhoi> danhSachCauHoi;
        int index = 0;
        private void laycauhoi()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;


            int rw = 0; //số hàng tối đa có trong sheet
            int cl = 0; //số cột tối đa có trong sheet

            danhSachCauHoi = new List<cauhoi>();
           
            xlApp = new Excel.Application();
            string url = Application.StartupPath + "\\cau-hoi-1.xlsx";
            xlWorkBook = xlApp.Workbooks.Open(url, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;      //Đếm số hàng
            cl = range.Columns.Count;   //Đếm số cột 
            for (int i = 2; i <= 11; i++)
            {

                danhSachCauHoi.Add(new cauhoi(
                 (string)(range.Cells[i, 1] as Excel.Range).Value2,
                (range.Cells[i, 2] as Excel.Range).Value2.ToString(),
                 (range.Cells[i, 3] as Excel.Range).Value2.ToString(),
                (range.Cells[i, 4] as Excel.Range).Value2.ToString(),
                (range.Cells[i, 5] as Excel.Range).Value2.ToString(),
                 (range.Cells[i, 6] as Excel.Range).Value2.ToString()
                 ));
               
            }
          
        xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            

        }
        private void loadcauhoi(int index)
        {
            thoigian = 40;
            tmthoigian.Start();
            btnA.Visible = true;
            btnB.Visible = true;
            btnC.Visible = true;
            btnD.Visible = true;
         
            if (index < danhSachCauHoi.Count)
            {
                cauHoiHienTai = danhSachCauHoi[index];
                tbcauhoi.Text = cauHoiHienTai.cauHoi;
                btnA.Text = cauHoiHienTai.A;
                btnB.Text = cauHoiHienTai.B;
                btnC.Text = cauHoiHienTai.C;
                btnD.Text = cauHoiHienTai.D;
            }
            else
            {//đã trả lời hết câu hỏi
             //Nhận thưởng
                fnhanthuong f = new fnhanthuong();
                this.Hide();
                f.ShowDialog();
            }
            if (index == 0)
            {
                btn1.BackColor = Color.Red;
            }
            if (index == 1)
            {
                btn2.BackColor = Color.Red;
            }
            if (index == 2)
            {
                btn3.BackColor = Color.Red;
            }
            if (index == 3)
            {
                btn4.BackColor = Color.Red;
            }
            if (index == 4)
            {
                btn5.BackColor = Color.Red;
            }
            if (index == 5)
            {
                btn6.BackColor = Color.Red;
            }
            if (index == 6)
            {
                btn7.BackColor = Color.Red;
            }
            if (index == 7)
            {
                btn8.BackColor = Color.Red;
            }
            if (index == 8)
            {
                btn9.BackColor = Color.Red;
            }
            if (index == 9)
            {
                btn10.BackColor = Color.Red;
            }
           
        }
        private void thongbaosai()
            
        {
            Thread.Sleep(2000);
            btndunglai.Visible = false;
            pntrogiup.Visible = false;
            btnA.Visible = false;
            btnB.Visible = false;
            btnC.Visible = false;
            btnD.Visible = false;
            btndiem.Visible = true;
                btndiem.Text = "Tiền thưởng nhận được là:0$";
            btndapan.Visible = true;
            btndapan.Text = "Trả lời sai !!"+"Đáp án là:"+cauHoiHienTai.dapan;
            tmthoigian.Stop();
        }
        int demdiem = 0;
        int tinhdiem(int x)
        {
            if (x == 0)
                return 100000;
            if (x == 1)
                return 200000;
            if (x == 2)
                return 400000;
            if (x == 3)
                return 800000;
            if (x == 4)
                return 1000000;
            if (x == 5)
                return 2000000;
            if (x == 6)
                return 4000000;
            if (x == 7)
                return 8000000;
            if (x == 8)
                return 20000000;
            if (x == 9)
                return 50000000;
            else
            return 0;
        }
        private void thongbaodung()
        {
            Thread.Sleep(2000);
            btndapan.Visible = true;
            btndapan.Text = "Trả lời đúng !!";
            btndiem.Visible = true;
            btndiem.Text ="Tiền thưởng hiện tại là:"+tinhdiem(demdiem).ToString() + "$";
            demdiem++;
            tmthoigian.Stop();
        }
        private void fplay_Load(object sender, EventArgs e)
        {
            laycauhoi();
            
            loadcauhoi(index);
        }
        private void lbcauhoi_Click(object sender, EventArgs e)
        {

        }
        private void button11_Click(object sender, EventArgs e)
        {
           

        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btncauhoi_Click(object sender, EventArgs e)
        {
          
        }
        private void grbtrogiup_Enter(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {

        }
        private void btndunglai_Click(object sender, EventArgs e)
        {
            lbthongbao.Visible = true;
            lbthongbao.Text = btndiem.Text;
            grbhelp.Visible = true;
            btnhettime.Visible = true;
            btndunglai.Visible = false;
            pntrogiup.Visible = false;
            btndunglai.Visible = false;
            btneexit.Visible = false;
            btnA.Enabled = false;
            btnB.Enabled = false;
            btnC.Enabled = false;
            btnD.Enabled = false;
            tmthoigian.Stop();
                
        }

        private void btneexit_Click(object sender, EventArgs e)
        {
            fmain f = new fmain();
            this.Hide();
            f.ShowDialog();
            this.Show();
        }

        private void BtnA_Click(object sender, EventArgs e)
        {
            lbthongbao.Visible = false;
            btnA.BackColor = Color.Yellow;
            btnA.Refresh();
            if (cauHoiHienTai.dapan == btnA.Text)
            {
                thongbaodung();
                index++;
                loadcauhoi(index);
                btnA.BackColor = SystemColors.Control;

            }
            else
            {
                thongbaosai();

            }
           
        }

        private void BtnB_Click(object sender, EventArgs e)
        {
            lbthongbao.Visible = false;
            btnB.BackColor = Color.Yellow;
            btnB.Refresh();
            if (cauHoiHienTai.dapan == btnB.Text)
            {
                thongbaodung();
                index++;
                loadcauhoi(index);
                btnB.BackColor = SystemColors.Control;
            }
            else
            {
                thongbaosai();

            }
           
        }

        private void BtnC_Click(object sender, EventArgs e)
        {
            lbthongbao.Visible = false;
            btnC.BackColor = Color.Yellow;
            btnC.Refresh();
            if (cauHoiHienTai.dapan == btnC.Text)
            {
                thongbaodung();
                index++;
                loadcauhoi(index);
                btnC.BackColor = SystemColors.Control;
            }
            else
            {
                thongbaosai();

            }
           
        }

        private void BtnD_Click(object sender, EventArgs e)
        {
            lbthongbao.Visible = false;
            btnD.BackColor = Color.Yellow;
            btnD.Refresh();
            if (cauHoiHienTai.dapan == btnD.Text)
            {
                thongbaodung();
                index++;
                loadcauhoi(index);
                btnD.BackColor = SystemColors.Control;
            }
            else
            {
                thongbaosai();

            }
           
        }

        private void btn5050_Click(object sender, EventArgs e)
        {
            btn5050.Visible = false;
            if (cauHoiHienTai.dapan == btnA.Text)
            {
                btnB.Visible = false;
                btnB.Text = "";
                btnC.Visible = false;
                btnC.Text = "";

            }

            if (btndapan.Text == btnB.Text)
            {

                btnA.Visible = false;
                btnA.Text = "";
                btnD.Visible = false;
                btnD.Text = "";

            }

            if (btndapan.Text == btnC.Text)
            {
                btnB.Visible = false;
                btnB.Text = "";
                btnD.Visible = false;
                btnD.Text = "";

            }

            if (btndapan.Text == btnD.Text)
            {

                btnB.Visible = false;
                btnB.Text = "";
                btnA.Visible = false;
                btnA.Text = "";

            }
        }
        int thoigiangoi;
        private void btngoidien_Click(object sender, EventArgs e)
        {
            btngoidien.Visible = false;
            thoigiangoi = 0;
            tmgoidien.Start();
        }

        private void tmgoidien_Tick(object sender, EventArgs e)
        {
            int demcau = 0;
            demcau = index + 1;
            thoigiangoi++;
            grbhelp.Visible = true;
            lbhelp.Visible = true;
            lbhelp.ForeColor = Color.Red;
            if (thoigiangoi == 10)
            {
                lbhelp.Text = "-Hệ thống đang kết nối đến số điện thoại yêu cầu...";
            }
            if (thoigiangoi == 50)
            {
                lbhelp.Text = "-Alo...";
            }
            if (thoigiangoi == 60)
            {
                lbhelp.Text += "\n-Hi, mình đang ở câu hỏi số " + demcau.ToString() + " và cần bạn trợ giúp";
            }

            if (thoigiangoi == 90)
            {
                lbhelp.Text = "--Câu hỏi là:  \n" + tbcauhoi.Text + "\nA: " + btnA.Text + "   B: " + btnB.Text + "    C: " + btnC.Text + " D: " + btnD.Text;
            }
            if (thoigiangoi == 150)
            {
                lbhelp.Text = "-Theo mình thì đáp án là: " + cauHoiHienTai.dapan;
            }
            if (thoigiangoi == 180)
            {
                lbhelp.Text = "\n-Chắc chắn không?\n-Chắc chắn mà, Vì mình gặp câu này rồi\n-Cám ơn nhé!?\n-Không có gì!!!";
            }
            if (thoigiangoi == 200)
            {
                lbhelp.Text = "";
                grbhelp.Visible = false;
                tmgoidien.Stop();
                tmthoigian.Start();
            }
        }

        private void grbhelp_Enter(object sender, EventArgs e)
        {

        }
        int thoigiankhangia = 0;
        private void tmkhangia_Tick(object sender, EventArgs e)
        {
            thoigiankhangia++;
            grbhelp.Visible = true;
            lbhelp.Visible = true;
            lbhelp.ForeColor = Color.Red;
            lbhelp.Text = "Đang lấy ý kiến khán giả trường quay.\n\n";

            if (thoigiankhangia > 50)
            {
                if (cauHoiHienTai.dapan == btnA.Text)
                {
                    lbhelp.Text += "            \nA:   70%    B:    13%   C:     7%    D:    10%";

                }
                if (cauHoiHienTai.dapan == btnB.Text)
                {
                    lbhelp.Text += "         \nA:    5%    B:    60%     C:     25%    D:    10%";

                }
                if (cauHoiHienTai.dapan == btnC.Text)
                {
                    lbhelp.Text += "         \nA:   40%     B:    0%     C:     60%     D:    0%";

                }
                if (cauHoiHienTai.dapan == btnD.Text)
                {
                    lbhelp.Text += "          \nA:   3%       B:    7%      C:     10%    D:    80%";

                }
            }

            if (thoigiankhangia == 140)
            {
                grbhelp.Visible = false;
                lbhelp.Text = "";
                tmkhangia.Stop();
                tmthoigian.Start();

            }
        }

        private void btnkhangia_Click(object sender, EventArgs e)
        {
            lbhelp.Text = "";
            thoigiankhangia = 0;
            tmkhangia.Start();
          //  btnkhangia.Enabled = false;
            btnkhangia.Visible = false;
        }
      
        private void tmdapan_Tick(object sender, EventArgs e)
        {
            

        }
        int thoigian = 40;
        private void tmthoigian_Tick(object sender, EventArgs e)
        {
            tmthoigian.Start();
            lbthoigian.Text = (thoigian).ToString();
            thoigian--;
          
            if (thoigian < 0)
            {
                lbthoigian.Text = "0";

            }
            if (thoigian < -1)
            {
                grbhelp.Visible = true;
                lbthongbao.Visible = true;
                lbthongbao.Text = "\n\nBạn đã hết thời gian!!!!!!!!!";
                btnA.Enabled = false;
                btnB.Enabled = false;
                btnC.Enabled = false;
                btnD.Enabled = false;
            }
            if(thoigian==-5)
            {
                grbhelp.Visible = true;
                lbthongbao.Visible = true;
                lbthongbao.Text = "\n-Cảm ơn bạn đã đến với chương trình của chúng tôi!!";
                btnhettime.Visible = true;
                tmthoigian.Stop();
            }
            
        }

        private void btnhettime_Click(object sender, EventArgs e)
        {
            fmain f = new fmain();
            this.Hide();
            f.ShowDialog();
            this.Show();
        }

        private void fplay_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.A)
            {
                 btnA.PerformClick();
            }
            if (e.KeyCode == Keys.B)
            {
                btnB.PerformClick();
            }
            if (e.KeyCode == Keys.C)
            {
                btnC.PerformClick();
            }
            if (e.KeyCode == Keys.D)
            {
                btnD.PerformClick();
            }

        }
    }
 }   

