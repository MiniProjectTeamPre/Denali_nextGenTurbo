using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Denali_nextGenTurbo {
    public partial class user_id : Form {
        public user_id() {
            InitializeComponent();
        }

        public string pathFile = "UserID.txt";
        private void user_id_Load(object sender, EventArgs e) {

        }
        private void user_id_FormClosed(object sender, FormClosedEventArgs e) {
            if(flag_id_pass) File.WriteAllText(pathFile, textBox1.Text);
            else File.WriteAllText(pathFile, string.Empty);
        }

        private bool flag_id_pass = false;
        private void textBox1_KeyDown(object sender, KeyEventArgs e) {
            //ตรวจสอบเบื้องต้นว่า id นั้นผิดฟอแมท หรือไม่ ถ้าผิดจะเด้งออกทันที
            if (e.KeyCode != Keys.Enter) return;
            if (textBox1.Text.Count() != 5) {
                MessageBox.Show("not format");
                return;
            }

            if (TeamPrecision.PRISM.cUsers.UserLogin(textBox1.Text)) {
                goto idComplete;
            } else {
                MessageBox.Show("_wม่พบข้อมูลในระบบ");
            }

            //ชั่วคราว==============(หรือถาวรวะ แม่งอยู่มานานละ)
            //if (check_id_in_txtFile()) goto SkipRest;

            //ตรวจสอบ id ด้วยระบบ rest ว่ามีอยู่ในระบบ หรือไม่
            //if (check_id_with_rest()) goto idComplete;

            //ถ้า id ไม่มีอยู่ในระบบ จะ return ออกไป ให้ใส่ id ใหม่
            return;

            //เมื่อใส่รหัสเสร็จแล้วโปรแกรมในส่วน ใส่รหัสจะปิดตัวลงเอง
            SkipRest:
            idComplete:
            flag_id_pass = true;
            this.Close();
        }
        private bool check_id_in_txtFile() {
            string[] id = { "81087", "81155", "71027", "71020", "36027", "71099", "71064", "72006", "72009", "36254",
                            "37218", "22134", "71045", "15151", "34899", "33735", "72035", "62019", "60080", "54086",
                            "53431", "71054", "72057", "72098", "30282", "62074"};

            foreach (string id_ in id) {
                if (textBox1.Text == id_) return true;
            }


            List<string> idInFile = new List<string>();
            try {
                string[] hhgh = File.ReadAllLines("up_data_login_id_add_new.txt");
                idInFile = hhgh.ToList<string>();
            } catch { }

            foreach (string id_ in idInFile) {
                if (textBox1.Text == id_) return true;
            }

            MessageBox.Show("ไม่พบข้อมูลในระบบ!!!");
            return false;
        }
        private bool check_id_with_rest() {
            //RestAPI restAPI = new RestAPI();
            //string result = restAPI.getEmployeeID(textBox1.Text);
            //if (result != textBox1.Text) {
            //    MessageBox.Show(result);
            //    return false;
            //}

            return true;
        }

        private void timer1_Tick(object sender, EventArgs e) {
            timer1.Enabled = false;
            this.Activate();

            if (Form.ActiveForm != this) {
                this.WindowState = FormWindowState.Minimized;
                this.WindowState = FormWindowState.Normal;
            }
        }
    }
}
