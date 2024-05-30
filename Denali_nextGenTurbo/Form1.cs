using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.OCR;
using Emgu.CV.Structure;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace Denali_nextGenTurbo {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }
        private bool payTest = false;  //set เป็น true ถ้าหากจะรันในคอมส่วนตัวที่ขึ้น prism ไม่ได้
        [System.Runtime.InteropServices.DllImport("user32")]
        public static extern int GetAsyncKeyState(int vKey);
        public SetupPay.FormPay setupPay = new SetupPay.FormPay();
        public bool setUpPayFlag = false;
        private bool focusFormTick = false;
        public PrismTest prismTest = new PrismTest();
        public DMM dmm = new DMM();
        private ImageClass imageClass = new ImageClass();
        private Position position = new Position();
        private LogFile logFile = new LogFile();
        private TestSN testSN = new TestSN();
        private AutoItX3Lib.AutoItX3 autoit = new AutoItX3Lib.AutoItX3();
        private SaveDataLogCsv saveDataCsv = new SaveDataLogCsv();
        private Focus focus = new Focus();
        private Tester tester = new Tester();


        #region============================================= Display_Message ==============================================
        private Color[] LogMsgTypeColor = { Color.Blue, Color.Green, Color.Black, Color.Orange, Color.Red };
        public enum LogMsgType { Incoming_Blue, Outgoing_Green, Normal_Black, Warning_Orange, Error_Red };
        public void Log(LogMsgType msgtype, string msg) {
            try {
                rtb_log.Invoke(new EventHandler(delegate {
                    if (rtb_log.Text.Length > rtb_log.MaxLength) {
                        rtb_log.Text = string.Empty;
                    }

                    rtb_log.SelectedText = string.Empty;
                    rtb_log.SelectionFont = new Font(rtb_log.SelectionFont, FontStyle.Bold);
                    rtb_log.SelectionColor = LogMsgTypeColor[(int)msgtype];
                    rtb_log.AppendText(msg);
                    rtb_log.ScrollToCaret();
                }));
            } catch { }
        }

        #endregion

        #region====================================================== Event ======================================================
        private void Form1_Load(object sender, EventArgs e) {
            Process[] pname = Process.GetProcessesByName("Denali_nextGenTurbo");
            if (pname.Length == 2) {
                MessageBox.Show("_โปรแกรมนี้ เปิดใช้งานอยู่");
                Application.Exit();
                return;
            }

            GetEmployeeID();
            setup();

            tb_wo.Focus();
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e) {
            Save_vSrollBar();

            setupPay.write_text(position.headConfig.positionFormOpenX, this.Location.X.ToString(), position.nameFile);
            setupPay.write_text(position.headConfig.positionFormOpenY, this.Location.Y.ToString(), position.nameFile);
        }
        private void tb_wo_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter) return;

            tb_wo.Text = tb_wo.Text.ToUpper();
            prismTest.wo = tb_wo.Text;

            try {
                string[] getWO;

                if (payTest) {
                    getWO = new string[]{
                    "SUCCESS",
                    "FG-63T305-LF",
                    "sadflasdjflajdslk",
                    "123",
                    "5"
                    };

                } else {
                    getWO = TeamPrecision.PRISM.cSNs.getWO(prismTest.wo, prismTest.xml.processName);
                }
                
                if (getWO[0] == prismTest.success) {
                    prismTest.result = getWO[0];
                    prismTest.fg = getWO[1];
                    prismTest.desc = getWO[2];
                    prismTest.orderQty = getWO[3];
                    prismTest.outputQty = getWO[4];
                    tb_fg.Text = prismTest.fg;
                    tb_desc.Text = prismTest.desc;
                    tb_orderQty.Text = (Convert.ToDouble(prismTest.orderQty)).ToString();
                    tb_outputQty.Text = (Convert.ToDouble(prismTest.outputQty)).ToString();
                    tb_customer.Text = "SST";
                    GetSpecMinMax();
                    CreateFolderDataLog2();

                    tb_sn.Focus();
                } else {
                    MessageBox.Show(prismTest.messageErr.getWO);
                    MessageBox.Show(prismTest.fg + "\r\n" + prismTest.desc);
                }
            } catch {
                MessageBox.Show(prismTest.messageErr.getWO);
                MessageBox.Show(prismTest.fg + "\r\n" + prismTest.desc);
            }

            //if (function_timeout(prismTest.getWO, 5000) == prismTest.success) {
            //    tb_fg.Text = prismTest.fg;
            //    tb_desc.Text = prismTest.desc;
            //    tb_orderQty.Text = (Convert.ToDouble(prismTest.orderQty)).ToString();
            //    tb_outputQty.Text = (Convert.ToDouble(prismTest.outputQty)).ToString();
            //    tb_customer.Text = "SST";

            //    GetSpecMinMax();
            //    CreateFolderDataLog2();

            //    tb_sn.Focus();

            //} else {
            //    MessageBox.Show(prismTest.messageErr.getWO);
            //    MessageBox.Show(prismTest.fg + "\r\n" + prismTest.desc);
            //}
        }
        private void tb_sn_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter) return;

            //if (tb_wo.Text.Length == 0) {
            //    MessageBox.Show("กรุณาใส่ WO ก่อน");
            //    tb_sn.Text = String.Empty;
            //    return;
            //}

            if (!CheckSN()) {
                return;
            }

            if (!CheckSnDouble()) {
                return;
            }

            SendToSN();

            tm_scanSN.Enabled = true;
        }
        private void tb_setTestTime_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter) return;

            string time = tb_setTestTime.Text;
            try {
                int timer = Convert.ToInt32(time);
            } catch {

                MessageBox.Show(logFile.mesSage.error);
                return;
            }

            setupPay.write_text(logFile.headConfig.setTestTime, time, logFile.nameFile);
            tb_setTestTime.BackColor = Color.Lime;
            DelaymS(100);
            tb_setTestTime.BackColor = Color.White;
        }
        private void tb_timeOutImage_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter) return;

            string timeOut = tb_timeOutImage.Text;
            try {
                int timer = Convert.ToInt32(timeOut);
            } catch {

                MessageBox.Show(logFile.mesSage.error);
                return;
            }

            setupPay.write_text(testSN.headConfig.setTimeOut, timeOut, position.nameFile);
            tb_timeOutImage.BackColor = Color.Lime;
            DelaymS(100);
            tb_timeOutImage.BackColor = Color.White;
        }

        private void bt_imageRefer_Click(object sender, EventArgs e) {
            SetFormImage();
            SetEventImage_GetImage();
        }
        private void bt_position1_Click(object sender, EventArgs e) {
            imageClass.position.head = imageClass.position.head1;
            SetFormImage();
            SetEventImage_GetPositionClick();
        }
        private void bt_position2_Click(object sender, EventArgs e) {
            imageClass.position.head = imageClass.position.head2;
            SetFormImage();
            SetEventImage_GetPositionClick();
        }
        private void bt_position3_Click(object sender, EventArgs e) {
            imageClass.position.head = imageClass.position.head3;
            SetFormImage();
            SetEventImage_GetPositionClick();
        }
        private void bt_position4_Click(object sender, EventArgs e) {
            imageClass.position.head = imageClass.position.head4;
            SetFormImage();
            SetEventImage_GetPositionClick();
        }
        private void bt_positionExecute_Click(object sender, EventArgs e) {
            imageClass.position.head = imageClass.position.execute;
            SetFormImage();
            SetEventImage_GetPositionClick();
        }
        private void bt_exeCute_Click(object sender, EventArgs e) {
            if (!CheckSnBeforeClick()) {
                MessageBox.Show("กรูณาสแกน SN ให้ครบก่อน");
                return;
            }

            bt_exeCute.Enabled = false;

            if (!tester.debugTestBattery) {
                this.WindowState = FormWindowState.Minimized;
                logFile.Clear(this);
                DelaymS(1000);
                autoit.MouseClick("LEFT", 0, 0, 1, -1);
                DelaymS(500);

                SnToTurBo();

                DelaymS(dmm.spec.delayWaitExeCute);
                ClickToExeCute();
                this.WindowState = FormWindowState.Normal;
            }

            tb_sn.Focus();
            DelaymS(50);

            SetFlagLogFile();
            logFile.stopWatch.Restart();
            tm_timeOutLog.Enabled = true;
        }
        private void bt_pathLogXml_Click(object sender, EventArgs e) {
            FolderBrowserDialog browser = new FolderBrowserDialog();

            if (browser.ShowDialog() == DialogResult.OK) {
                logFile.pathLog = browser.SelectedPath + "\\";
                setupPay.write_text(logFile.headConfig.pathLog, logFile.pathLog, logFile.nameFile);
            }
        }
        private void bt_clearStatusAll_Click(object sender, EventArgs e) {
            dmm.statusRunning[0] = false;
            dmm.statusRunning[1] = false;
            dmm.statusRunning[2] = false;
            dmm.statusRunning[3] = false;
            DelaymS(50);

            lb_status1.Text = imageClass.close;
            lb_status2.Text = imageClass.close;
            lb_status3.Text = imageClass.close;
            lb_status4.Text = imageClass.close;
            lb_status1.BackColor = Color.Red;
            lb_status2.BackColor = Color.Red;
            lb_status3.BackColor = Color.Red;
            lb_status4.BackColor = Color.Red;

            tb_snHead1.Text = string.Empty;
            tb_snHead2.Text = string.Empty;
            tb_snHead3.Text = string.Empty;
            tb_snHead4.Text = string.Empty;
            tb_snHead1.BackColor = Color.White;
            tb_snHead2.BackColor = Color.White;
            tb_snHead3.BackColor = Color.White;
            tb_snHead4.BackColor = Color.White;
            tb_snHead1.Visible = false;
            tb_snHead2.Visible = false;
            tb_snHead3.Visible = false;
            tb_snHead4.Visible = false;

            tb_voltHead1.Text = string.Empty;
            tb_voltHead2.Text = string.Empty;
            tb_voltHead3.Text = string.Empty;
            tb_voltHead4.Text = string.Empty;
            tb_voltHead1.BackColor = Color.White;
            tb_voltHead2.BackColor = Color.White;
            tb_voltHead3.BackColor = Color.White;
            tb_voltHead4.BackColor = Color.White;
            tb_voltHead1.Visible = false;
            tb_voltHead2.Visible = false;
            tb_voltHead3.Visible = false;
            tb_voltHead4.Visible = false;

            testSN.flagTurbo[0] = false;
            testSN.flagTurbo[1] = false;
            testSN.flagTurbo[2] = false;
            testSN.flagTurbo[3] = false;
            testSN.flagReady[0] = true;
            testSN.flagReady[1] = true;
            testSN.flagReady[2] = true;
            testSN.flagReady[3] = true;
            SetFinalLabelVisable(1);
            SetFinalLabelVisable(2);
            SetFinalLabelVisable(3);
            SetFinalLabelVisable(4);

            imageClass.flagCheck = true;
            logFile.readLogFull = true;
            prismTest.flagWaitUpPrism = false;
            dmm.flagRead = false;
        }

        private void vScrollBar1_ValueChanged(object sender, EventArgs e) {
            int inVert = 0;

            inVert = vScrollBar1.Maximum - vScrollBar1.Value;
            imageClass.scrollBar = inVert;
            lb_vScrollBar.Text = inVert.ToString();

            setupPay.write_text(position.headConfig.vScrollBar, inVert.ToString(), position.nameFile);
        }
        private void bt_setPrism_Click(object sender, EventArgs e) {
            setUpPayFlag = true;
            setupPay.show_dialog();
            setUpPayFlag = false;

            prismTest.ReadConfig(setupPay);
            prismTest.xml.WriteXml();
        }
        private void tm_key_Tick(object sender, EventArgs e) {
            ShowPosition();

            GetAsyncKeyState(Convert.ToInt32(Keys.LControlKey));
            GetAsyncKeyState(Convert.ToInt32(Keys.RControlKey));

            if (Convert.ToBoolean(GetAsyncKeyState(Convert.ToInt32(Keys.LControlKey))) ||
                Convert.ToBoolean(GetAsyncKeyState(Convert.ToInt32(Keys.RControlKey)))) {

                if (imageClass.position.timerKeyCtrl < 5) {
                    imageClass.position.timerKeyCtrl++;
                    return;

                }
            } else {
                imageClass.position.timerKeyCtrl = 0;
                return;
            }


            if (Convert.ToBoolean(GetAsyncKeyState(Convert.ToInt32(Keys.LControlKey))) ||
                Convert.ToBoolean(GetAsyncKeyState(Convert.ToInt32(Keys.RControlKey)))) {

                imageClass.form.Close();
                Thread.Sleep(100);
                this.Show();
                tm_key.Enabled = false;
            }
        }
        private void tm_scanSN_Tick(object sender, EventArgs e) {
            tm_scanSN.Enabled = false;

            testSN.scanSnSupport = string.Empty;
        }
        private void tm_timeOutLog_Tick(object sender, EventArgs e) {
            int timeSup = (int)((logFile.testTime - (int)logFile.stopWatch.ElapsedMilliseconds) / 1000);
            TimeSpan span = TimeSpan.FromSeconds(timeSup);
            bt_exeCute.Text = "EXECUTE " + span.ToString();

            if (logFile.stopWatch.ElapsedMilliseconds > logFile.testTime) {
                tm_timeOutLog.Enabled = false;
                bt_exeCute.Text = "EXECUTE";

                imageClass.flagCheck = true;
                logFile.stopWatch.Stop();

                if (!logFile.head1) {
                    tb_snHead1.BackColor = Color.Red;
                    SetFinalLabelFail(1);
                }
                if (!logFile.head2) {
                    tb_snHead2.BackColor = Color.Red;
                    SetFinalLabelFail(2);
                }
                if (!logFile.head3) {
                    tb_snHead3.BackColor = Color.Red;
                    SetFinalLabelFail(3);
                }
                if (!logFile.head4) {
                    tb_snHead4.BackColor = Color.Red;
                    SetFinalLabelFail(4);
                }
            }
        }
        private void tm_focusForm_Tick(object sender, EventArgs e) {
            focusFormTick = true;
        }

        private void tb_snHead1_KeyPress(object sender, KeyPressEventArgs e) {
            e.Handled = true;
        }
        private void tb_snHead2_KeyPress(object sender, KeyPressEventArgs e) {
            e.Handled = true;
        }
        private void tb_snHead3_KeyPress(object sender, KeyPressEventArgs e) {
            e.Handled = true;
        }
        private void tb_snHead4_KeyPress(object sender, KeyPressEventArgs e) {
            e.Handled = true;
        }
        private void tb_voltHead1_KeyPress(object sender, KeyPressEventArgs e) {
            e.Handled = true;
        }
        private void tb_voltHead2_KeyPress(object sender, KeyPressEventArgs e) {
            e.Handled = true;
        }
        private void tb_voltHead3_KeyPress(object sender, KeyPressEventArgs e) {
            e.Handled = true;
        }
        private void tb_voltHead4_KeyPress(object sender, KeyPressEventArgs e) {
            e.Handled = true;
        }
        #endregion

        #region====================================================== Function ======================================================
        public string function_timeout(Func<string> function, int timeout) {
            Task<string> task = Task.Run(function);
            if (task.Wait(timeout)) return task.Result;
            else return "over timeout";
        }
        public void DelaymS(int mS) {
            Stopwatch stopwatchDelaymS = new Stopwatch();
            stopwatchDelaymS.Restart();
            while (mS > stopwatchDelaymS.ElapsedMilliseconds) {
                if (!stopwatchDelaymS.IsRunning) stopwatchDelaymS.Start();
                Application.DoEvents();
            }
            stopwatchDelaymS.Stop();
        }
        private void setup() {
            setupPay.SelectTab = SetupPay.tabPage.TAB1;
            setupPay.set_nameTab(prismTest.nameFile);
            setupPay.SelectTab = SetupPay.tabPage.TAB2;
            setupPay.set_nameTab(position.nameFile);
            setupPay.SelectTab = SetupPay.tabPage.TAB3;
            setupPay.set_nameTab(logFile.nameFile);
            setupPay.SelectTab = SetupPay.tabPage.TAB4;
            setupPay.set_nameTab(dmm.spec.nameFile);
            setupPay.SelectTab = SetupPay.tabPage.TAB5;
            setupPay.set_nameTab(saveDataCsv.nameFile);
            setupPay.SelectTab = SetupPay.tabPage.TAB6;
            setupPay.set_nameTab(tester.nameFile);
            setupPay.setup();

            Set_vScrollBar();
            GetPositionConfig();
            GetTimeOutImage();
            imageClass.GetStatus(this);
            prismTest.ReadConfig(setupPay);
            prismTest.xml.WriteXml();
            testSN.length = Convert.ToInt32(prismTest.xml.digitSN);
            GetLogConfig();
            GetSpec();
            CreateFolderDataLog();
            tester.GetConfig(this);

            RunBackGroundWorker();
            ColorFormLoad();
        }
        private void ColorFormLoad() {
            if (prismTest.mode == PrismTest.Define.debug) {
                this.BackColor = Color.Gold;
                tab_test.BackColor = Color.Gold;
                tab_setup.BackColor = Color.Gold;
                tab_checkImage.BackColor = Color.Gold;
                tab_log.BackColor = Color.Gold;
            }
        }
        private void GetPositionConfig() {
            position.number1.X = Convert.ToInt32(setupPay.read_text(position.headConfig.position1_X, position.nameFile));
            position.number1.Y = Convert.ToInt32(setupPay.read_text(position.headConfig.position1_Y, position.nameFile));
            position.number1.Width = Convert.ToInt32(setupPay.read_text(position.headConfig.position1_Width, position.nameFile));
            position.number1.Height = Convert.ToInt32(setupPay.read_text(position.headConfig.position1_Height, position.nameFile));
            position.number2.X = Convert.ToInt32(setupPay.read_text(position.headConfig.position2_X, position.nameFile));
            position.number2.Y = Convert.ToInt32(setupPay.read_text(position.headConfig.position2_Y, position.nameFile));
            position.number2.Width = Convert.ToInt32(setupPay.read_text(position.headConfig.position2_Width, position.nameFile));
            position.number2.Height = Convert.ToInt32(setupPay.read_text(position.headConfig.position2_Height, position.nameFile));
            position.number3.X = Convert.ToInt32(setupPay.read_text(position.headConfig.position3_X, position.nameFile));
            position.number3.Y = Convert.ToInt32(setupPay.read_text(position.headConfig.position3_Y, position.nameFile));
            position.number3.Width = Convert.ToInt32(setupPay.read_text(position.headConfig.position3_Width, position.nameFile));
            position.number3.Height = Convert.ToInt32(setupPay.read_text(position.headConfig.position3_Height, position.nameFile));
            position.number4.X = Convert.ToInt32(setupPay.read_text(position.headConfig.position4_X, position.nameFile));
            position.number4.Y = Convert.ToInt32(setupPay.read_text(position.headConfig.position4_Y, position.nameFile));
            position.number4.Width = Convert.ToInt32(setupPay.read_text(position.headConfig.position4_Width, position.nameFile));
            position.number4.Height = Convert.ToInt32(setupPay.read_text(position.headConfig.position4_Height, position.nameFile));
            position.number5.X = Convert.ToInt32(setupPay.read_text(position.headConfig.position5_X, position.nameFile));
            position.number5.Y = Convert.ToInt32(setupPay.read_text(position.headConfig.position5_Y, position.nameFile));
            position.number5.Width = Convert.ToInt32(setupPay.read_text(position.headConfig.position5_Width, position.nameFile));
            position.number5.Height = Convert.ToInt32(setupPay.read_text(position.headConfig.position5_Height, position.nameFile));
            position.execute.X = Convert.ToInt32(setupPay.read_text(position.headConfig.positionExecute_X, position.nameFile));
            position.execute.Y = Convert.ToInt32(setupPay.read_text(position.headConfig.positionExecute_Y, position.nameFile));
            position.positionFormOpenX = Convert.ToInt32(setupPay.read_text(position.headConfig.positionFormOpenX, position.nameFile));
            position.positionFormOpenY = Convert.ToInt32(setupPay.read_text(position.headConfig.positionFormOpenY, position.nameFile));
            this.Location = new Point(position.positionFormOpenX, position.positionFormOpenY);
        }
        private void GetTimeOutImage() {
            testSN.setTimeOut = Convert.ToInt32(setupPay.read_text(testSN.headConfig.setTimeOut, position.nameFile));
            tb_timeOutImage.Text = testSN.setTimeOut.ToString();
        }
        private void Set_vScrollBar() {
            vScrollBar1.Maximum = Screen.PrimaryScreen.Bounds.Width;

            int inVert = Convert.ToInt32(setupPay.read_text(position.headConfig.vScrollBar, position.nameFile));
            vScrollBar1.Value = vScrollBar1.Maximum - inVert;
        }
        private void Save_vSrollBar() {
            setupPay.write_text(position.headConfig.vScrollBar, lb_vScrollBar.Text, position.nameFile);
        }
        private bool CheckSN() {
            if (tb_sn.Text.Length != testSN.length) {
                MessageBox.Show("Not " + testSN.length + " Digit!!");
                tb_sn.Text = string.Empty;
                return false;
            }

            return true;
        }
        private bool CheckSnDouble() {
            if (tb_sn.Text == testSN.scanSnSupport) {
                tb_sn.Text = string.Empty;
                return false;
            }

            testSN.scanSnSupport = tb_sn.Text;
            return true;
        }
        private void SendToSN() {
            if (testSN.flagTurbo[0] && testSN.flagReady[0]) {
                tb_snHead1.Text = tb_sn.Text;
                tb_sn.Text = string.Empty;
                testSN.flagReady[0] = false;
                return;
            }

            if (testSN.flagTurbo[1] && testSN.flagReady[1]) {
                tb_snHead2.Text = tb_sn.Text;
                tb_sn.Text = string.Empty;
                testSN.flagReady[1] = false;
                return;
            }

            if (testSN.flagTurbo[2] && testSN.flagReady[2]) {
                tb_snHead3.Text = tb_sn.Text;
                tb_sn.Text = string.Empty;
                testSN.flagReady[2] = false;
                return;
            }

            if (testSN.flagTurbo[3] && testSN.flagReady[3]) {
                tb_snHead4.Text = tb_sn.Text;
                tb_sn.Text = string.Empty;
                testSN.flagReady[3] = false;
                return;
            }

            MessageBox.Show(testSN.mesSage.snFull);
            tb_sn.Text = string.Empty;
        }
        private void RunBackGroundWorker() {
            //bgwk_image.RunWorkerAsync();
            //bgwk_dmm.RunWorkerAsync();

            Application.Idle += NormalMode;
        }
        private void GetEmployeeID() {
            user_id userID = new user_id();
            string id;

            userID.ShowDialog();
            id = File.ReadAllText(userID.pathFile);
            File.Delete(userID.pathFile);

            if (id == string.Empty) {
                this.Close();
                return;
            }

            tb_user.Text = id;
            TeamPrecision.PRISM.cSettingValues.EmployeeID = id;
            setupPay.write_text(prismTest.headConfig.employeeID, id, prismTest.nameFile);
        }
        private void GetLogConfig() {
            tb_setTestTime.Text = setupPay.read_text(logFile.headConfig.setTestTime, logFile.nameFile);
            logFile.testTime = Convert.ToInt32(tb_setTestTime.Text);

            logFile.showFail = Convert.ToBoolean(setupPay.read_text(logFile.headConfig.showFail, logFile.nameFile));
            logFile.logConfig.firmwareVersion = setupPay.read_text(logFile.headConfig.firmwareVersion, logFile.nameFile);
            logFile.logConfig.processName = setupPay.read_text(logFile.headConfig.processName, logFile.nameFile);
            logFile.logConfig.job = setupPay.read_text(logFile.headConfig.job, logFile.nameFile);
            logFile.logConfig.simapn = setupPay.read_text(logFile.headConfig.simapn, logFile.nameFile);
            logFile.pathLog = setupPay.read_text(logFile.headConfig.pathLog, logFile.nameFile);
        }
        private void GetSpec() {
            dmm.nameDMM = setupPay.read_text(dmm.spec.headConfig.nameDmm, dmm.spec.nameFile);
            dmm.relay.port = setupPay.read_text(dmm.spec.headConfig.comportRelay, dmm.spec.nameFile);
            dmm.spec.delayWaitExeCute = Convert.ToInt32(setupPay.read_text(dmm.spec.headConfig.delayWaitExeCute, dmm.spec.nameFile));
        }
        private void GetSpecMinMax() {
            dmm.spec.allFg = setupPay.read_text(dmm.spec.headConfig.batteryFg1, dmm.spec.nameFile);
            if (dmm.spec.allFg.Contains(prismTest.fg.Replace("FG-", "").Replace("-LF", ""))) {
                tb_min.Text = setupPay.read_text(dmm.spec.headConfig.batteryMin1, dmm.spec.nameFile);
                tb_max.Text = setupPay.read_text(dmm.spec.headConfig.batteryMax1, dmm.spec.nameFile);
                dmm.spec.min = Convert.ToDouble(tb_min.Text);
                dmm.spec.max = Convert.ToDouble(tb_max.Text);
                return;
            }

            dmm.spec.allFg = setupPay.read_text(dmm.spec.headConfig.batteryFg2, dmm.spec.nameFile);
            if (dmm.spec.allFg.Contains(prismTest.fg.Replace("FG-", "").Replace("-LF", ""))) {
                tb_min.Text = setupPay.read_text(dmm.spec.headConfig.batteryMin2, dmm.spec.nameFile);
                tb_max.Text = setupPay.read_text(dmm.spec.headConfig.batteryMax2, dmm.spec.nameFile);
                dmm.spec.min = Convert.ToDouble(tb_min.Text);
                dmm.spec.max = Convert.ToDouble(tb_max.Text);
                return;
            }

            dmm.spec.allFg = setupPay.read_text(dmm.spec.headConfig.batteryFg3, dmm.spec.nameFile);
            if (dmm.spec.allFg.Contains(prismTest.fg.Replace("FG-", "").Replace("-LF", ""))) {
                tb_min.Text = setupPay.read_text(dmm.spec.headConfig.batteryMin3, dmm.spec.nameFile);
                tb_max.Text = setupPay.read_text(dmm.spec.headConfig.batteryMax3, dmm.spec.nameFile);
                dmm.spec.min = Convert.ToDouble(tb_min.Text);
                dmm.spec.max = Convert.ToDouble(tb_max.Text);
                return;
            }

            dmm.spec.allFg = setupPay.read_text(dmm.spec.headConfig.batteryFg4, dmm.spec.nameFile);
            if (dmm.spec.allFg.Contains(prismTest.fg.Replace("FG-", "").Replace("-LF", ""))) {
                tb_min.Text = setupPay.read_text(dmm.spec.headConfig.batteryMin4, dmm.spec.nameFile);
                tb_max.Text = setupPay.read_text(dmm.spec.headConfig.batteryMax4, dmm.spec.nameFile);
                dmm.spec.min = Convert.ToDouble(tb_min.Text);
                dmm.spec.max = Convert.ToDouble(tb_max.Text);
                return;
            }

            MessageBox.Show("ไม่พบ Spec Min Max ใน FG ที่ระบุ\r\nโปรดตรวจสอบ Config File");
        }
        private void CreateFolderDataLog() {
            saveDataCsv.path = setupPay.read_text(saveDataCsv.headConfig.path, saveDataCsv.nameFile);

            if (!Directory.Exists(saveDataCsv.path)) {
                Directory.CreateDirectory(saveDataCsv.path);
            }
        }
        private void CreateFolderDataLog2() {
            string pathWo = saveDataCsv.path + "\\" + prismTest.fg + "\\" + prismTest.wo.Replace("/", "_");

            if (!Directory.Exists(pathWo)) {
                Directory.CreateDirectory(pathWo);
            }
        }
        private void FocusForm() {
            if (setUpPayFlag || !focusFormTick || !cb_focusForm.Checked) {
                return;
            }

            focusFormTick = false;
            this.Activate();
            if (Form.ActiveForm != this) {
                this.WindowState = FormWindowState.Minimized;
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void SetFormImage() {
            this.Hide();
            Thread.Sleep(100);
            imageClass.form = new Form();
            imageClass.form.Size = new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            imageClass.form.Location = new Point(0, 0);
            imageClass.form.ControlBox = false;
            imageClass.form.TransparencyKey = Color.Red;
            imageClass.form.BackColor = Color.Red;
            //image.form.Opacity = 0.01;
            imageClass.form.Cursor = Cursors.Cross;
            imageClass.graphics = imageClass.form.CreateGraphics();
            imageClass.pen = new Pen(Color.Blue);
            imageClass.pen.Width = 2;
        }
        private void SetEventImage_GetImage() {
            imageClass.form.MouseDown += new MouseEventHandler(MouseDown);
            imageClass.form.MouseMove += new MouseEventHandler(MouseMove);
            imageClass.form.MouseUp += new MouseEventHandler(MouseUp_GetImage);
            imageClass.form.Show();
        }
        private void SetEventImage_GetPosition() {
            imageClass.form.MouseDown += new MouseEventHandler(MouseDown);
            imageClass.form.MouseMove += new MouseEventHandler(MouseMove);
            imageClass.form.MouseUp += new MouseEventHandler(MouseUp_GetPosition);
            imageClass.form.Show();
        }
        private void SetEventImage_GetPositionClick() {
            imageClass.form.MouseDown += new MouseEventHandler(MouseDown_GetPosition);
            imageClass.form.Show();
        }
        private void GetImage() {
            using (Bitmap bitmap = new Bitmap(Convert.ToInt32(Math.Sqrt(Math.Pow(Convert.ToDouble(imageClass.startPoint.X -
                imageClass.stopPoint.X), 2))), Convert.ToInt32(Math.Sqrt(Math.Pow(Convert.ToDouble(imageClass.startPoint.Y -
                imageClass.stopPoint.Y), 2))))) {

                using (Graphics graphics = Graphics.FromImage(bitmap)) {

                    if (imageClass.startPoint.X > imageClass.stopPoint.X) {
                        imageClass.startPoint = new Point(imageClass.stopPoint.X, imageClass.startPoint.Y);
                    }

                    if (imageClass.startPoint.Y > imageClass.stopPoint.Y) {
                        imageClass.startPoint = new Point(imageClass.startPoint.X, imageClass.stopPoint.Y);
                    }

                    graphics.CopyFromScreen(imageClass.startPoint, Point.Empty, bitmap.Size);
                }

                imageClass.bitmapSup = new Bitmap((Image)bitmap);
                imageClass.SaveStatus(this, (Image)imageClass.bitmapSup);
            }
        }
        private void GetPosition() {
            double widthCal_1 = Convert.ToDouble(imageClass.startPoint.X - imageClass.stopPoint.X);
            double widthCal_2 = Math.Pow(widthCal_1, 2);
            double widthCal_3 = Math.Sqrt(widthCal_2);
            int width = Convert.ToInt32(widthCal_3);

            double heightCal_1 = Convert.ToDouble(imageClass.startPoint.Y - imageClass.stopPoint.Y);
            double heightCal_2 = Math.Pow(heightCal_1, 2);
            double heightCal_3 = Math.Sqrt(heightCal_2);
            int height = Convert.ToInt32(heightCal_3);

            if (imageClass.startPoint.X > imageClass.stopPoint.X) {
                imageClass.startPoint = new Point(imageClass.stopPoint.X, imageClass.startPoint.Y);
            }

            if (imageClass.startPoint.Y > imageClass.stopPoint.Y) {
                imageClass.startPoint = new Point(imageClass.startPoint.X, imageClass.stopPoint.Y);
            }

            AddPosition2(width, height);
        }
        private void AddPosition2(int width, int height) {
            if (imageClass.position.head == imageClass.position.head1) {
                position.number1.X = imageClass.startPoint.X;
                position.number1.Y = imageClass.startPoint.Y;
                position.number1.Width = width;
                position.number1.Height = height;
                SavePositionHead1();
            }

            if (imageClass.position.head == imageClass.position.head2) {
                position.number2.X = imageClass.startPoint.X;
                position.number2.Y = imageClass.startPoint.Y;
                position.number2.Width = width;
                position.number2.Height = height;
                SavePositionHead2();
            }

            if (imageClass.position.head == imageClass.position.head3) {
                position.number3.X = imageClass.startPoint.X;
                position.number3.Y = imageClass.startPoint.Y;
                position.number3.Width = width;
                position.number3.Height = height;
                SavePositionHead3();
            }

            if (imageClass.position.head == imageClass.position.head4) {
                position.number4.X = imageClass.startPoint.X;
                position.number4.Y = imageClass.startPoint.Y;
                position.number4.Width = width;
                position.number4.Height = height;
                SavePositionHead4();
            }

            if (imageClass.position.head == imageClass.position.head5) {
                position.number5.X = imageClass.startPoint.X;
                position.number5.Y = imageClass.startPoint.Y;
                position.number5.Width = width;
                position.number5.Height = height;
                SavePositionHead5();
            }
        }
        private void AddPosition() {
            if (imageClass.position.head == imageClass.position.head1) {
                position.number1.X = imageClass.startPoint.X;
                position.number1.Y = imageClass.startPoint.Y;
                SavePositionHead1();
            }

            if (imageClass.position.head == imageClass.position.head2) {
                position.number2.X = imageClass.startPoint.X;
                position.number2.Y = imageClass.startPoint.Y;
                SavePositionHead2();
            }

            if (imageClass.position.head == imageClass.position.head3) {
                position.number3.X = imageClass.startPoint.X;
                position.number3.Y = imageClass.startPoint.Y;
                SavePositionHead3();
            }

            if (imageClass.position.head == imageClass.position.head4) {
                position.number4.X = imageClass.startPoint.X;
                position.number4.Y = imageClass.startPoint.Y;
                SavePositionHead4();
            }

            if (imageClass.position.head == imageClass.position.head5) {
                position.number5.X = imageClass.startPoint.X;
                position.number5.Y = imageClass.startPoint.Y;
                SavePositionHead5();
            }

            if (imageClass.position.head == imageClass.position.execute) {
                position.execute.X = imageClass.startPoint.X;
                position.execute.Y = imageClass.startPoint.Y;
                SavePositionExcute();
            }
        }
        private void SavePositionHead1() {
            setupPay.write_text(position.headConfig.position1_X, position.number1.X.ToString(), position.nameFile);
            setupPay.write_text(position.headConfig.position1_Y, position.number1.Y.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position1_Width, position.number1.Width.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position1_Height, position.number1.Height.ToString(), position.nameFile);
        }
        private void SavePositionHead2() {
            setupPay.write_text(position.headConfig.position2_X, position.number2.X.ToString(), position.nameFile);
            setupPay.write_text(position.headConfig.position2_Y, position.number2.Y.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position2_Width, position.number2.Width.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position2_Height, position.number2.Height.ToString(), position.nameFile);
        }
        private void SavePositionHead3() {
            setupPay.write_text(position.headConfig.position3_X, position.number3.X.ToString(), position.nameFile);
            setupPay.write_text(position.headConfig.position3_Y, position.number3.Y.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position3_Width, position.number3.Width.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position3_Height, position.number3.Height.ToString(), position.nameFile);
        }
        private void SavePositionHead4() {
            setupPay.write_text(position.headConfig.position4_X, position.number4.X.ToString(), position.nameFile);
            setupPay.write_text(position.headConfig.position4_Y, position.number4.Y.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position4_Width, position.number4.Width.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position4_Height, position.number4.Height.ToString(), position.nameFile);
        }
        private void SavePositionHead5() {
            setupPay.write_text(position.headConfig.position5_X, position.number5.X.ToString(), position.nameFile);
            setupPay.write_text(position.headConfig.position5_Y, position.number5.Y.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position5_Width, position.number5.Width.ToString(), position.nameFile);
            //setupPay.write_text(position.headConfig.position5_Height, position.number5.Height.ToString(), position.nameFile);
        }
        private void SavePositionExcute() {
            setupPay.write_text(position.headConfig.positionExecute_X, position.execute.X.ToString(), position.nameFile);
            setupPay.write_text(position.headConfig.positionExecute_Y, position.execute.Y.ToString(), position.nameFile);
        }
        private void ShowPosition() {
            int delete = 8;

            imageClass.graphics.Clear(Color.Red);
            imageClass.graphics.DrawRectangle(imageClass.pen, position.number1.X - delete, position.number1.Y - delete,
                position.number1.Width, position.number1.Height);

            imageClass.graphics.DrawRectangle(imageClass.pen, position.number2.X - delete, position.number2.Y - delete,
                position.number2.Width, position.number2.Height);

            imageClass.graphics.DrawRectangle(imageClass.pen, position.number3.X - delete, position.number3.Y - delete,
                position.number3.Width, position.number3.Height);

            imageClass.graphics.DrawRectangle(imageClass.pen, position.number4.X - delete, position.number4.Y - delete,
                position.number4.Width, position.number4.Height);

            imageClass.graphics.DrawRectangle(imageClass.pen, position.number5.X - delete, position.number5.Y - delete,
                position.number5.Width, position.number5.Height);
        }

        private void MouseDown(object sender, MouseEventArgs e) {
            imageClass.startPoint = new Point(Control.MousePosition.X, Control.MousePosition.Y);
            imageClass.penFlag = true;
        }
        private void MouseMove(object sender, MouseEventArgs e) {
            if (!imageClass.penFlag) return;

            int pointMinX = ((imageClass.startPoint.X + Control.MousePosition.X) -
                Convert.ToInt32(Math.Sqrt(Math.Pow(Convert.ToDouble(imageClass.startPoint.X -
                Control.MousePosition.X), 2)))) / 2;

            int pointMinY = ((imageClass.startPoint.Y + Control.MousePosition.Y) -
                Convert.ToInt32(Math.Sqrt(Math.Pow(Convert.ToDouble(imageClass.startPoint.Y -
                Control.MousePosition.Y), 2)))) / 2;

            imageClass.graphics.Clear(Color.Red);
            imageClass.graphics.DrawRectangle(imageClass.pen, pointMinX - 10, pointMinY - 10,
                Convert.ToInt32(Math.Sqrt(Math.Pow(Convert.ToDouble(Control.MousePosition.X -
                imageClass.startPoint.X), 2))) + 4,
                Convert.ToInt32(Math.Sqrt(Math.Pow(Convert.ToDouble(Control.MousePosition.Y -
                imageClass.startPoint.Y), 2)) + 4));
        }
        private void MouseUp_GetImage(object sender, MouseEventArgs e) {
            imageClass.stopPoint = new Point(Control.MousePosition.X, Control.MousePosition.Y);

            GetImage();

            imageClass.form.Close();
            Thread.Sleep(100);
            this.Show();

            imageClass.penFlag = false;
        }
        private void MouseUp_GetPosition(object sender, MouseEventArgs e) {
            imageClass.stopPoint = new Point(Control.MousePosition.X, Control.MousePosition.Y);

            GetPosition();

            imageClass.form.Close();
            Thread.Sleep(100);
            this.Show();

            imageClass.penFlag = false;
        }
        private void MouseDown_GetPosition(object sender, MouseEventArgs e) {
            imageClass.startPoint = new Point(Control.MousePosition.X, Control.MousePosition.Y);

            AddPosition();

            imageClass.form.Close();
            Thread.Sleep(100);
            this.Show();
        }

        private void ContainsImage(TextBox sn, TextBox volt, PictureBox pictureBox, Label label, int head) {
            Image<Bgr, byte> source = new Image<Bgr, Byte>(new Bitmap(pictureBox.Image)); //max
            Image<Bgr, byte> reference = new Image<Bgr, byte>(imageClass.define.pathImageRefer +
                imageClass.define.lastNamePNG); //min
            //Image<Bgr, byte> imageShow = source.Copy();

            using (Image<Gray, float> result = source.MatchTemplate(reference, imageClass.typeContain)) {

                double[] minValues, maxValues;
                Point[] minLocations, maxLocations;
                result.MinMax(out minValues, out maxValues, out minLocations, out maxLocations);

                if (maxValues[0] > imageClass.define.ContainsValue) {
                    //Rectangle match = new Rectangle(maxLocations[0], imageShow.Size);
                    //imageShow.Draw(match, new Bgr(Color.Red), 3);
                    testSN.ClearTimer(head);
                    label.Text = imageClass.open;
                    label.BackColor = Color.Lime;
                    sn.Visible = true;
                    volt.Visible = true;
                    testSN.flagTurbo[head - 1] = true;

                } else {
                    if (!testSN.CheckTimeOut(head)) {
                        return;
                    }
                    label.Text = imageClass.close;
                    label.BackColor = Color.Red;
                    sn.Text = string.Empty;
                    sn.BackColor = Color.White;
                    sn.Visible = false;
                    volt.Text = string.Empty;
                    volt.BackColor = Color.White;
                    volt.Visible = false;
                    testSN.flagTurbo[head - 1] = false;
                    testSN.flagReady[head - 1] = true;
                    SetFinalLabelVisable(head);
                }
            }
        }
        private void ContainsImage2(TextBox sn, TextBox volt, Image image, Label label, int head) {
            Bitmap source = (Bitmap)image.Clone();
            Bitmap reference = (Bitmap)pb_imageRefer.Image.Clone();
            int pixel = 1;
            bool flag_compar = false;

            for (int i_shift = 0; i_shift < source.Width - reference.Width; i_shift++) {
                for (int j_shift = 0; j_shift < source.Height - reference.Height; j_shift++) {
                    flag_compar = true;
                    for (int i = 0; i < reference.Width; i += pixel) {
                        for (int j = 0; j < reference.Height; j += pixel) {
                            if (reference.GetPixel(i, j) != source.GetPixel(i + i_shift, j + j_shift)) {
                                i = reference.Width;
                                flag_compar = false;
                                break;
                            }
                        }
                    }
                    if (flag_compar) {
                        break;
                    }
                }
                if (flag_compar) {
                    break;
                }
            }

            if (flag_compar) {
                testSN.ClearTimer(head);
                label.Text = imageClass.open;
                label.BackColor = Color.Lime;
                sn.Visible = true;
                volt.Visible = true;
                testSN.flagTurbo[head - 1] = true;
            } else {

                if (!testSN.CheckTimeOut(head)) {
                    return;
                }
                dmm.statusRunning[head - 1] = false;
                DelaymS(50);
                label.Text = imageClass.close;
                label.BackColor = Color.Red;
                sn.Text = string.Empty;
                sn.BackColor = Color.White;
                sn.Visible = false;
                volt.Text = string.Empty;
                volt.BackColor = Color.White;
                volt.Visible = false;
                testSN.flagTurbo[head - 1] = false;
                testSN.flagReady[head - 1] = true;
                SetFinalLabelVisable(head);
            }
        }
        private void GetImageByPosition(PictureBox pictureBox, int X, int Y, int width, int height) {
            //Screen.PrimaryScreen.Bounds.Width
            //Screen.PrimaryScreen.Bounds.Height
            Bitmap bitmap;

            try {
                bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            } catch {
                return;
            }

            Graphics graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(X, Y, 0, 0, bitmap.Size);

            pictureBox.Image = bitmap;

            MemoryLeakSolve();
        }
        private void MemoryLeakSolve() {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        private void ShowImageStatus() {
            GetImageByPosition(pb_status1, position.number1.X - imageClass.scrollBar, position.number1.Y -
                (position.number1.Height / 2), position.number1.Width, position.number1.Height);

            GetImageByPosition(pb_status2, position.number2.X - imageClass.scrollBar, position.number2.Y -
                (position.number2.Height / 2), position.number2.Width, position.number2.Height);

            GetImageByPosition(pb_status3, position.number3.X - imageClass.scrollBar, position.number3.Y -
                (position.number3.Height / 2), position.number3.Width, position.number3.Height);

            GetImageByPosition(pb_status4, position.number4.X - imageClass.scrollBar, position.number4.Y -
                (position.number4.Height / 2), position.number4.Width, position.number4.Height);
        }
        private void ShowImageStatus_DoWork() {
            imageClass.flagDoWork = true;
            bgwk_image.ReportProgress(0);

            while (imageClass.flagDoWork) {
                Thread.Sleep(50);
            }
        }
        private void ContainsImage_DoWork() {
            imageClass.flagDoWork = true;
            bgwk_image.ReportProgress(1);

            while (imageClass.flagDoWork) {
                Thread.Sleep(50);
            }
        }

        private void LogConnectRelay(bool status) {
            if (status) {
                bgwk_dmm.ReportProgress(0);
            } else {

                bgwk_dmm.ReportProgress(1);
            }
        }
        private void SetColorLabel(Label label, Color color) {
            label.ForeColor = color;
        }
        private void ReadVolt_DoWork() {
            dmm.flagDoWork = true;
            bgwk_dmm.ReportProgress(2);

            while (dmm.flagDoWork) {
                Thread.Sleep(50);
            }
        }
        private void FinePortDmm_DoWork() {
            dmm.flagDoWork = true;
            bgwk_dmm.ReportProgress(3);

            while (dmm.flagDoWork) {
                Thread.Sleep(50);
            }
        }
        private void DisPlayFocusSN_DoWork() {
            dmm.flagDoWork = true;
            bgwk_dmm.ReportProgress(4);

            while (dmm.flagDoWork) {
                Thread.Sleep(50);
            }
        }
        private void DisPlayFocusSN() {
            gb_head1.BackColor = SystemColors.Control;
            gb_head2.BackColor = SystemColors.Control;
            gb_head3.BackColor = SystemColors.Control;
            gb_head4.BackColor = SystemColors.Control;

            if (testSN.flagTurbo[0] && testSN.flagReady[0]) {
                gb_head1.BackColor = Color.Aqua;
                if (!focus.newUnit1) {
                    focus.newUnit1 = true;
                    tb_sn.Focus();
                }
                return;
            }

            if (testSN.flagTurbo[1] && testSN.flagReady[1]) {
                gb_head2.BackColor = Color.Aqua;
                if (!focus.newUnit2) {
                    focus.newUnit2 = true;
                    tb_sn.Focus();
                }
                return;
            }

            if (testSN.flagTurbo[2] && testSN.flagReady[2]) {
                gb_head3.BackColor = Color.Aqua;
                if (!focus.newUnit3) {
                    focus.newUnit3 = true;
                    tb_sn.Focus();
                }
                return;
            }

            if (testSN.flagTurbo[3] && testSN.flagReady[3]) {
                gb_head4.BackColor = Color.Aqua;
                if (!focus.newUnit4) {
                    focus.newUnit4 = true;
                    tb_sn.Focus();
                }
                return;
            }

            if (tb_snHead1.Text.Length != 0 || tb_snHead2.Text.Length != 0 || tb_snHead3.Text.Length != 0 || tb_snHead4.Text.Length != 0) {
                bt_exeCute.Enabled = true;
                bt_exeCute.Focus();
                focus.Clear();
                return;
            }

            if (tb_customer.Text.Length != 0) {
                tb_sn.Focus();
                return;
            }
        }

        private void SnToTurBo() {
            if (tb_snHead1.Text.Length == testSN.length) {
                autoit.MouseClick("LEFT", position.number1.X, position.number1.Y, 1, -1);
                Thread.Sleep(100);
                autoit.Send(tb_snHead1.Text);
                Thread.Sleep(100);
            }

            if (tb_snHead2.Text.Length == testSN.length) {
                autoit.MouseClick("LEFT", position.number2.X, position.number2.Y, 1, -1);
                Thread.Sleep(100);
                autoit.Send(tb_snHead2.Text);
                Thread.Sleep(100);
            }

            if (tb_snHead3.Text.Length == testSN.length) {
                autoit.MouseClick("LEFT", position.number3.X, position.number3.Y, 1, -1);
                Thread.Sleep(100);
                autoit.Send(tb_snHead3.Text);
                Thread.Sleep(100);
            }

            if (tb_snHead4.Text.Length == testSN.length) {
                autoit.MouseClick("LEFT", position.number4.X, position.number4.Y, 1, -1);
                Thread.Sleep(100);
                autoit.Send(tb_snHead4.Text);
                Thread.Sleep(100);
            }
        }
        private void ClickToExeCute() {
            autoit.MouseClick("LEFT", position.execute.X, position.execute.Y, 1, -1);
        }
        private void SetFlagLogFile() {
            pgb_test.Value = 0;
            pgb_test.Maximum = 0;

            if (tb_snHead1.Visible) {
                logFile.head1 = false;
                saveDataCsv.head1 = false;
                dmm.statusRunning[0] = true;
                tb_snHead1.BackColor = Color.White;
                tb_voltHead1.BackColor = Color.White;
                lb_finalStatus1.Visible = false;
                SetFinalLabelWait(1);
                pgb_test.Maximum += 2;
            }
            if (tb_snHead2.Visible) {
                logFile.head2 = false;
                saveDataCsv.head2 = false;
                dmm.statusRunning[1] = true;
                tb_snHead2.BackColor = Color.White;
                tb_voltHead2.BackColor = Color.White;
                lb_finalStatus2.Visible = false;
                SetFinalLabelWait(2);
                pgb_test.Maximum += 2;
            }
            if (tb_snHead3.Visible) {
                logFile.head3 = false;
                saveDataCsv.head3 = false;
                dmm.statusRunning[2] = true;
                tb_snHead3.BackColor = Color.White;
                tb_voltHead3.BackColor = Color.White;
                lb_finalStatus3.Visible = false;
                SetFinalLabelWait(3);
                pgb_test.Maximum += 2;
            }
            if (tb_snHead4.Visible) {
                logFile.head4 = false;
                saveDataCsv.head4 = false;
                dmm.statusRunning[3] = true;
                tb_snHead4.BackColor = Color.White;
                tb_voltHead4.BackColor = Color.White;
                lb_finalStatus4.Visible = false;
                SetFinalLabelWait(4);
                pgb_test.Maximum += 2;
            }
            logFile.readLogFull = false;
            prismTest.flagWaitUpPrism = true;
            imageClass.flagCheck = false;
            dmm.flagRead = true;
        }
        private bool CheckSnBeforeClick() {
            if (tb_snHead1.Visible) {
                if (tb_snHead1.Text.Length != testSN.length) {
                    return false;
                }
            }
            if (tb_snHead2.Visible) {
                if (tb_snHead2.Text.Length != testSN.length) {
                    return false;
                }
            }
            if (tb_snHead3.Visible) {
                if (tb_snHead3.Text.Length != testSN.length) {
                    return false;
                }
            }
            if (tb_snHead4.Visible) {
                if (tb_snHead4.Text.Length != testSN.length) {
                    return false;
                }
            }

            return true;
        }

        private void ReadXmlFileTimeOut_DoWork() {
            imageClass.flagDoWork = true;
            bgwk_image.ReportProgress(5);

            while (imageClass.flagDoWork) {
                DelaymS(50);
            }
        }
        private void ReadXmlFileTimeOut() {
            if (!logFile.stopWatch.IsRunning) {
                logFile.stopWatch.Restart();

            } else {
                if (logFile.stopWatch.ElapsedMilliseconds > logFile.testTime) {
                    imageClass.flagCheck = true;
                    logFile.stopWatch.Stop();

                    if (!logFile.head1) {
                        tb_snHead1.BackColor = Color.Red;
                        SetFinalLabelFail(1);
                    }
                    if (!logFile.head2) {
                        tb_snHead2.BackColor = Color.Red;
                        SetFinalLabelFail(2);
                    }
                    if (!logFile.head3) {
                        tb_snHead3.BackColor = Color.Red;
                        SetFinalLabelFail(3);
                    }
                    if (!logFile.head4) {
                        tb_snHead4.BackColor = Color.Red;
                        SetFinalLabelFail(4);
                    }
                }
            }
        }
        private void ReadXmlFile_DoWork() {
            imageClass.flagDoWork = true;
            bgwk_image.ReportProgress(2);

            while (imageClass.flagDoWork) {
                DelaymS(50);
            }
        }
        private void ReadXmlFile() {
            bool flag = false;

            if (!logFile.head1) {
                flag = true;
                Log(LogMsgType.Normal_Black, "\nRead Xml file head 1");
                if (logFile.ReadXML(tb_snHead1.Text)) {
                    logFile.head1 = true;
                    CheckXmlFile(tb_snHead1, prismTest.upPrismFct8_1);
                }
            }
            if (!logFile.head2) {
                flag = true;
                Log(LogMsgType.Normal_Black, "\nRead Xml file head 2");
                if (logFile.ReadXML(tb_snHead2.Text)) {
                    logFile.head2 = true;
                    CheckXmlFile(tb_snHead2, prismTest.upPrismFct8_2);
                }
            }
            if (!logFile.head3) {
                flag = true;
                Log(LogMsgType.Normal_Black, "\nRead Xml file head 3");
                if (logFile.ReadXML(tb_snHead3.Text)) {
                    logFile.head3 = true;
                    CheckXmlFile(tb_snHead3, prismTest.upPrismFct8_3);
                }
            }
            if (!logFile.head4) {
                flag = true;
                Log(LogMsgType.Normal_Black, "\nRead Xml file head 4");
                if (logFile.ReadXML(tb_snHead4.Text)) {
                    logFile.head4 = true;
                    CheckXmlFile(tb_snHead4, prismTest.upPrismFct8_4);
                }
            }

            if (!flag) {
                logFile.readLogFull = true;
            }
        }
        private void CheckXmlFile(TextBox sn, PrismTest.UpPrismFct8 jsonPrism) {
            jsonPrism.SerialNumber = logFile.serialNumber;
            if (!logFile.result) {
                sn.BackColor = Color.Red;
                if (logFile.showFail) {
                    MessageBox.Show("TurBo Fail!!");
                }
                return;
            }

            jsonPrism.FirmwareVersion = logFile.firmwareVersion;
            if (logFile.firmwareVersion != logFile.logConfig.firmwareVersion) {
                sn.BackColor = Color.Red;
                if (logFile.showFail) {
                    MessageBox.Show("Firmware Version Fail!!");
                }
                return;
            }

            jsonPrism.ProcessName = logFile.processName;
            if (logFile.processName != logFile.logConfig.processName) {
                sn.BackColor = Color.Red;
                if (logFile.showFail) {
                    MessageBox.Show("Process Name Fail!!");
                }
                return;
            }

            jsonPrism.Job = logFile.job;
            if (logFile.job != logFile.logConfig.job) {
                sn.BackColor = Color.Red;
                if (logFile.showFail) {
                    MessageBox.Show("Job Fail!!");
                }
                return;
            }

            jsonPrism.SIMAPN = logFile.simapn;
            if (logFile.simapn != logFile.logConfig.simapn) {
                sn.BackColor = Color.Red;
                if (logFile.showFail) {
                    MessageBox.Show("SIMAPN Fail!!");
                }
                return;
            }

            if (payTest) {
                sn.BackColor = Color.Lime;
                return;

            } else {
                string[] status = TeamPrecision.PRISM.cSNs.CheckStatusSNv2(sn.Text, tb_wo.Text);

                if (status[1].Contains("ยังไม่ผ่าน")) {
                    MessageBox.Show("ยังไม่ผ่าน FCT7");
                    sn.BackColor = Color.Red;
                    return;
                }

                if (status[1].Contains("FCT7 สถานะ : FAIL")) {
                    MessageBox.Show(status[1]);
                    sn.BackColor = Color.Red;
                    return;
                }

                DataTable dataTable = TeamPrecision.PRISM.cSNs.ChecKSerialTeam(sn.Text);
                string serialASM = dataTable.Rows[0].ItemArray[0].ToString();
                DataTable dataPrismSup = TeamPrecision.PRISM.cSNs.getDataTest(serialASM, "FCT7");

                if (dataPrismSup.Rows.Count == 0) {
                    sn.BackColor = Color.Red;
                    if (logFile.showFail) {
                        MessageBox.Show("data FCT7 Fail!!");
                    }
                    return;
                }

                string[] deviceIdSup = new string[1];
                for (int loop = 0; loop < dataPrismSup.Rows.Count; loop++) {

                    if (dataPrismSup.Rows[loop]["PROC_NAME"].ToString() == "FCT7") {
                        string dataPrism = dataPrismSup.Rows[loop].ItemArray[3].ToString();
                        deviceIdSup = dataPrism.Replace("Value: ", "ฅ").Replace(" <br>", "ฅ").Replace(",", "ฅ").Split('ฅ');
                        break;
                    }
                }

                jsonPrism.SerialASM = serialASM;
                jsonPrism.ICCID = deviceIdSup[3];
                jsonPrism.DeviceID = logFile.deviceID;
                if (logFile.deviceID != deviceIdSup[1]) {
                    sn.BackColor = Color.Red;
                    MessageBox.Show("Prism Fail!!\r\nหมายเลข Device ID ใน Log File ไม่ตรงกับ หมายเลขใน Prism\r\n" +
                        "ให้แจ้งทาง TE ตรวจสอบ");
                    return;
                }

                sn.BackColor = Color.Lime;
            }
        }
        private void WaitUpPrism_DoWork() {
            imageClass.flagDoWork = true;
            bgwk_image.ReportProgress(3);

            while (imageClass.flagDoWork) {
                Thread.Sleep(50);
            }
        }
        private bool WaitUpPrism(TextBox sn, TextBox volt, PrismTest.UpPrismFct8 jsonPrism, int head) {
            string resultPrism = prismTest.pass;

            if (!sn.Visible && !volt.Visible) {
                return true;
            }

            if (tester.debugTestBattery) {
                jsonPrism.Volt = volt.Text;
                jsonPrism.SerialNumber = sn.Text;

                if (volt.BackColor == Color.Red) {
                    jsonPrism.Status = prismTest.fail;
                    SetFinalLabelFail(head);
                } else {
                    jsonPrism.Status = prismTest.pass;
                    SetFinalLabelPass(head);
                }
                SaveData(jsonPrism.Status, jsonPrism, head);
                return true;
            }

            if (sn.BackColor == Color.White || volt.BackColor == Color.White) {
                return false;
            }

            jsonPrism.Volt = volt.Text;

            if (sn.BackColor == Color.Red || volt.BackColor == Color.Red) {
                if (!prismTest.upFailToPrism) {
                    jsonPrism.Status = prismTest.fail;
                    SaveData(prismTest.fail, jsonPrism, head);
                    SetFinalLabelFail(head);
                    return true;
                }

                SaveData(prismTest.fail, jsonPrism, head);
                SetFinalLabelFail(head);
                resultPrism = prismTest.fail;

            } else {
                jsonPrism.Status = prismTest.pass;
                SaveData(resultPrism, jsonPrism, head);
            }

            if (prismTest.mode != PrismTest.Define.debug) {
                if (Convert.ToBoolean(prismTest.xml.checkProcessBefore)) {
                    string[] status = TeamPrecision.PRISM.cSNs.CheckStatusSNv2(sn.Text, tb_wo.Text);

                    if (status[1].Contains(prismTest.xml.processBefore)) {
                        sn.BackColor = Color.Red;
                        SetFinalLabelFail(head);
                        if (logFile.showFail) {
                            MessageBox.Show("Prism Fail!!\r\n" + status[0] + "\r\n" + status[1]);
                        }
                        return true;
                    }
                }

                string dataUpPrism = new JavaScriptSerializer().Serialize(jsonPrism);
                if (TeamPrecision.PRISM.cResults.SaveTestResult(sn.Text, resultPrism, dataUpPrism) != prismTest.success) {
                    sn.BackColor = Color.Red;
                    SetFinalLabelFail(head);
                    if (logFile.showFail) {
                        MessageBox.Show("Prism Fail!!\r\nอัพขึ้นระบบไม่ได้");
                    }
                    return true;
                }

                //string note = logFile.logFileNameNotePrism;
                string note = sn.Text + "-" + jsonPrism.DeviceID + "-Denali";
                if (TeamPrecision.PRISM.cSNs.updateTransExtSN(jsonPrism.SerialASM, note, prismTest.xml.employeeID,
                    prismTest.xml.computerName, prismTest.xml.processName) != prismTest.success) {
                    //sn.BackColor = Color.Red;
                    //SetFinalLabelFail(head);
                    //if (logFile.showFail) {
                    //    MessageBox.Show("Prism Note Fail!!\r\nอัพขึ้นระบบไม่ได้");
                    //}
                    //return true;
                }

                function_timeout(prismTest.getWO, 5000);
                tb_orderQty.Text = Convert.ToDouble(prismTest.orderQty).ToString();
                tb_outputQty.Text = Convert.ToDouble(prismTest.outputQty).ToString();
            }

            if (sn.BackColor != Color.Red && volt.BackColor != Color.Red) {
                SetFinalLabelPass(head);
                SetFinalLabelPass(head);
            }
            return true;
        }
        private void SaveData(string result, PrismTest.UpPrismFct8 jsonPrism, int head) {
            string csvPath = saveDataCsv.path + "\\" + prismTest.fg + "\\" + prismTest.wo.Replace("/", "_") + "\\" + prismTest.pass + ".csv";

            if (result == prismTest.fail) {
                csvPath = csvPath.Replace(prismTest.pass, prismTest.fail);
            }
            CreateHeadLog(csvPath);

            string date = DateTime.Now.ToString("yyyy/MM/dd", CultureInfo.CreateSpecificCulture("en-US"));
            string time = DateTime.Now.ToString("HH:mm:ss", CultureInfo.CreateSpecificCulture("en-US"));
            string dataResult = jsonPrism.SerialNumber + ",";
            dataResult += date + ",";
            dataResult += time + ",";
            dataResult += head + ",";
            dataResult += jsonPrism.Volt + ",";
            dataResult += jsonPrism.SerialASM + ",";
            dataResult += "'" + jsonPrism.DeviceID + ",";
            dataResult += "'" + jsonPrism.ICCID + ",";
            dataResult += jsonPrism.FirmwareVersion + ",";
            dataResult += jsonPrism.ProcessName + ",";
            dataResult += jsonPrism.Job + ",";
            dataResult += jsonPrism.SIMAPN + ",";
            dataResult += jsonPrism.Status;

            StreamWriter swOut = new StreamWriter(csvPath, true);
            while (true) {
                try {
                    swOut.WriteLine(dataResult);
                } catch {
                    MessageBox.Show("_กรุณาปิด log file csv ก่อน");
                    continue;
                }
                break;
            }
            swOut.Close();
        }
        private void CreateHeadLog(string fileName) {
            string header = "SerialNumber,Date,Time,Head,Volt,SerialASM,DeviceID,ICCID,FirmwareVersion," +
            "ProcessName,Job,SIMAPN,Status";

            bool SameHeaderOrNot = CheckHeaderLog(header, fileName);

            if ((!File.Exists(fileName)) | (SameHeaderOrNot == false)) {
                StreamWriter swOut = new StreamWriter(fileName, true);
                swOut.WriteLine(header);
                swOut.Close();
            }
        }
        private bool CheckHeaderLog(string CurrentHeader, string fileName) {
            string PreviousHeader = null;
            string[] CurrentHeader_split = CurrentHeader.Split(',');

            const Int32 BufferSize = 500;
            try {
                using (var fileStream = File.OpenRead(fileName))
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize)) {
                    String line;
                    while ((line = streamReader.ReadLine()) != null) {
                        List<string> names = new List<string>(line.Split(','));
                        if (names[0] == "SerialNumber" || names[0] == CurrentHeader_split[0]) {
                            PreviousHeader = line;
                        }
                    }
                    fileStream.Close();
                    if (string.Compare(CurrentHeader, PreviousHeader) != 0) {
                        return false;
                    }
                }
            } catch { }

            return true;
        }
        private void SetFinalLabelPass(int head) {
            Label label = new Label();

            switch (head) {
                case 1:
                    label = lb_finalStatus1;
                    break;
                case 2:
                    label = lb_finalStatus2;
                    break;
                case 3:
                    label = lb_finalStatus3;
                    break;
                case 4:
                    label = lb_finalStatus4;
                    break;
            }

            label.Text = prismTest.pass;
            label.BackColor = Color.Lime;
            label.Visible = true;
        }
        private void SetFinalLabelFail(int head) {
            Label label = new Label();

            switch (head) {
                case 1:
                    label = lb_finalStatus1;
                    break;
                case 2:
                    label = lb_finalStatus2;
                    break;
                case 3:
                    label = lb_finalStatus3;
                    break;
                case 4:
                    label = lb_finalStatus4;
                    break;
            }

            label.Text = prismTest.fail;
            label.BackColor = Color.Red;
            label.Visible = true;
        }
        private void SetFinalLabelWait(int head) {
            Label label = new Label();

            switch (head) {
                case 1:
                    label = lb_finalStatus1;
                    break;
                case 2:
                    label = lb_finalStatus2;
                    break;
                case 3:
                    label = lb_finalStatus3;
                    break;
                case 4:
                    label = lb_finalStatus4;
                    break;
            }

            label.Text = prismTest.wait;
            label.BackColor = Color.Blue;
            label.Visible = true;
        }
        private void SetFinalLabelVisable(int head) {
            Label label = new Label();

            switch (head) {
                case 1:
                    label = lb_finalStatus1;
                    break;
                case 2:
                    label = lb_finalStatus2;
                    break;
                case 3:
                    label = lb_finalStatus3;
                    break;
                case 4:
                    label = lb_finalStatus4;
                    break;
            }

            label.Visible = false;
        }
        #endregion

        #region================================================== Pocess ===================================================
        private bool startFlag = false;
        private void NormalMode(object sender, EventArgs e) {
            if (!startFlag) {
                startFlag = true;
                if (dmm.relay.Connect(this)) {
                    Log(LogMsgType.Outgoing_Green, dmm.relay.mesSage.connected);

                } else {
                    Log(LogMsgType.Error_Red, dmm.relay.mesSage.connectErr);
                    SetColorLabel(lb_volt1, Color.Red);
                    SetColorLabel(lb_volt2, Color.Red);
                    SetColorLabel(lb_volt3, Color.Red);
                    SetColorLabel(lb_volt4, Color.Red);
                    SetColorLabel(lb_vHead1, Color.Red);
                    SetColorLabel(lb_vHead2, Color.Red);
                    SetColorLabel(lb_vHead3, Color.Red);
                    SetColorLabel(lb_vHead4, Color.Red);
                }
                dmm.FinePortDmm();
                tb_wo.Focus();
                return;
            }

            if (imageClass.flagCheck) {
                ShowImageStatus();
                ContainsImage(tb_snHead1, tb_voltHead1, pb_status1, lb_status1, 1);
                ContainsImage(tb_snHead2, tb_voltHead2, pb_status2, lb_status2, 2);
                ContainsImage(tb_snHead3, tb_voltHead3, pb_status3, lb_status3, 3);
                ContainsImage(tb_snHead4, tb_voltHead4, pb_status4, lb_status4, 4);
                DisPlayFocusSN();
                FocusForm();

            } else {
                if (dmm.flagRead) {
                    dmm.flagRead = false;
                    if (tb_voltHead1.Visible) {
                        dmm.ReadVolt(this, tb_voltHead1, 1);
                        pgb_test.Value += 1;
                    }
                    if (tb_voltHead2.Visible) {
                        dmm.ReadVolt(this, tb_voltHead2, 2);
                        pgb_test.Value += 1;
                    }
                    if (tb_voltHead3.Visible) {
                        dmm.ReadVolt(this, tb_voltHead3, 3);
                        pgb_test.Value += 1;
                    }
                    if (tb_voltHead4.Visible) {
                        dmm.ReadVolt(this, tb_voltHead4, 4);
                        pgb_test.Value += 1;
                    }
                }
                if (!logFile.readLogFull) {
                    if (tester.debugTestBattery) {
                        logFile.readLogFull = true;

                    } else {
                        ReadXmlFile();
                    }
                    
                } else {
                    bool flag = true;
                    if (!saveDataCsv.head1) {
                        if (!WaitUpPrism(tb_snHead1, tb_voltHead1, prismTest.upPrismFct8_1, 1)) {
                            flag = false;
                        } else {
                            saveDataCsv.head1 = true;
                            if (tb_snHead1.Visible) {
                                pgb_test.Value += 1;
                            }
                        }
                    }
                    if (!saveDataCsv.head2) {
                        if (!WaitUpPrism(tb_snHead2, tb_voltHead2, prismTest.upPrismFct8_2, 2)) {
                            flag = false;
                        } else {
                            saveDataCsv.head2 = true;
                            if (tb_snHead2.Visible) {
                                pgb_test.Value += 1;
                            }
                        }
                    }
                    if (!saveDataCsv.head3) {
                        if (!WaitUpPrism(tb_snHead3, tb_voltHead3, prismTest.upPrismFct8_3, 3)) {
                            flag = false;
                        } else {
                            saveDataCsv.head3 = true;
                            if (tb_snHead3.Visible) {
                                pgb_test.Value += 1;
                            }
                        }
                    }
                    if (!saveDataCsv.head4) {
                        if (!WaitUpPrism(tb_snHead4, tb_voltHead4, prismTest.upPrismFct8_4, 4)) {
                            flag = false;
                        } else {
                            saveDataCsv.head4 = true;
                            if (tb_snHead4.Visible) {
                                pgb_test.Value += 1;
                            }
                        }
                    }
                    if (flag) {
                        prismTest.flagWaitUpPrism = false;
                        imageClass.flagCheck = true;
                        tm_timeOutLog.Enabled = false;
                        bt_exeCute.Text = "EXECUTE";
                        logFile.stopWatch.Stop();
                    }
                }
                

            }
            
            Thread.Sleep(100);
        }

        private void bgwk_dmm_DoWork(object sender, DoWorkEventArgs e) {
            LogConnectRelay(dmm.relay.Connect(this));
            FinePortDmm_DoWork();

            while (true) {
                if (dmm.flagRead) {
                    dmm.flagRead = false;
                    ReadVolt_DoWork();
                }

                DisPlayFocusSN_DoWork();
                Thread.Sleep(150);
            }
        }
        private void bgwk_dmm_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            if (e.ProgressPercentage == 0) {
                Log(LogMsgType.Outgoing_Green, dmm.relay.mesSage.connected);
            }

            if (e.ProgressPercentage == 1) {
                Log(LogMsgType.Error_Red, dmm.relay.mesSage.connectErr);
                SetColorLabel(lb_volt1, Color.Red);
                SetColorLabel(lb_volt2, Color.Red);
                SetColorLabel(lb_volt3, Color.Red);
                SetColorLabel(lb_volt4, Color.Red);
                SetColorLabel(lb_vHead1, Color.Red);
                SetColorLabel(lb_vHead2, Color.Red);
                SetColorLabel(lb_vHead3, Color.Red);
                SetColorLabel(lb_vHead4, Color.Red);
            }

            if (e.ProgressPercentage == 2) {
                if (tb_voltHead1.Visible) {
                    dmm.ReadVolt(this, tb_voltHead1, 1);
                }
                if (tb_voltHead2.Visible) {
                    dmm.ReadVolt(this, tb_voltHead2, 2);
                }
                if (tb_voltHead3.Visible) {
                    dmm.ReadVolt(this, tb_voltHead3, 3);
                }
                if (tb_voltHead4.Visible) {
                    dmm.ReadVolt(this, tb_voltHead4, 4);
                }
            }

            if (e.ProgressPercentage == 3) {
                dmm.FinePortDmm();
                tb_wo.Focus();
            }

            if (e.ProgressPercentage == 4) {
                DisPlayFocusSN();
            }

            dmm.flagDoWork = false;
        }

        private void bgwk_image_DoWork(object sender, DoWorkEventArgs e) {
            while (true) {

                if (imageClass.flagCheck) {
                    ShowImageStatus_DoWork();
                    Log(LogMsgType.Normal_Black, "\nCompare image check");
                    ContainsImage_DoWork();
                    goto JumpImage;
                }

                if (!logFile.readLogFull) {
                    ReadXmlFileTimeOut_DoWork();
                    ReadXmlFile_DoWork();
                }

                if (prismTest.flagWaitUpPrism) {
                    Log(LogMsgType.Normal_Black, "\nWait Up Prism");
                    WaitUpPrism_DoWork();
                }

                JumpImage:
                Thread.Sleep(50);
            }
        }
        private void bgwk_image_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            if (e.ProgressPercentage == 0) {
                ShowImageStatus();
            }

            if (e.ProgressPercentage == 1) {
                //ContainsImage2(tb_snHead1, tb_voltHead1, pb_status1.Image, lb_status1, 1);
                //ContainsImage2(tb_snHead2, tb_voltHead2, pb_status2.Image, lb_status2, 2);
                //ContainsImage2(tb_snHead3, tb_voltHead3, pb_status3.Image, lb_status3, 3);
                //ContainsImage2(tb_snHead4, tb_voltHead4, pb_status4.Image, lb_status4, 4);

                ContainsImage(tb_snHead1, tb_voltHead1, pb_status1, lb_status1, 1);
                ContainsImage(tb_snHead2, tb_voltHead2, pb_status2, lb_status2, 2);
                ContainsImage(tb_snHead3, tb_voltHead3, pb_status3, lb_status3, 3);
                ContainsImage(tb_snHead4, tb_voltHead4, pb_status4, lb_status4, 4);
            }

            if (e.ProgressPercentage == 2) {
                ReadXmlFile();
            }

            if (e.ProgressPercentage == 3) {
                bool flag = true;
                if (!saveDataCsv.head1) {
                    if (!WaitUpPrism(tb_snHead1, tb_voltHead1, prismTest.upPrismFct8_1, 1)) {
                        flag = false;
                    } else {
                        saveDataCsv.head1 = true;
                    }
                }
                if (!saveDataCsv.head2) {
                    if (!WaitUpPrism(tb_snHead2, tb_voltHead2, prismTest.upPrismFct8_2, 2)) {
                        flag = false;
                    } else {
                        saveDataCsv.head2 = true;
                    }
                }
                if (!saveDataCsv.head3) {
                    if (!WaitUpPrism(tb_snHead3, tb_voltHead3, prismTest.upPrismFct8_3, 3)) {
                        flag = false;
                    } else {
                        saveDataCsv.head3 = true;
                    }
                }
                if (!saveDataCsv.head4) {
                    if (!WaitUpPrism(tb_snHead4, tb_voltHead4, prismTest.upPrismFct8_4, 4)) {
                        flag = false;
                    } else {
                        saveDataCsv.head4 = true;
                    }
                }
                if (flag) {
                    prismTest.flagWaitUpPrism = false;
                    imageClass.flagCheck = true;
                    tm_timeOutLog.Enabled = false;
                    bt_exeCute.Text = "EXECUTE";
                    logFile.stopWatch.Stop();
                }
            }

            if (e.ProgressPercentage == 5) {
                ReadXmlFileTimeOut();
            }

            imageClass.flagDoWork = false;
        }
        #endregion


        public class PrismTest {
            /// <summary>nameFile = "prism_config"</summary>
            public string nameFile { get; set; }
            public HeadConfig headConfig { get; set; }
            public string sn { get; set; }
            public string result { get; set; }
            public string dataSummary { get; set; }
            public string fg { get; set; }
            public string desc { get; set; }
            public string orderQty { get; set; }
            public string outputQty { get; set; }
            public string wo { get; set; }
            /// <summary>Value = "SUCCESS"</summary>
            public string success { get; set; }
            /// <summary>Value = "PASS"</summary>
            public string pass { get; set; }
            /// <summary>Value = "FAIL"</summary>
            public string fail { get; set; }
            /// <summary>Value = "WAIT"</summary>
            public string wait { get; set; }
            public MessageErr messageErr { get; set; }
            public XML xml { get; set; }
            /// <summary>เอาไว้บอกให้เตรียมที่จะอัพ Prism</summary>
            public bool flagWaitUpPrism { get; set; }
            public bool upFailToPrism { get; set; }
            public UpPrismFct8 upPrismFct8_1 { get; set; }
            public UpPrismFct8 upPrismFct8_2 { get; set; }
            public UpPrismFct8 upPrismFct8_3 { get; set; }
            public UpPrismFct8 upPrismFct8_4 { get; set; }
            public string mode { get; set; }

            public PrismTest() {
                nameFile = "prism_config";
                headConfig = new HeadConfig();
                success = "SUCCESS";
                messageErr = new MessageErr();
                xml = new XML();
                pass = "PASS";
                fail = "FAIL";
                wait = "WAIT";
                upPrismFct8_1 = new UpPrismFct8();
                upPrismFct8_2 = new UpPrismFct8();
                upPrismFct8_3 = new UpPrismFct8();
                upPrismFct8_4 = new UpPrismFct8();
            }
            public string upData() {
                return TeamPrecision.PRISM.cResults.SaveTestResult(sn, result, dataSummary);
            }
            public string getWO() {
                string result = string.Empty;

                try {
                    string[] getWO = TeamPrecision.PRISM.cSNs.getWO(wo, xml.processName);

                    result = getWO[0];
                    fg = getWO[1];
                    desc = getWO[2];
                    orderQty = getWO[3];
                    outputQty = getWO[4];

                } catch { }

                return result;
            }
            public void ReadConfig(SetupPay.FormPay setupPay) {
                xml.mode = setupPay.read_text(headConfig.mode, nameFile);
                xml.employeeID = setupPay.read_text(headConfig.employeeID, nameFile);
                xml.processName = setupPay.read_text(headConfig.processName, nameFile);
                xml.stationName = setupPay.read_text(headConfig.stationName, nameFile);
                xml.computerName = setupPay.read_text(headConfig.computerName, nameFile);
                xml.databaseName = setupPay.read_text(headConfig.databaseName, nameFile);
                xml.databaseServerTPP = setupPay.read_text(headConfig.databaseServerTPP, nameFile);
                xml.databaseServerTPR = setupPay.read_text(headConfig.databaseServerTPR, nameFile);
                xml.checkProcessBefore = setupPay.read_text(headConfig.checkProcessBefore, nameFile);
                xml.processBefore = setupPay.read_text(headConfig.processBefore, nameFile);
                xml.digitSN = setupPay.read_text(headConfig.digitSN, nameFile);

                upFailToPrism = Convert.ToBoolean(setupPay.read_text(headConfig.upFailToPrism, nameFile));
                mode = setupPay.read_text(headConfig.mode, nameFile);
            }

            public class HeadConfig {
                public string mode { get; set; }
                public string employeeID { get; set; }
                public string processName { get; set; }
                public string stationName { get; set; }
                public string computerName { get; set; }
                public string databaseName { get; set; }
                public string databaseServerTPP { get; set; }
                public string databaseServerTPR { get; set; }
                public string checkProcessBefore { get; set; }
                public string processBefore { get; set; }
                public string digitSN { get; set; }
                public string upFailToPrism { get; set; }

                public HeadConfig() {
                    mode = "Mode";
                    employeeID = "Employee ID";
                    processName = "Process Name";
                    stationName = "Station Name";
                    computerName = "Computer Name";
                    databaseName = "Database Name";
                    databaseServerTPP = "Database Server TPP";
                    databaseServerTPR = "Database Server TPR";
                    checkProcessBefore = "Check Process Before";
                    processBefore = "Process Before";
                    digitSN = "Digit SN";
                    upFailToPrism = "Up Fail To Prism";
                }
            }
            public class MessageErr {
                /// <summary>getWO = "Get WO Error"</summary>
                public string getWO = "Get WO Error";
            }
            public class XML {
                private string databaseServerEnCode { get; set; }
                private string databaseNameEncode { get; set; }
                public string mode { get; set; }
                public string employeeID { get; set; }
                public string processName { get; set; }
                public string stationName { get; set; }
                public string computerName { get; set; }
                public string databaseName { get; set; }
                public string databaseServerTPP { get; set; }
                public string databaseServerTPR { get; set; }
                public string checkProcessBefore { get; set; }
                public string processBefore { get; set; }
                public string digitSN { get; set; }

                public void WriteXml() {
                    if (databaseName == Define.prismTPR) {
                        databaseServerEnCode = EnCode.serverTPR;
                        databaseNameEncode = EnCode.nameTPR;
                    } else {

                        databaseServerEnCode = EnCode.serverTPP;
                        databaseNameEncode = EnCode.nameTPP;
                    }

                    XmlTextWriter writer = new XmlTextWriter(Define.pathXml, Encoding.UTF8);
                    writer.WriteStartDocument();
                    writer.Formatting = Formatting.Indented;
                    writer.Indentation = 2;
                    writer.WriteStartElement("cSettingSerial");
                    writer.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                    writer.WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
                    writer.WriteStartElement("TestingMode");
                    writer.WriteString(mode);
                    writer.WriteEndElement();
                    writer.WriteStartElement("DatabaseServer");
                    writer.WriteString(databaseServerEnCode);
                    writer.WriteEndElement();
                    writer.WriteStartElement("DatabaseName");
                    writer.WriteString(databaseNameEncode);
                    writer.WriteEndElement();
                    writer.WriteStartElement("DatabaseUser");
                    writer.WriteString("qNMPB0293rI=");
                    writer.WriteEndElement();
                    writer.WriteStartElement("DatabasePassword");
                    writer.WriteString("m/2+3pRMmYg=");
                    writer.WriteEndElement();
                    writer.WriteStartElement("ComputerName");
                    writer.WriteString(computerName);
                    writer.WriteEndElement();
                    writer.WriteStartElement("StationName");
                    writer.WriteString(stationName);
                    writer.WriteEndElement();
                    writer.WriteStartElement("ProcessName");
                    writer.WriteString(processName);
                    writer.WriteEndElement();
                    writer.WriteStartElement("UsePasswordWhenLogin");
                    writer.WriteString("false");
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                    writer.Close();
                }

                public static class EnCode {
                    /// <summary>Value = "Vd147+pBWChvWVcRsdZvHQ=="</summary>
                    public static readonly string serverTPR = "Vd147+pBWChvWVcRsdZvHQ==";
                    /// <summary>Value = "awiuCMQfI7kyjwyD3O/AmA=="</summary>
                    public static readonly string nameTPR = "awiuCMQfI7kyjwyD3O/AmA==";
                    /// <summary>Value = "Vd147+pBWCihy3FzdahxTg=="</summary>
                    public static readonly string serverTPP = "Vd147+pBWCihy3FzdahxTg==";
                    /// <summary>Value = "U/AFYtHi4S8yjwyD3O/AmA=="</summary>
                    public static readonly string nameTPP = "U/AFYtHi4S8yjwyD3O/AmA==";
                }
                public static class Define {
                    /// <summary>Value = "TPR_PRISM"</summary>
                    public static readonly string prismTPR = "TPR_PRISM";
                    /// <summary>Value = "TeamPrecision.PRISM.Setting.xml"</summary>
                    public static readonly string pathXml = "TeamPrecision.PRISM.Setting.xml";
                }
            }
            public class UpPrismFct8 {
                public string SerialASM { get; set; }
                public string SerialNumber { get; set; }
                public string DeviceID { get; set; }
                public string ICCID { get; set; }
                public string FirmwareVersion { get; set; }
                public string ProcessName { get; set; }
                public string Job { get; set; }
                public string SIMAPN { get; set; }
                public string Volt { get; set; }
                public string Status { get; set; }
            }
            public static class Define {
                /// <summary>Value = "Debug"</summary>
                public static readonly string debug = "Debug";
            }
        }
        public class Position {
            /// <summary>Value = "image_config"</summary>
            public string nameFile { get; set; }
            public HeadConfig headConfig { get; set; }
            public Number1 number1 { get; set; }
            public Number2 number2 { get; set; }
            public Number3 number3 { get; set; }
            public Number4 number4 { get; set; }
            public Number5 number5 { get; set; }
            public Execute execute { get; set; }
            public int positionFormOpenX { get; set; }
            public int positionFormOpenY { get; set; }

            public Position() {
                nameFile = "image_config";
                headConfig = new HeadConfig();
                number1 = new Number1();
                number2 = new Number2();
                number3 = new Number3();
                number4 = new Number4();
                number5 = new Number5();
                execute = new Execute();
            }

            public class HeadConfig {
                public string vScrollBar { get; set; }
                public string position1_X { get; set; }
                public string position1_Y { get; set; }
                public string position1_Width { get; set; }
                public string position1_Height { get; set; }
                public string position2_X { get; set; }
                public string position2_Y { get; set; }
                public string position2_Width { get; set; }
                public string position2_Height { get; set; }
                public string position3_X { get; set; }
                public string position3_Y { get; set; }
                public string position3_Width { get; set; }
                public string position3_Height { get; set; }
                public string position4_X { get; set; }
                public string position4_Y { get; set; }
                public string position4_Width { get; set; }
                public string position4_Height { get; set; }
                public string position5_X { get; set; }
                public string position5_Y { get; set; }
                public string position5_Width { get; set; }
                public string position5_Height { get; set; }
                public string positionExecute_X { get; set; }
                public string positionExecute_Y { get; set; }
                public string positionFormOpenX { get; set; }
                public string positionFormOpenY { get; set; }

                public HeadConfig() {
                    vScrollBar = "V Scroll Bar";
                    position1_X = "Position Head 1 X";
                    position1_Y = "Position Head 1 Y";
                    position1_Width = "Position Head 1 Width";
                    position1_Height = "Position Head 1 Height";
                    position2_X = "Position Head 2 X";
                    position2_Y = "Position Head 2 Y";
                    position2_Width = "Position Head 2 Width";
                    position2_Height = "Position Head 2 Height";
                    position3_X = "Position Head 3 X";
                    position3_Y = "Position Head 3 Y";
                    position3_Width = "Position Head 3 Width";
                    position3_Height = "Position Head 3 Height";
                    position4_X = "Position Head 4 X";
                    position4_Y = "Position Head 4 Y";
                    position4_Width = "Position Head 4 Width";
                    position4_Height = "Position Head 4 Height";
                    position5_X = "Position Head 5 X";
                    position5_Y = "Position Head 5 Y";
                    position5_Width = "Position Head 5 Width";
                    position5_Height = "Position Head 5 Height";
                    positionExecute_X = "Position Execute X";
                    positionExecute_Y = "Position Execute Y";
                    positionFormOpenX = "Position Form Open X";
                    positionFormOpenY = "Position Form Open Y";
                }
            }
            public class Number1 {
                public int X { get; set; }
                public int Y { get; set; }
                public int Width { get; set; }
                public int Height { get; set; }
            }
            public class Number2 {
                public int X { get; set; }
                public int Y { get; set; }
                public int Width { get; set; }
                public int Height { get; set; }
            }
            public class Number3 {
                public int X { get; set; }
                public int Y { get; set; }
                public int Width { get; set; }
                public int Height { get; set; }
            }
            public class Number4 {
                public int X { get; set; }
                public int Y { get; set; }
                public int Width { get; set; }
                public int Height { get; set; }
            }
            public class Number5 {
                public int X { get; set; }
                public int Y { get; set; }
                public int Width { get; set; }
                public int Height { get; set; }
            }
            public class Execute {
                public int X { get; set; }
                public int Y { get; set; }
            }
        }
        public class LogFile {
            /// <summary>Value = "logFile_config"</summary>
            public string nameFile { get; set; }
            public int testTime { get; set; }
            public HeadConfig headConfig { get; set; }
            public MesSage mesSage { get; set; }
            public string pathLog { get; set; }
            /// <summary>เอาไว้บอกว่ารอบนี้ ได้อ่าน Xml ครบทุกหัวแล้ว</summary>
            public bool readLogFull { get; set; }
            public bool result { get; set; }
            public string deviceID { get; set; }
            public string serialNumber { get; set; }
            public string firmwareVersion { get; set; }
            public string processName { get; set; }
            public string job { get; set; }
            public string simapn { get; set; }
            public bool head1 { get; set; }
            public bool head2 { get; set; }
            public bool head3 { get; set; }
            public bool head4 { get; set; }
            public LogConfig logConfig { get; set; }
            public bool showFail { get; set; }
            public string logFileNameNotePrism { get; set; }
            public Stopwatch stopWatch { get; set; }

            public LogFile() {
                nameFile = "logFile_config";
                headConfig = new HeadConfig();
                mesSage = new MesSage();
                readLogFull = true;
                logConfig = new LogConfig();
                head1 = true;
                head2 = true;
                head3 = true;
                head4 = true;
                stopWatch = new Stopwatch();
                simapn = string.Empty;
            }
            public bool ReadXML(string snCustommer) {
                bool resultSup = false;
                List<string> nameFile = new List<string>();
                try {
                    string[] getFile = Directory.GetFiles(pathLog);
                    nameFile = getFile.ToList<string>();
                } catch { }

                List<string> pathFileSelect = nameFile.FindAll(element => element.Contains(snCustommer));

                foreach (string path in pathFileSelect) {
                    XmlSerializer reader = new XmlSerializer(typeof(ConfigurationReport));
                    StreamReader file;
                    while (true) {
                        try {
                            file = new StreamReader(path);
                            break;
                        } catch { }
                        Thread.Sleep(50);
                    }
                    ConfigurationReport overView = (ConfigurationReport)reader.Deserialize(file);

                    //string[] dateTimeSup1 = overView.DeviceDetails.ReportTime.Split(' ');
                    //string[] dateTimeSup2 = dateTimeSup1[0].Split('/');
                    //string[] dateTimeSup3 = dateTimeSup1[1].Split(':');
                    //int date = Convert.ToInt32(dateTimeSup2[1]);
                    //int month = Convert.ToInt32(dateTimeSup2[0]);
                    //int year = Convert.ToInt32(dateTimeSup2[2]);
                    //int hour = Convert.ToInt32(dateTimeSup3[0]);
                    //int min = Convert.ToInt32(dateTimeSup3[1]);
                    //int sec = Convert.ToInt32(dateTimeSup3[2]);
                    //DateTime now = DateTime.Now;
                    //DateTime dateTimeFile = new DateTime(year, month, date, hour, min, sec);

                    //double minutes = now.Subtract(dateTimeFile).TotalMinutes;  //หน่วยเป็นนาที, วินาทีจะเป็น 0 - 99

                    //if ((minutes * 60) > (testTime / 1000)) {
                    //    continue;
                    //}

                    deviceID = overView.DeviceDetails.DeviceID;
                    serialNumber = overView.DeviceDetails.SerialNumber;
                    firmwareVersion = overView.DeviceDetails.FirmwareVersion;
                    processName = overView.DeviceDetails.ProcessName;
                    job = overView.DeviceDetails.Job;
                    try {
                        simapn = overView.Configuration.SIMAPN;
                    } catch { }

                    string pathPass = "FCT8_PASS\\";
                    string pathFail = "FCT8_FAIL\\";
                    string pathSelect = string.Empty;

                    if (path.Contains(Define.pass)) {
                        result = true;
                        pathSelect = pathPass;

                    } else {
                        pathSelect = pathFail;
                    }
                    
                    if (!Directory.Exists(pathLog + pathPass)) {
                        Directory.CreateDirectory(pathLog + pathPass);
                    }
                    if (!Directory.Exists(pathLog + pathFail)) {
                        Directory.CreateDirectory(pathLog + pathFail);
                    }

                    file.Close();
                    logFileNameNotePrism = path.Replace(pathLog, string.Empty);
                    int num = 1;
                    string nameLogSup = string.Empty;
                    while (true) {
                        try {
                            File.Move(path, pathLog + pathSelect + logFileNameNotePrism.Replace(".xml", string.Empty) + nameLogSup + ".xml");
                            break;
                        } catch {
                            nameLogSup = "_" + num;
                            num++;
                        }
                        Thread.Sleep(50);
                    }
                    resultSup = true;
                    break;
                }

                return resultSup;
            }
            public void Clear(Form1 form) {
                List<string> nameFile = new List<string>();
                try {
                    string[] getFile = Directory.GetFiles(pathLog);
                    nameFile = getFile.ToList<string>();
                } catch { }

                List<string> pathFileSelect = nameFile.FindAll(element => element.Contains(form.tb_snHead1.Text));
                foreach (string path in pathFileSelect) {
                    File.Delete(path);
                }

                pathFileSelect = nameFile.FindAll(element => element.Contains(form.tb_snHead2.Text));
                foreach (string path in pathFileSelect) {
                    File.Delete(path);
                }

                pathFileSelect = nameFile.FindAll(element => element.Contains(form.tb_snHead3.Text));
                foreach (string path in pathFileSelect) {
                    File.Delete(path);
                }

                pathFileSelect = nameFile.FindAll(element => element.Contains(form.tb_snHead4.Text));
                foreach (string path in pathFileSelect) {
                    File.Delete(path);
                }
            }

            #region Class File Xml
            public class Source {
                public string Application { get; set; }
            }
            public class DeviceDetails {
                public string ReportTime { get; set; }
                public string ProductName { get; set; }
                public string SerialNumber { get; set; }
                public string DeviceID { get; set; }
                public string FirmwareVersion { get; set; }
                public string ProcessName { get; set; }
                public string Job { get; set; }
            }
            public class Configuration {
                public string UnitStartTimeStamp { get; set; }
                public string UnitStopTimeStamp { get; set; }
                public string NumberofCustomerResets { get; set; }
                public string UserInformation { get; set; }
                public string TimestampConfigurationTime { get; set; }
                public string ExternalPowerONResetCounter { get; set; }
                public string EEPROMSize { get; set; }
                public string StartUpDelayTime { get; set; }
                public string TripLength { get; set; }
                public string TemperatureMeasurementInterval { get; set; }
                public string ControlByte0 { get; set; }
                public string ControlByte1 { get; set; }
                public string MinimumBatteryVoltage { get; set; }
                public string LowBatteryThreshold { get; set; }
                public string LowBatteryDeadband { get; set; }
                public string BatteryLevelSamplingInterval { get; set; }
                public string LightSensorSamplingInterval { get; set; }
                public string RunModeCheckInInterval { get; set; }
                public string PhoneHomeCheckInInterval { get; set; }
                public string CellScanMethod { get; set; }
                public string SocketPrimaryIPAddress { get; set; }
                public string SocketPrimaryPort { get; set; }
                public string SocketTimeout { get; set; }
                public string MaximumRetries { get; set; }
                public string MinimumSignal { get; set; }
                public string BeeperDuration { get; set; }
                public string BeeperInterval { get; set; }
                public string NTPIPAddress { get; set; }
                public string NTPPort { get; set; }
                public string NTPSyncRetryInterval { get; set; }
                public string SIMAPN { get; set; }
                public string SIMUsername { get; set; }
                public string SIMPassword { get; set; }
                public string SIMAuthenticationType { get; set; }
                public string MaxModemOnTime { get; set; }
                public string ModemActivityTimeout { get; set; }
                public string ModemOFFTime { get; set; }
                public string ModemMinimumBatteryLevel { get; set; }
                public string CPUType { get; set; }
                public string HibernateCheckInInterval { get; set; }
                public string OutofCoverageWakeupInterval { get; set; }
                public string OutofCoverageNoNetworkTime { get; set; }
                public string AlphaNumericSN { get; set; }
                public string ConfigurationFlag { get; set; }
                public string HumidityMeasurementInterval { get; set; }
                public string SensorEnable1 { get; set; }
                public string SensorEnable2 { get; set; }
            }
            public class Transforms {
                public string Transform { get; set; }
            }
            public class Reference {
                public Transforms Transforms { get; set; }
                public string DigestMethod { get; set; }
                public string DigestValue { get; set; }
            }
            public class SignedInfo {
                public string CanonicalizationMethod { get; set; }
                public string SignatureMethod { get; set; }
                public Reference Reference { get; set; }
            }
            public class Signature {
                public SignedInfo SignedInfo { get; set; }
                public string SignatureValue { get; set; }
            }
            public class ConfigurationReport {
                public Source Source { get; set; }
                public DeviceDetails DeviceDetails { get; set; }
                public Configuration Configuration { get; set; }
                public Signature Signature { get; set; }
            }
            #endregion

            public class HeadConfig {
                public string setTestTime { get; set; }
                public string firmwareVersion { get; set; }
                public string processName { get; set; }
                public string job { get; set; }
                public string simapn { get; set; }
                public string pathLog { get; set; }
                public string showFail { get; set; }

                public HeadConfig() {
                    setTestTime = "Set Test Time";
                    firmwareVersion = "Firmware Version";
                    processName = "Process Name";
                    job = "Job";
                    simapn = "SIMAPN";
                    pathLog = "Path Log";
                    showFail = "Show Fail";
                }
            }
            public class MesSage {
                public string error { get; set; }

                public MesSage() {
                    error = "Error!! Read Log Fail Xml";
                }
            }
            public static class Define {
                /// <summary>Value = "dd/MM/yyyy"</summary>
                public static readonly string dateFormat = "dd/MM/yyyy";
                /// <summary>Value = "HH:mm:ss"</summary>
                public static readonly string timeFormat = "HH:mm:ss";
                /// <summary>Value = "en-US"</summary>
                public static readonly string usFormat = "en-US";
                /// <summary>Value = "Passed"</summary>
                public static readonly string pass = "Passed";
            }
            public class LogConfig {
                public string firmwareVersion { get; set; }
                public string processName { get; set; }
                public string job { get; set; }
                public string simapn { get; set; }
            }
        }
        public class TestSN {
            /// <summary>Value = 4</summary>
            public int maxHead { get; set; }
            public int length { get; set; }
            public bool[] flagTurbo { get; set; }
            public bool[] flagReady { get; set; }
            public MesSage mesSage { get; set; }
            public Stopwatch timeOut1 { get; set; }
            public Stopwatch timeOut2 { get; set; }
            public Stopwatch timeOut3 { get; set; }
            public Stopwatch timeOut4 { get; set; }
            public int setTimeOut { get; set; }
            public HeadConfig headConfig { get; set; }
            public string scanSnSupport { get; set; }

            public TestSN() {
                maxHead = 4;
                flagTurbo = new bool[4];
                flagReady = new bool[] { true, true, true, true };
                mesSage = new MesSage();
                timeOut1 = new Stopwatch();
                timeOut2 = new Stopwatch();
                timeOut3 = new Stopwatch();
                timeOut4 = new Stopwatch();
                timeOut1.Start();
                timeOut2.Start();
                timeOut3.Start();
                timeOut4.Start();
                headConfig = new HeadConfig();
            }
            public bool CheckTimeOut(int head) {
                int overTime = 99999;

                switch (head) {
                    case 1:
                        if (timeOut1.ElapsedMilliseconds < setTimeOut) {
                            return false;
                        }
                        if (timeOut1.ElapsedMilliseconds > overTime) {
                            timeOut1.Restart();
                        }
                        return true;
                    case 2:
                        if (timeOut2.ElapsedMilliseconds < setTimeOut) {
                            return false;
                        }
                        if (timeOut2.ElapsedMilliseconds > overTime) {
                            timeOut2.Restart();
                        }
                        return true;
                    case 3:
                        if (timeOut3.ElapsedMilliseconds < setTimeOut) {
                            return false;
                        }
                        if (timeOut3.ElapsedMilliseconds > overTime) {
                            timeOut3.Restart();
                        }
                        return true;
                    case 4:
                        if (timeOut4.ElapsedMilliseconds < setTimeOut) {
                            return false;
                        }
                        if (timeOut4.ElapsedMilliseconds > overTime) {
                            timeOut4.Restart();
                        }
                        return true;
                }

                return true;
            }
            public void ClearTimer(int head) {
                switch (head) {
                    case 1:
                        timeOut1.Restart();
                        break;
                    case 2:
                        timeOut2.Restart();
                        break;
                    case 3:
                        timeOut3.Restart();
                        break;
                    case 4:
                        timeOut4.Restart();
                        break;
                }
            }

            public class HeadConfig {
                public string setTimeOut { get; set; }

                public HeadConfig() {
                    setTimeOut = "Set Time Out";
                }
            }
            public class MesSage {
                /// <summary>Value = "SN ครบจำนวนแล้ว"</summary>
                public string snFull { get; set; }

                public MesSage() {
                    snFull = "SN ครบจำนวนแล้ว";
                }
            }
        }
        public class SaveDataLogCsv {
            /// <summary>Value = "saveData_config"</summary>
            public string nameFile { get; set; }
            public string path { get; set; }
            public HeadConfig headConfig { get; set; }
            public string pathPass { get; set; }
            public string pathFail { get; set; }
            public bool head1 { get; set; }
            public bool head2 { get; set; }
            public bool head3 { get; set; }
            public bool head4 { get; set; }

            public SaveDataLogCsv() {
                nameFile = "saveData_config";
                headConfig = new HeadConfig();
            }
            public class HeadConfig {
                public string path { get; set; }

                public HeadConfig() {
                    path = "Path";
                }
            }
        }
        public class Focus {
            public bool newUnit1 { get; set; }
            public bool newUnit2 { get; set; }
            public bool newUnit3 { get; set; }
            public bool newUnit4 { get; set; }

            public void Clear() {
                newUnit1 = false;
                newUnit2 = false;
                newUnit3 = false;
                newUnit4 = false;
            }
        }
        public class Tester {
            /// <summary>Value = "tester_config"</summary>
            public string nameFile { get; set; }
            public bool debugTestBattery { get; set; }
            public HeadConfig headConfig { get; set; }

            public Tester() {
                nameFile = "tester_config";
                headConfig = new HeadConfig();
            }
            public void GetConfig(Form1 form) {
                debugTestBattery = Convert.ToBoolean(form.setupPay.read_text(headConfig.debugTestBattery, nameFile));
            }
            public class HeadConfig {
                public string debugTestBattery { get; set; }

                public HeadConfig() {
                    debugTestBattery = "Debug test Battery";
                }
            }
        }
    }
}
