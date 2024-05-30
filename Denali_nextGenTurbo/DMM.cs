using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Denali_nextGenTurbo {
    public class DMM {
        private Connect dmm { get; set; }
        public string nameDMM { get; set; }
        private string nameDmmConnect { get; set; }
        /// <summary>Value = "IDM-8341"</summary>
        private string dmmIdm8341 { get; set; }
        /// <summary>Value = "U3606B"</summary>
        private string dmmU3606 { get; set; }
        private Define define { get; set; }
        public Relay relay { get; set; }
        public bool flagDoWork { get; set; }
        public Spec spec { get; set; }
        /// <summary>เอาไว้บอกให้เริ่มอ่าน และบอกว่าอ่านเสร็จทุกหัวแล้ว</summary>
        public bool flagRead { get; set; }
        private string rxReceive { get; set; }
        private bool rxReceiveFlag { get; set; }
        private SerialPort serialPort { get; set; }
        public bool[] statusRunning { get; set; }

        public DMM() {
            dmm = new Connect();
            define = new Define();
            relay = new Relay();
            spec = new Spec();
            dmmIdm8341 = "IDM-8341";
            dmmU3606 = "U3606B";
            serialPort = new SerialPort();
            statusRunning = new bool[4];
        }
        public void ReadVolt(Form1 form, TextBox textBox, int head) {
            textBox.Text = "Reading...";

            form.Log(Form1.LogMsgType.Incoming_Blue, "\nOn Relay Card Head " + head);
            OnRelay(head);

            if (nameDMM == dmmIdm8341) {
                DmmReadIdm8341(form, textBox, head);
            }
            if (nameDMM == dmmU3606) {
                DmmReadU3606(form, textBox, head);
            }

            form.Log(Form1.LogMsgType.Incoming_Blue, "\nOff Relay Card Head " + head);
            OffRelay(head);
        }
        public void ClearTextBox(TextBox textBox) {
            textBox.Text = "";
            textBox.BackColor = Color.White;
        }
        private void OnRelay(int head) {
            switch (head) {
                case 1:
                    relay.Relay_On(1, 9);
                    relay.Relay_On(1, 13);
                    break;
                case 2:
                    relay.Relay_On(1, 10);
                    relay.Relay_On(1, 14);
                    break;
                case 3:
                    relay.Relay_On(1, 11);
                    relay.Relay_On(1, 15);
                    break;
                case 4:
                    relay.Relay_On(1, 12);
                    relay.Relay_On(1, 16);
                    break;
            }
        }
        private void OffRelay(int head) {
            switch (head) {
                case 1:
                    relay.Relay_Off(1, 9);
                    relay.Relay_Off(1, 13);
                    break;
                case 2:
                    relay.Relay_Off(1, 10);
                    relay.Relay_Off(1, 14);
                    break;
                case 3:
                    relay.Relay_Off(1, 11);
                    relay.Relay_Off(1, 15);
                    break;
                case 4:
                    relay.Relay_Off(1, 12);
                    relay.Relay_Off(1, 16);
                    break;
            }
        }
        public void FinePortDmm() {
            if (nameDMM == dmmIdm8341) {
                FinePortIDM8341();
            }
            if (nameDMM == dmmU3606) {
                FinePortU3606();
            }
        }
        private void FinePortU3606() {
            nameDmmConnect = string.Empty;

            Ivi.Visa.Interop.ResourceManager mngr = new Ivi.Visa.Interop.ResourceManager();
            try {
                string[] namePort = mngr.FindRsrc("?*");
                foreach (string port in namePort) {
                    if (port.Contains("0x0957::0x4D18")) {
                        nameDmmConnect = port;
                    }
                }
            } catch { }

            if (nameDmmConnect == string.Empty) {
                MessageBox.Show("DMM Error!!");
            }
        }
        private void FinePortIDM8341() {
            nameDmmConnect = string.Empty;

            ManagementObjectSearcher objOSDetails2 = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'");
            ManagementObjectCollection osDetailsCollection2 = objOSDetails2.Get();
            foreach (ManagementObject usblist in osDetailsCollection2) {
                if (!usblist["Description"].ToString().Contains("IDM834")) continue;
                string[] port = usblist.GetPropertyValue("NAME").ToString().Split('(', ')');
                nameDmmConnect = port[1];
            }

            if (nameDmmConnect == string.Empty) {
                MessageBox.Show("DMM Error!!");
                return;
            }

            serialPort.PortName = nameDmmConnect;
            serialPort.BaudRate = 9600;
            serialPort.DataBits = 8;
            serialPort.StopBits = StopBits.One;
            serialPort.Parity = Parity.None;
            serialPort.Handshake = Handshake.None;
            serialPort.RtsEnable = true;
            serialPort.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);
        }

        private void DisconnectU3606() {
            dmm.DisConnectInstr();
        }
        private bool ConnectU3606() {
            if (dmm.ConnectInstr(nameDmmConnect) == define.connected) {
                return true;

            } else {
                return false;
            }
        }
        private void DmmReadU3606(Form1 form, TextBox textBox, int head) {
            double value = 0;
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Restart();

            ConnectU3606();
            while (stopWatch.ElapsedMilliseconds <= 5000 && statusRunning[head - 1]) {
                string resultString = dmm.QueryString(":MEAS:VOLT:DC?");

                try {
                    value = Convert.ToDouble(resultString);
                } catch { }

                form.Log(Form1.LogMsgType.Incoming_Blue, "\nRead Volt Dc = " + value);
                if (value > spec.max || value < spec.min) {
                    textBox.Text = value.ToString();
                    //Thread.Sleep(50);
                    form.DelaymS(50);
                    continue;
                }

                textBox.Text = value.ToString();
                textBox.BackColor = Color.Lime;
                break;
            }

            DisconnectU3606();
            if (textBox.BackColor != Color.Lime && statusRunning[head - 1])
            {
                textBox.BackColor = Color.Red;
            }
        }
        private void DmmReadIdm8341(Form1 form, TextBox textBox, int head) {
            double value = 0;
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Restart();

            if (!OpenPortIdm8341(form, head)) {
                if (statusRunning[head - 1]) {
                    textBox.Text = "Dmm Error!!";
                    textBox.BackColor = Color.Red;
                }
                return;
            }

            serialPort.Write("CONF:VOLT:DC DEF\n");

            while (stopWatch.ElapsedMilliseconds <= 5000 && statusRunning[head - 1]) {
                if (!WriteIdm8341(form, "VAL1?\n")) {
                    textBox.Text = "Dmm Error!!";
                    textBox.BackColor = Color.Red;
                    continue;
                }

                try {
                    value = Convert.ToDouble(rxReceive);
                } catch { }

                form.Log(Form1.LogMsgType.Incoming_Blue, "\nRead Volt Dc = " + value);
                if (value > spec.max || value < spec.min) {
                    textBox.Text = value.ToString();
                    //Thread.Sleep(50);
                    form.DelaymS(50);
                    continue;
                }

                textBox.Text = value.ToString();
                textBox.BackColor = Color.Lime;
                break;
            }

            if (textBox.BackColor != Color.Lime && statusRunning[head - 1])
            {
                textBox.BackColor = Color.Red;
            }
        }
        private void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e) {
            Thread.Sleep(50);
            rxReceive = serialPort.ReadExisting();
            rxReceive = rxReceive.Replace("\n", "").Replace("\r", "");
            serialPort.DiscardInBuffer();
            serialPort.DiscardOutBuffer();
            rxReceiveFlag = true;
        }
        private bool OpenPortIdm8341(Form1 form, int head) {
            Stopwatch time = new Stopwatch();
            time.Restart();

            while (time.ElapsedMilliseconds < 3000 && statusRunning[head - 1]) {
                try {
                    serialPort.Open();
                    time.Stop();
                    break;
                } catch {
                    form.DelaymS(250);
                }

                try {
                    serialPort.Close();
                } catch { }
            }

            if (time.IsRunning) {
                try {
                    serialPort.Close();
                } catch { }

                return false;
            }

            return true;
        }
        private bool WriteIdm8341(Form1 form, string cmd) {
            Stopwatch timeSup = new Stopwatch();

            rxReceiveFlag = false;
            serialPort.Write(cmd);
            timeSup.Restart();

            while (timeSup.ElapsedMilliseconds <= 2000) {
                if (!rxReceiveFlag) {
                    form.DelaymS(50);
                    continue;
                }
                rxReceiveFlag = false;
                timeSup.Stop();
                break;
            }

            if (timeSup.IsRunning) {
                return false;
            }

            return true;
        }

        public class Define {
            public string connected { get; set; }
            /// <summary>Value = "PASS"</summary>
            public string pass { get; set; }
            /// <summary>Value = "FAIL"</summary>
            public string fail { get; set; }
            /// <summary>Value = "WAIT"</summary>
            public string wait { get; set; }

            public Define() {
                connected = "Connected";
                pass = "PASS";
                fail = "FAIL";
                wait = "WAIT";
            }
        }
        public class Relay {
            public SerialPort relayPort { get; set; }
            private Form formArduino { get; set; }
            public MesSage_ mesSage { get; set; }
            public bool status { get; set; }
            public string port { get; set; }

            public Relay() {
                relayPort = new SerialPort();
                mesSage = new MesSage_();
            }
            public bool Connect(Form1 form) {
                ClosePort();
                SetFormFinePort();
                bool flagRelay = FinePort(form);
                formArduino.Close();

                if (!flagRelay) {
                    MessageBox.Show(MesSage.connectError);
                    return false;
                }

                if (!relayPort.IsOpen) {
                    MessageBox.Show(MesSage.noDevice);
                    return false;
                }

                Write(CMD.initial);

                string[] resultScanI2c = Write(CMD.scanI2c).Replace("\r\n", "|").Split('|');
                if (resultScanI2c.Length != 2) {
                    MessageBox.Show(MesSage.scanI2cError);
                    return false;
                }

                if (!CheckCard(resultScanI2c[1]))
                {
                    return false;
                }

                OffAllRelay();
                status = true;
                return true;
            }
            private void SetFormFinePort() {
                Label label = new Label();
                FontFamily fontFamily = new FontFamily("Arial");

                label.Text = "Find prot arduino...";
                label.Size = new Size(350, 75);
                label.Font = new Font(fontFamily, 30, FontStyle.Bold, GraphicsUnit.Pixel);

                formArduino = new Form();
                formArduino.Size = new Size(400, 100);
                formArduino.ControlBox = false;
                formArduino.StartPosition = FormStartPosition.CenterScreen;
                formArduino.Controls.Add(label);
                formArduino.Show();
            }
            public static void DelaymS(int mS) {
                Stopwatch stopwatchDelaymS = new Stopwatch();
                stopwatchDelaymS.Restart();
                while (mS > stopwatchDelaymS.ElapsedMilliseconds) {
                    if (!stopwatchDelaymS.IsRunning) stopwatchDelaymS.Start();
                    Application.DoEvents();
                }
                stopwatchDelaymS.Stop();
            }
            private void ClosePort() {
                try {
                    relayPort.Close();
                } catch { }
            }
            private void GetPort() {

            }
            public string Write(string cmd) {
                string result = string.Empty;
                Stopwatch stopwatch = new Stopwatch();

                try {
                    relayPort.DiscardInBuffer();
                    relayPort.DiscardOutBuffer();
                    relayPort.Write(cmd);
                } catch {

                    return MesSage.writeError;
                }

                DelaymS(500);

                result = Read();
                DelaymS(250);
                result += ReadAgain();

                return result;
            }
            private string Read() {
                string result = string.Empty;
                Stopwatch stopwatch = new Stopwatch();

                stopwatch.Restart();
                while (stopwatch.ElapsedMilliseconds < 2500) {
                    try {
                        result = relayPort.ReadExisting();
                    } catch { }

                    if (result == "") {
                        DelaymS(100);
                        continue;
                    }

                    break;
                }

                return result;
            }
            private string ReadAgain() {
                string result = string.Empty;

                while (true) {
                    string resultSup = string.Empty;

                    try {
                        resultSup = relayPort.ReadExisting();
                    } catch { }

                    if (resultSup != "") {
                        result += resultSup;
                        DelaymS(100);
                        continue;
                    }

                    break;
                }

                return result;
            }
            private bool FinePort(Form1 form) {
                bool flagOpen = false;

                if (!OpenPort(port)){
                    goto LabelFinePort;
                }

                string rxSup = Write(CMD.checkPort);

                if (rxSup.Contains(Rx.checkPort)){
                    flagOpen = true;
                    return flagOpen;

                } else{
                    ClosePort();
                }

            LabelFinePort:
                ManagementObjectSearcher Manager = new ManagementObjectSearcher(PORT.nameManager);
                ManagementObjectCollection Manager2 = Manager.Get();

                foreach (ManagementObject portList in Manager2) {

                    if (portList[PORT.descripTion].ToString() != PORT.nameArduino1 &&
                        portList[PORT.descripTion].ToString() != PORT.nameArduino2 &&
                        portList[PORT.descripTion].ToString() != PORT.nameArduino3 &&
                        portList[PORT.descripTion].ToString() != PORT.nameArduino4) continue;

                    string[] arrport = portList.GetPropertyValue("NAME").ToString().Split('(', ')');

                    if(arrport.Length < 2) {
                        continue;
                    }

                    if (!OpenPort(arrport[1])) {
                        continue;
                    }

                    string rx = Write(CMD.checkPort);

                    if (rx.Contains(Rx.checkPort)) {
                        form.setupPay.write_text(form.dmm.spec.headConfig.comportRelay, arrport[1], form.dmm.spec.nameFile);
                        flagOpen = true;
                        break;

                    } else {
                        ClosePort();
                        continue;
                    }
                }

                return flagOpen;
            }
            private bool OpenPort(string port) {
                bool flagOpen = false;
                Stopwatch stopwatch = new Stopwatch();

                relayPort = new SerialPort(port);
                stopwatch.Restart();
                while (stopwatch.ElapsedMilliseconds < 2500) {
                    ClosePort();

                    try {
                        relayPort.Open();
                    } catch {
                        DelaymS(250);
                        continue;
                    }

                    DelaymS(500);
                    flagOpen = true;
                    break;
                }

                return flagOpen;
            }
            private bool CheckCard(string scanI2c) {
                bool result = true;
                int numRelayCard = 1;

                List<string> scanI2cList = new List<string>(scanI2c.Split(','));
                scanI2cList.RemoveAll(item => item == "");
                List<string> listCard = new List<string> { "20", "21", "22", "23", "24", "25", "26", "27" };
                for (int i = 1; i <= 8; i++) {
                    bool flag = false;
                    foreach (string list in scanI2cList) {
                        if (listCard.Contains(list)) {
                            scanI2cList.Remove(list);
                            flag = true;
                            break;
                        }
                    }

                    if (i <= numRelayCard && !flag) {
                        MessageBox.Show("_ใส่การ์ดรีเลย์ไม่ครบ");
                        result = false;
                        break;
                    }
                }

                return result;
            }
            public void OffAllRelay() {
                Write(CMD.offAll);
            }
            public void Relay_On(int card, int bit) {
                RelayControl(card.ToString(), bit.ToString(), "1");
                DelaymS(50);
            }
            public void Relay_Off(int card, int bit) {
                RelayControl(card.ToString(), bit.ToString(), "0");
                DelaymS(50);
            }
            public void RelayControl(string card, string bit, string cmd) {
                string[] result = { "" };
                for (int i = 1; i <= 3; i++) {
                    result = write_23017("23017," + card + "," + bit + "," + cmd + "\n").Replace("\r\n", "|").Split('|');
                    try { string hh = result[2]; } catch { continue; }
                    if (result[2] == "65535") continue;//65535
                    break;
                }
            }
            public string write_23017(string cmd) {
                label_write_23017_sup:
                try {
                    relayPort.DiscardInBuffer();
                    relayPort.DiscardOutBuffer();
                    relayPort.Write(cmd);
                } catch { }
                string rx = "";
                Stopwatch s = new Stopwatch();
                s.Restart();
                while (s.ElapsedMilliseconds < 2500) {
                    //DelaymS(250);
                    try { rx = relayPort.ReadExisting(); } catch { }
                    if (rx != "") {
                        //LogResetArduino2(rx, cmd);
                        s.Stop();
                        break; 
                    }
                    DelaymS(50);
                }
                if (s.IsRunning) {
                    try {
                        relayPort.DtrEnable = true;
                        relayPort.RtsEnable = true;
                    } catch { }
                    DelaymS(5000);
                    try {
                        relayPort.DtrEnable = false;
                        relayPort.RtsEnable = false;
                    } catch { }
                    try { relayPort.Close(); } catch { }
                    try { relayPort.Open(); } catch { }
                    LogResetArduino(cmd);
                    goto label_write_23017_sup;
                }
                string sup_rx = "";
                s.Restart();
                while (s.ElapsedMilliseconds < 100) {
                    try { sup_rx = relayPort.ReadExisting(); } catch { }
                    if (sup_rx != "") { rx += sup_rx; s.Restart(); continue; }
                    DelaymS(10);
                }
                return rx;
            }
            /// <summary>
            /// Save log reset arduino to csv
            /// </summary>
            private void LogResetArduino(string cmd) {
                string path = "D:\\LogResetArduino";

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                DateTime now = DateTime.Now;
                StreamWriter swOut = new StreamWriter(path + "\\" + now.Year + "_" + now.Month + ".csv", true);
                string time = now.Day.ToString("00") + ":" + now.Hour.ToString("00") + ":" + now.Minute.ToString("00") + ":" + now.Second.ToString("00");
                swOut.WriteLine(time + ",Reset Arduino," + cmd);
                swOut.Close();
            }
            private void LogResetArduino2(string rx, string cmd) {
                string path = "D:\\LogWriteArduino";

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                DateTime now = DateTime.Now;
                StreamWriter swOut = new StreamWriter(path + "\\" + now.Year + "_" + now.Month + ".csv", true);
                string time = now.Day.ToString("00") + ":" + now.Hour.ToString("00") + ":" + now.Minute.ToString("00") + ":" + now.Second.ToString("00");
                swOut.WriteLine(time + ",Write," + cmd + "," + rx);
                swOut.Close();
            }


            public static class PORT {
                /// <summary>Value = "SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"</summary>
                public static readonly string nameManager = "SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'";
                /// <summary>Value = "Description"</summary>
                public static readonly string descripTion = "Description";
                /// <summary>Value = "USB-SERIAL CH340"</summary>
                public static readonly string nameArduino1 = "USB-SERIAL CH340";
                /// <summary>Value = "USB Serial Port"</summary>
                public static readonly string nameArduino2 = "USB Serial Port";
                /// <summary>Value = "USB Serial Device"</summary>
                public static readonly string nameArduino3 = "USB Serial Device";
                /// <summary>Value = "Arduino Mega 2560"</summary>
                public static readonly string nameArduino4 = "Arduino Mega 2560";
            }
            public static class MesSage {
                /// <summary>Value = "Write Error"</summary>
                public static readonly string writeError = "Write Error";
                /// <summary>Value = "Cannot connect to Realay Card!!"</summary>
                public static readonly string connectError = "Cannot connect to Realay Card!!";
                /// <summary>Value = "USB-SERIAL CH340\" not have in device"</summary>
                public static readonly string noDevice = "USB-SERIAL CH340\" not have in device";
                /// <summary>Value = "23017,SCANI2C\" retrun not format"</summary>
                public static readonly string scanI2cError = "23017,SCANI2C\" retrun not format";
                /// <summary>Value = "Relay Card : Connect\n"</summary>
                public static readonly string connected = "Relay Card : Connect\n";
                /// <summary>Value = "Relay Card : No Connect\n"</summary>
                public static readonly string connectErr = "Relay Card : No Connect\n";
            }
            public static class CMD {
                /// <summary>Value = "23017,CHECKPORT\n"</summary>
                public static readonly string checkPort = "23017,CHECKPORT\n";
                /// <summary>Value = "23017,INITIAL\n"</summary>
                public static readonly string initial = "23017,INITIAL\n";
                /// <summary>Value = "23017,SCANI2C\n"</summary>
                public static readonly string scanI2c = "23017,SCANI2C\n";
                /// <summary>Value = "23017,OFFALLRELAY\n"</summary>
                public static readonly string offAll = "23017,OFFALLRELAY\n";
            }
            public static class Rx {
                /// <summary>Value = "RELAY_PORT_BY_DESIGN"</summary>
                public static readonly string checkPort = "RELAY_PORT_BY_DESIGN";
            }
            public class MesSage_ {
                public string connected = MesSage.connected;
                public string connectErr = MesSage.connectErr;
            }
        }
        public class Spec {
            /// <summary>nameFile = "spec_config"</summary>
            public string nameFile { get; set; }
            public string allFg { get; set; }
            public double min { get; set; }
            public double max { get; set; }
            public HeadConfig headConfig { get; set; }
            /// <summary>Value = "Read Spec Min Max Error"</summary>
            public string error { get; set; }
            public int delayWaitExeCute { get; set; }

            public Spec() {
                nameFile = "spec_config";
                headConfig = new HeadConfig();
                error = "Read Spec Min Max Error";
            }
            public class HeadConfig {
                public string nameDmm { get; set; }
                public string comportRelay { get; set; }
                public string batteryFg1 { get; set; }
                public string batteryMin1 { get; set; }
                public string batteryMax1 { get; set; }
                public string batteryFg2 { get; set; }
                public string batteryMin2 { get; set; }
                public string batteryMax2 { get; set; }
                public string batteryFg3 { get; set; }
                public string batteryMin3 { get; set; }
                public string batteryMax3 { get; set; }
                public string batteryFg4 { get; set; }
                public string batteryMin4 { get; set; }
                public string batteryMax4 { get; set; }
                public string delayWaitExeCute { get; set; }
                public HeadConfig() {
                    nameDmm = "Name Dmm";
                    comportRelay = "Comport Relay";
                    batteryFg1 = "Battery 1 (FG)";
                    batteryMin1 = "Battery 1 (Min)";
                    batteryMax1 = "Battery 1 (Max)";
                    batteryFg2 = "Battery 2 (FG)";
                    batteryMin2 = "Battery 2 (Min)";
                    batteryMax2 = "Battery 2 (Max)";
                    batteryFg3 = "Battery 3 (FG)";
                    batteryMin3 = "Battery 3 (Min)";
                    batteryMax3 = "Battery 3 (Max)";
                    batteryFg4 = "Battery 4 (FG)";
                    batteryMin4 = "Battery 4 (Min)";
                    batteryMax4 = "Battery 4 (Max)";
                    delayWaitExeCute = "Delay Wait ExeCute";
                }
            }
        }
    }
}
