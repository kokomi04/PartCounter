using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Xml;
using System.Timers;
using System.IO;
using System.Threading;
using System.Diagnostics;
using Microsoft.Win32;
using PartsCounter;
using System.Drawing.Printing;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AppRS232
{
    public partial class Setup : Form
    {
        bool boolcheckRuler = true;
        bool running = false;
        SerialPort P = new SerialPort();
        SerialPort P1 = new SerialPort();
        string InputData = string.Empty;
        string InputData1 = string.Empty;
        string DataSN = string.Empty;
        string DataThick = string.Empty;
        string DataMax;
        string DataPrintSN;
        string pathcurrentNow = Path.GetDirectoryName(Application.ExecutablePath);
        delegate void SetTextCallback(string Text);
        const double pi = 3.14159;
        SQLiteDatabase newSQLite = new SQLiteDatabase("test.db");
        protected override void WndProc(ref Message message)
        {
            if (message.Msg == SingleInstance.WM_SHOWFIRSTINSTANCE)
            {
                ShowWindow();
            }
            base.WndProc(ref message);
        }
        public void ShowWindow()
        {
            Win32API.ShowToFront(this.Handle);
        }
        public Setup()
        {
            InitializeComponent();
            Path.GetDirectoryName(Application.ExecutablePath);
            /*if (!File.Exists("C:\\Windows\\Fonts\\gunship.ttf"))
            {
                File.Copy(Path.GetDirectoryName(Application.ExecutablePath) + "\\gunship.ttf", "C:\\Windows\\Fonts");
                rtb_Log.Text += Environment.NewLine + "Copy OK ^^ ";
            }*/
            string version1 = "1.1";
            tmSystem.Start();
            this.CenterToScreen();
            string[] ports = SerialPort.GetPortNames();
            txtCOM.Items.AddRange(ports);
            cbCOMport.Items.AddRange(ports);
            //P+P1 port
            P.ReadTimeout = 500;
            P.WriteTimeout = 500;
            P.DataReceived += new SerialDataReceivedEventHandler(DataReceive);
            P1.ReadTimeout = 500;
            P1.WriteTimeout = 500;
            P1.DataReceived += new SerialDataReceivedEventHandler(DataReceive1);
            string[] BaudRate = { "1200", "2400", "4800", "9600", "19200", "38400", "57600", "115200" };
            txtBaudRate.Items.AddRange(BaudRate);
            string[] DataBits = { "6", "7", "8" };
            txtDataBits.Items.AddRange(DataBits);
            string[] Parity = { "None", "Odd", "Even" };
            txtParity.Items.AddRange(Parity);
            string[] StopBits = { "1", "1.5", "2" };
            txtStopBit.Items.AddRange(StopBits);
            this.Text = "Part Counter - Version : " + version1 + " (BuildDate : " + File.GetLastWriteTime("PartsCounter.exe") + ")";
            cbbPicth.SelectedItem = "2";
            txtPicthGet.Visible = false;                       
        }

        private void txtSetup_Load(object sender, EventArgs e)
        {
            //-----this code to fix screen full fill
            this.Height = Screen.PrimaryScreen.WorkingArea.Height;
            this.Width = Screen.PrimaryScreen.WorkingArea.Width;
            this.Location = Screen.PrimaryScreen.WorkingArea.Location;
            this.WindowState = FormWindowState.Maximized;
            //-------------
            RegistryKey key = Registry.CurrentUser.CreateSubKey(@"HKEY_LOCAL_MACHINE\\SOFTWARE\\Ambit\\CMTestProgram\\");
            if (key.GetValue("boolVitme") != null)
            {
                if (key.GetValue("boolVitme").ToString() == "0")
                    chBvitme.Checked = false;
                else
                    chBvitme.Checked = true;
            }
            else chBvitme.Checked = true;
            if (key.GetValue("boolSendData") != null)
            {
                if (key.GetValue("boolSendData").ToString() == "0")
                    chkBoxSendITForm.Checked = false;
                else
                    chkBoxSendITForm.Checked = true;
            }
            else chkBoxSendITForm.Checked = true;
            if (key.GetValue("ComPort") != null)
            {
                txtCOM.SelectedIndex = txtCOM.Items.IndexOf(key.GetValue("ComPort"));
                //P.PortName = txtCOM.SelectedItem.ToString();
            }
            else txtCOM.SelectedIndex = 0;
            if (key.GetValue("ComPort1") != null)
            {
                cbCOMport.SelectedIndex = txtCOM.Items.IndexOf(key.GetValue("ComPort1"));
                //P1.PortName = cbCOMport.SelectedItem.ToString();
            }
            else cbCOMport.SelectedIndex = 0;
            if (key.GetValue("RulerDefault") != null)
            {
                txtDefault.Text = key.GetValue("RulerDefault").ToString();
            }
            if (key.GetValue("strConnection") != null)
            {
                txtConnection.Text = key.GetValue("strConnection").ToString();
            }
            if (key.GetValue("strCAMPro") != null)
            {
                txtCameraP.Text = key.GetValue("strCAMPro").ToString();
            }
            if (key.GetValue("MichaelLeo") != null)
            {
                txtConst.Text = key.GetValue("MichaelLeo").ToString();
            }
            txtBaudRate.SelectedIndex = 3;
            txtDataBits.SelectedIndex = 2;
            txtParity.SelectedIndex = 0;
            txtStopBit.SelectedIndex = 0;
            btnConnect_Click(sender, e);
            //this.AcceptButton = btnRun;
            rtb_Log.Text += "Press Enter to start counting";
            //display database
            //dtgv.DefaultCellStyle.Font = new Font("Arial", 11);
            //dtgv.DataSource = newSQLite.ExecuteQueryDataTable("select * from bang1");
            //dtgv.Columns[0].Width = 200;
            //dtgv.Columns[1].Width = 132;
            //dtgv.Columns[2].Width = 132;
            //btnDB.Visible = false;
            /// check Ruler default
            if (boolcheckRuler)
            {
                if (txtDefault.Text == "" || txtDefault.Text == null)
                {
                    boolcheckRuler = false;
                    MessageBox.Show("ruler=" + txtDefault.Text);
                    MessageBox.Show("Rỗng! Hãy cài đặt gốc cho thước đo!!Gọi: 0389233499-A Quân/0395105561-A Thành");
                    this.btnRun.Text = "Error";
                    btnRun.BackColor = Color.Red;
                }
                else if (!checkRuler())
                {
                    boolcheckRuler = false;
                    MessageBox.Show("Sai số! Hãy cài đặt lại gốc cho thước đo!!Gọi: 0389233499-A Quân/0395105561-A Thành");
                    this.btnRun.Text = "Error";
                    btnRun.BackColor = Color.Red;
                }
            }
        }
        private void DataReceive(object obj, SerialDataReceivedEventArgs e)
        {
            InputData = P.ReadExisting();
            if (InputData != string.Empty)
            {
                SetText(InputData);
            }
        }
        private void DataReceive1(object obj, SerialDataReceivedEventArgs e)
        {
            InputData1 = P1.ReadExisting();
            if (InputData1 != string.Empty)
            {
                this.txtDataReceive.Text += InputData1;
            }
        }

        private void SetText(string text)
        {
            if (this.txtDataReceive.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.txtDataReceive.Text += text;
                DataThick += text;
            }
            //this.txtDataReceive.SelectionStart = this.txtDataReceive.Text.Length;
            //this.txtDataReceive.ScrollToCaret();
        }
        private void txtCOM_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (P.IsOpen)
            {
                P.Close();
            }
            RegistryKey key = Registry.CurrentUser.CreateSubKey(@"HKEY_LOCAL_MACHINE\\SOFTWARE\\Ambit\\CMTestProgram\\");
            key.SetValue("ComPort", txtCOM.SelectedItem.ToString());
            key.Close();
            P.PortName = txtCOM.SelectedItem.ToString();
        }

        private void txtBaudRate_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (P.IsOpen)
            {
                P.Close();
            }
            P.BaudRate = Convert.ToInt32(txtBaudRate.Text);
        }

        private void txtDataBits_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (P.IsOpen)
            {
                P.Close();
            }
            P.DataBits = Convert.ToInt32(txtDataBits.Text);
        }

        private void txtParity_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (P.IsOpen)
            {
                P.Close();
            }
            switch (txtParity.SelectedItem.ToString())
            {
                case "Odd":
                    P.Parity = Parity.Odd;
                    break;
                case "None":
                    P.Parity = Parity.None;
                    break;
                case "Even":
                    P.Parity = Parity.Even;
                    break;
            }
        }

        private void txtStopBit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (P.IsOpen)
            {
                P.Close();
            }
            switch (txtStopBit.SelectedItem.ToString())
            {
                case "1":
                    P.StopBits = StopBits.One;
                    break;
                case "1.5":
                    P.StopBits = StopBits.OnePointFive;
                    break;
                case "2":
                    P.StopBits = StopBits.Two;
                    break;
            }
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (P.IsOpen)
                {
                    btnConnect.Enabled = false;
                    btnDisconnect.Enabled = true;
                    toolStripStatusLabel1.Text = " Connected to " + txtCOM.SelectedItem.ToString();
                }
                else
                {
                    btnConnect.Enabled = false;
                    btnDisconnect.Enabled = true;
                    P.PortName = txtCOM.SelectedItem.ToString();
                    toolStripStatusLabel1.Text = "Connected to " + txtCOM.SelectedItem.ToString();
                    P.Open();
                }
            }

            catch (Exception)
            {
                MessageBox.Show("Not Connected", "Thử lại", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try
            {
                if (P1.IsOpen)
                {
                    toolStripStatusLabel1.Text = "Ready";
                    return;
                }
                toolStripStatusLabel1.Text += ", Connected to " + cbCOMport.SelectedItem.ToString();
                //P1 port
                P1.BaudRate = 9600;
                P1.DataBits = 8;
                P1.Parity = System.IO.Ports.Parity.None;
                P1.StopBits = System.IO.Ports.StopBits.One;
                P1.PortName = cbCOMport.SelectedItem.ToString();
                P1.Open();
                //txtDataReceive.Text += "Conected to motor";
                //toolStripStatusLabel2.Text += ",Connected to " + cbCOMport.SelectedItem.ToString();
            }
            catch
            {
                MessageBox.Show("Cant connect to this port, please check again!");
            }
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            P.Close();
            P1.Close();
            btnConnect.Enabled = true;
            btnDisconnect.Enabled = false;
            toolStripStatusLabel1.Text = "Disconnected " + txtCOM.SelectedItem.ToString() + " Port";
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtDataReceive.Text = "";
            //txtDataSend.Text = "";
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            DialogResult Review = MessageBox.Show("Thoát chương trình?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (Review == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void tmSystem_Tick(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            if (P.IsOpen)
            {
                if (txtDataSend.Text == "")
                    MessageBox.Show("Lệnh gửi trống..!", "Information");
                else
                {
                    P.Write(txtDataSend.Text + "\r\n");
                }
            }
            else
                MessageBox.Show("Please check to open COM port!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private void SendToCom(string data)
        {
            if (P.IsOpen)
            {
                if (data == "")
                    MessageBox.Show("Lệnh gửi trống..!", "Information");
                else
                {
                    P.Write(data + "\r\n");
                }
            }
            else;
                //MessageBox.Show("Please check to open COM port!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private bool SendToComNew(SerialPort port, string dataS, string dataR, int delay)
        {
            txtDataReceive.Clear();
            if (port.IsOpen)
            {
                if (dataS == "")
                {
                    MessageBox.Show("Lệnh gửi trống..!", "Information");
                    return false;
                }
                else
                {
                    rtb_Log.Text += Environment.NewLine + "Send cmd to COM:" + dataS + " >>OK";
                    port.Write(dataS + "\r\n");
                    for (int i = 0; i < delay * 10; i++)
                    {
                        DelayMs(0, 100);
                        if (dataR == "")
                        {
                            return true;
                        }
                        if (txtDataReceive.Text.Contains(dataR))
                        {

                            rtb_Log.Text += Environment.NewLine + "COM reponse:" + txtDataReceive.Text;
                            return true;
                        }
                    }
                    rtb_Log.Text += Environment.NewLine + "COM reponse:" + txtDataReceive.Text + " >>No data";
                    return false;
                }
            }
            else
            {
                MessageBox.Show("Please check to open COM port!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private void SendToComMotor(string data)
        {
            if (P1.IsOpen)
            {
                if (data == "")
                    MessageBox.Show("Lệnh gửi trống..!", "Information");
                else
                {
                    P1.Write(data + "\r\n");
                }
            }
            else
                MessageBox.Show("Please check to open COM port!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            //this.btnRun.Enabled = false;
            this.btnRun.Text = "Running";
            btnRun.BackColor = Color.Yellow;
            rtb_Log.Clear();
            piBoxPN.Image = null;
            if (!openCAM())
            {
                return;
            }
            SendToComMotor("RUN1");
            piBoxPN.Load(pathcurrentNow + @"\cam2.PNG");
            if (!getDBValue())
            {
                MessageBox.Show("Chưa có dữ liệu, tìm thông tin thêm mới nhé!");
                btnRun.BackColor = Color.ForestGreen;
                this.btnRun.Text = "Ready";
                SendToComMotor("STOP");
                running = false;
                return;
            }
            DelayMs(0, 8000);
            List<double> listA = new List<double>();
            for (int i = 0; i < 80; i++)
            {
                SendToCom("S");
                DelayMs(0, 100);
                if (DataThick.Contains("S"))
                {
                    DataThick = DataThick.Substring(DataThick.IndexOf("X") + 1, 13).Trim();
                    listA.Add(Double.Parse(DataThick));
                    txtDataReceive.Clear();
                }
                else continue;
                DataThick = "";
            }
            if (listA.Count != 0)
            {
                double temp2 = PartsCounter2(listA.Min() + 6.5, 30, double.Parse(txtThick.Text), double.Parse(txtPicth.Text));
                rtb_Log.Text += Environment.NewLine + "Min:" + (listA.Min() + 6.5).ToString() + "mm";
                rtb_Log.Text += Environment.NewLine + "Max:" + (listA.Max() + 6.5).ToString() + "mm";
                if (temp2 < 0)
                {
                    temp2 = 0;
                }
                //rtb_Log.Text += Environment.NewLine + "Số lượng liệu thực:" + Convert.ToInt32(temp2).ToString();
                //Bu lieu theo gia tri khao sat
                temp2 = BuLieu(listA.Min(), double.Parse(txtThick.Text), temp2);
                rtb_Log.Text += Environment.NewLine + "Số lượng liệu là:" + Convert.ToInt32(temp2).ToString();
                txtQuantity.Text = Convert.ToInt32(temp2).ToString();
                SendToComMotor("STOP");
                DelayMs(0, 500);
            }
            else MessageBox.Show("Kiểm tra kết nối RS232");
            // string pathcurrentNow = Path.GetDirectoryName(Application.ExecutablePath);
            //string filename = DateTime.Today.ToString();
            // rtb_Log.SaveFile(pathcurrentNow+ filename);      
            //this.btnRun.Enabled = true;       
            btnRun.BackColor = Color.ForestGreen;
            this.btnRun.Text = "Finish";
            running = false;
        }
        private void Set_Default()
        {
            rtb_Log.Clear();
            piBoxPN.Image = null;
            btnRun.BackColor = Color.ForestGreen;
            this.btnRun.Text = "Ready";
            running = false;
        }
        //use this method to run
        private void Run_Count()
        {
            txtQuantity.Text = "";
            //this.btnRun.Enabled = false;
            this.btnRun.Text = "Running";
            btnRun.BackColor = Color.Yellow;
            rtb_Log.Clear();
            piBoxPN.Image = null;
            if (!openCAM())
            {
                Set_Default();
                return;
            }
            piBoxPN.Image = Image.FromFile(pathcurrentNow + @"/cam2.PNG");
            if (!getDBValue())
            {
                MessageBox.Show("Chưa có dữ liệu nhé!" + DataSN + ".");
                Set_Default();
                return;
            }
            if (!SendToComNew(P1, "START", "START_OK", 30))
            {
                MessageBox.Show("Check COM");
                Set_Default();
                return;
            }
            DelayMs(0, 300);
            SendToComNew(P1, "RUN1", "", 1);
            List<double> listA = new List<double>();
            for (int i = 0; i < 135; i++)
            {
                SendToCom("S");
                DelayMs(0, 100);
                if (DataThick.Contains("S"))
                {
                    DataThick = DataThick.Substring(DataThick.IndexOf("X") + 1, 13).Trim();
                    listA.Add(Double.Parse(DataThick));
                    txtDataReceive.Clear();
                }
                else continue;
                DataThick = "";
            }
            if (listA.Count != 0)
            {
                double temp2 = PartsCounter2(Math.Abs(listA.Average()) + 6.5, 30 + double.Parse(txtThick.Text) * double.Parse(txtConst.Text), double.Parse(txtThick.Text), double.Parse(txtPicth.Text));
                rtb_Log.Text += Environment.NewLine + "Min:" + (listA.Min() + 6.5).ToString() + "mm";
                rtb_Log.Text += Environment.NewLine + "Max:" + (listA.Max() + 6.5).ToString() + "mm";
                rtb_Log.Text += Environment.NewLine + "SL:" + Convert.ToInt32(temp2).ToString();
                /* if (Int32.Parse(DataMax) > 5000)
                 {
                     if (temp2 <= 7000 && temp2 > 1500)
                         temp2 = temp2 - 100.00;
                     else if (temp2 <= 1500)
                         temp2 = temp2 - 150.00;
                 }
                 if (Int32.Parse(DataMax) < 5001)
                 {
                     if (temp2 < 2500)
                         temp2 = temp2 - 35.00;
                 }*/
                rtb_Log.Text += Environment.NewLine + "SL change:" + Convert.ToInt32(temp2).ToString();
                if (temp2 < 0)
                {
                    temp2 = 0;
                }
                if(temp2 > Int32.Parse(DataMax))
                {
                    temp2 = Int32.Parse(DataMax);
                }
                //Bu lieu theo gia tri khao sat
                //temp2 = BuLieu(listA.Min()+6.5, double.Parse(txtThick.Text), temp2);
                rtb_Log.Text += Environment.NewLine + "Số lượng liệu là:" + Convert.ToInt32(temp2).ToString();
                txtQuantity.Text = Convert.ToInt32(temp2).ToString();
                SendToComNew(P1, "STOP", "RUN_OK", 2);
            }
            else
            {
                SendToComNew(P1, "RST", "RST_OK", 20);
                running = false;
                MessageBox.Show("Kiểm tra kết nối RS232");
            }
            Savelog();
            //try
            //{
            //    Savelog();
            //    PrintResult();
            //}
            //catch
            //{
            //    MessageBox.Show("Please check printer!!");
            //}
            btnRun.BackColor = Color.ForestGreen;
            this.btnRun.Text = "Finish";
            running = false;
            SendToComNew(P1, "RST", "RST_OK", 20);
            if (chkBoxSendITForm.Checked)
            {
                SendDataToITForm();
            }
        }
        private void Run_Count2()
        {
            txtQuantity.Text = "";
            DialogResult rv = MessageBox.Show("Hãy đưa đĩa liệu vào máy đo!", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (rv == DialogResult.No)
            {
                running = false;
                return;
            }
            //this.btnRun.Enabled = false;
            this.btnRun.Text = "Running";
            btnRun.BackColor = Color.Yellow;
            rtb_Log.Clear();
            piBoxPN.Image = null;
            if (!openCAM())
            {
                Set_Default();
                return;
            }
            piBoxPN.Image = Image.FromFile(pathcurrentNow + @"/cam2.PNG");
            if (!getDBValue())
            {
                MessageBox.Show("Kiểm tra lại database");
                Set_Default();
                return;
            }
            DelayMs(0, 300);
            SendToComNew(P1, "RUN1", "", 1);
            List<double> listA = new List<double>();
            for (int i = 0; i < 140; i++)
            {
                SendToCom("S");
                DelayMs(0, 100);
                if (DataThick.Contains("S"))
                {
                    DataThick = DataThick.Substring(DataThick.IndexOf("X") + 1, 13).Trim();
                    listA.Add(Double.Parse(DataThick));
                    txtDataReceive.Clear();
                }
                else continue;
                DataThick = "";
            }
            if (listA.Count != 0)
            {
                double temp2 = PartsCounter2(listA.Min() + 6.5, 30, double.Parse(txtThick.Text), double.Parse(txtPicth.Text));
                rtb_Log.Text += Environment.NewLine + "Min:" + (listA.Min() + 6.5).ToString() + "mm";
                rtb_Log.Text += Environment.NewLine + "Max:" + (listA.Max() + 6.5).ToString() + "mm";
                if (temp2 < 0)
                {
                    temp2 = 0;
                }
                //Bu lieu theo gia tri khao sat
                //temp2 = BuLieu(listA.Min() + 6.5, double.Parse(txtThick.Text), temp2);
                rtb_Log.Text += Environment.NewLine + "Số lượng liệu là:" + Convert.ToInt32(temp2).ToString();
                txtQuantity.Text = Convert.ToInt32(temp2).ToString();
                SendToComNew(P1, "STOP", "RUN_OK", 2);
            }
            else
            {
                running = false;
                MessageBox.Show("Kiểm tra kết nối RS232");
            }
            Savelog();
            //try
            //{
            //    Savelog();
            //    PrintResult();
            //}
            //catch
            //{
            //    MessageBox.Show("Please check printer!!");
            //}
            btnRun.BackColor = Color.ForestGreen;
            this.btnRun.Text = "Finish";
            running = false;
        }
        private void Savelog()
        {
            string filename = txtPN.Text + "-" + DateTime.Now.ToString("HHmmss-ddMMyyyy") + ".txt";
            rtb_Log.SaveFile(pathcurrentNow + "//Logs//" + filename);
        }
        private void checkCAM()
        {
            if (File.Exists(pathcurrentNow + "\\Barcode.txt"))
                File.Delete(pathcurrentNow + "\\Barcode.txt");
            var pros = System.Diagnostics.Process.Start(pathcurrentNow + "\\" + txtCameraP.Text.Trim());
            DelayMs(0, 4000);
            pros.Kill();
        }
        private bool openCAM()
        {
            lblmfg1.Text = "";
            txtPrint.Text = "";
            if (File.Exists(pathcurrentNow + "\\Barcode.txt"))
                File.Delete(pathcurrentNow + "\\Barcode.txt");
            var pros = System.Diagnostics.Process.Start(pathcurrentNow + "\\"+ txtCameraP.Text.Trim());
            DelayMs(0, 4000);
            for (int i = 0; i < 100; i++)
            {
                if (File.Exists(pathcurrentNow + "\\Barcode.txt"))
                {
                    pros.Kill();
                    break;
                }
                else DelayMs(0, 100);
            }
            if (!File.Exists(pathcurrentNow + "\\Barcode.txt"))
            {
                pros.Kill();
                SendToComMotor("STOP");
                MessageBox.Show("Kiểm tra kết nối Camera!", "Warning");
                //this.btnRun.Enabled = true;
                this.btnRun.BackColor = Color.ForestGreen;
                this.btnRun.Text = "Ready";
                return false;              
            }
            DelayMs(0, 1000);
            //File.SetAttributes(pathcurrentNow + "\\Barcode.txt", FileAttributes.Hidden);
            StreamReader str = new StreamReader(pathcurrentNow + "\\Barcode.txt");
            string s = str.ReadLine(); 
            if (s != null)
            {
                str.Close();
                rtb_Log.Text += "Scan barcode= " + s;
                if(!s.Contains("1_"))
                {
                    pros.Kill();
                    SendToComMotor("STOP");
                    btnRun.BackColor = Color.ForestGreen;
                    this.btnRun.Text = "Ready";
                    MessageBox.Show("Sai mã PN!");
                    return false;
                }
                if (s.Contains("CJAPAN"))
                {
                    s = s.Substring(0, s.LastIndexOf(","));
                }
                try
                {
                    DataSN = s.Substring(s.IndexOf("P") + 1, s.IndexOf(",") - 3).Trim();
                    if (DataSN.Contains(","))
                    {
                        DataSN = DataSN.Substring(0, DataSN.IndexOf(",") - 1).Trim();
                    }
                    DataMax = (s.Substring(s.IndexOf("Q")));
                    DataMax = DataMax.Substring(1, DataMax.IndexOf(",") - 1).Trim();
                    DataMax = DataMax.Substring(0, DataMax.LastIndexOf("0") + 1);
                    if (DataMax.Contains(","))
                    {
                        DataMax = DataMax.Substring(0, DataSN.IndexOf(",") - 1).Trim();
                    }
                    if (DataMax.Contains(" "))
                    {
                        DataMax = DataMax.Replace(" ", "");
                    }
                    rtb_Log.Text += Environment.NewLine + "PN=" + DataSN;
                    rtb_Log.Text += Environment.NewLine + "Limit=" + DataMax;
                    lblPN1.Text = DataSN;
                    lblQty1.Text = DataMax;
                    string a = s.Substring(s.IndexOf(",S") + 2);
                    //a = a.Substring(0, a.IndexOf(","));
                    if (a.Contains("MURATA"))
                    {
                        a= a.Substring(0,a.Length-2);
                        txtPrint.Text = a.Trim();
                    }
                    lblmfg1.Text = a.Trim();                    
                    DataPrintSN = s.Substring(s.LastIndexOf(",")).Trim();
                }
                catch
                {
                    pros.Kill();
                    SendToComMotor("STOP");
                    btnRun.BackColor = Color.ForestGreen;
                    this.btnRun.Text = "Ready";
                    MessageBox.Show("Sai mã PN!");
                    return false;
                }
                return true;
            }
            else
            {
                SendToComMotor("STOP");
                btnRun.BackColor = Color.ForestGreen;
                this.btnRun.Text = "Ready";
                MessageBox.Show("SCAN hỏng rồi!");
                return false;
            }
            //while ((s = str.ReadLine()) != null)
            //{
            //    if(s!=null)
            //    {
            //        pros.Kill();
            //        break;
            //    }               
            //}
            //string[] lines = File.ReadAllLines("C:\\Users\\V0957326.VN\\Desktop\\AppRS232\\AppRS232\\Barcode.txt");
            ////line = sr.ReadToEnd();
            //if(lines[0]!=null)
            //{
            //    txtDataReceive.Text =lines[0].ToString();
            //    pros.Kill();
            //}           
        }

        private bool openCAM2()
        {
            File.Delete(pathcurrentNow + "\\Barcode.txt");
            var pros = System.Diagnostics.Process.Start(pathcurrentNow + "\\"+ txtCameraP.Text.Trim());
            DelayMs(0, 3000);
            for (int i = 0; i < 300; i++)
            {
                if (File.Exists(pathcurrentNow + "\\Barcode.txt"))
                {
                    pros.Kill();
                    break;
                }
                else DelayMs(0, 100);
            }
            if (!File.Exists(pathcurrentNow + "\\Barcode.txt"))
            {
                pros.Kill();
                MessageBox.Show("Kiểm tra Camera!", "Warning");
                return false;
            }
            DelayMs(0, 1000);
            StreamReader str = new StreamReader(pathcurrentNow + "\\Barcode.txt");
            string s = str.ReadLine();
            if (s != null)
            {
                str.Close();
                rtb_PP.Text = "Scan barcode= " + s;
                try
                {
                    DataSN = s.Substring(s.IndexOf("P") + 1, s.IndexOf(",") - 3).Trim();
                    if (DataSN.Contains(","))
                    {
                        DataSN = DataSN.Substring(0, DataSN.IndexOf(",") - 1).Trim();
                    }
                    DataMax = (s.Substring(s.IndexOf("Q")));
                    DataMax = DataMax.Substring(1, DataMax.IndexOf(",") - 1).Trim();
                    DataMax = DataMax.Substring(0, DataMax.LastIndexOf("0") + 1);
                    rtb_PP.Text += Environment.NewLine + "\r\nPN=" + DataSN;
                    rtb_PP.Text += Environment.NewLine + "\r\nMaxCount=" + DataMax;
                }
                catch
                {
                    pros.Kill();
                    MessageBox.Show("PN incorrect format");
                    return false;
                }
                txtPNGet.Clear();
                txtPicthGet.Clear();
                txtThickGet.Clear();               
                txtPNGet.Text = DataSN.Trim();
                return true;
            }
            else
            {
                MessageBox.Show("SCAN Fail");
                return false;
            }
        }
        private bool openCAM3()
        {
            if (File.Exists(pathcurrentNow + "\\Barcode.txt"))
                File.Delete(pathcurrentNow + "\\Barcode.txt");
            var pros = System.Diagnostics.Process.Start(pathcurrentNow + "\\"+ txtCameraP.Text.Trim());
            DelayMs(0, 4000);
            for (int i = 0; i < 100; i++)
            {
                if (File.Exists(pathcurrentNow + "\\Barcode.txt"))
                {
                    pros.Kill();
                    break;
                }
                else DelayMs(0, 100);
            }
            if (!File.Exists(pathcurrentNow + "\\Barcode.txt"))
            {
                pros.Kill();
                SendToComMotor("STOP");
                MessageBox.Show("Kiểm tra kết nối Camera!", "Warning");
                //this.btnRun.Enabled = true;
                this.btnRun.BackColor = Color.ForestGreen;
                this.btnRun.Text = "Ready";
                return false;
            }
            pros.Kill();
            DelayMs(0, 1000);
            StreamReader str = new StreamReader(pathcurrentNow + "\\Barcode.txt");
            string s = str.ReadLine();
            if (s != null)
            {
                str.Close();
                rtb_Log.Text += "Scan barcode= " + s;
                try
                {
                    DataSN = s.Substring(s.IndexOf("P") + 1, s.IndexOf(",") - 3).Trim();
                    if (DataSN.Contains(","))
                    {
                        DataSN = DataSN.Substring(0, DataSN.IndexOf(",") - 1).Trim();
                    }
                    rtb_Log.Text += Environment.NewLine + "PN=" + DataSN;
                }
                catch
                {
                    pros.Kill();
                    SendToComMotor("STOP");
                    btnRun.BackColor = Color.ForestGreen;
                    this.btnRun.Text = "Ready";
                    MessageBox.Show("Sai mã PN!");
                    return false;
                }
                return true;
            }
            else
            {
                SendToComMotor("STOP");
                btnRun.BackColor = Color.ForestGreen;
                this.btnRun.Text = "Ready";
                MessageBox.Show("SCAN hỏng rồi!");
                return false;
            }
        }

        private int PartsCounter(double R1, double R2, double thick, double picth)
        {
            int Result = 0;
            int N = (int)((R1 - R2) / thick);
            double temp = pi * (R1 + R2);
            Result = (int)(temp * N / picth);
            return Result;
        }
        private double PartsCounter2(double R1, double R2, double thick, double picth)
        {
            double Result = 0;
            Result = (R1 * R1 - R2 * R2) * pi / (thick * picth);
            //Result = Result - (30 + thick) * pi / (2 * picth);
            return Result;
        }

        private void DelayMs(int diff, int time)
        {
            DateTime dt1 = DateTime.Now;
            diff = 0;
            while (diff < time)
            {
                DateTime dt2 = DateTime.Now;
                TimeSpan ts = dt2.Subtract(dt1);
                diff = (int)ts.TotalMilliseconds;
                Application.DoEvents();
            }
        }

        private bool getDBValue()
        {
            try
            {
                if (!newSQLite.checkPN(DataSN))
                {
                    SendToComMotor("STOP");
                    // this.btnRun.Enabled = true;
                    btnRun.BackColor = Color.ForestGreen;
                    this.btnRun.Text = "Ready";
                    return false;
                }
                //DataTable dt = db.selectDB1("select *from bang1 where PN='" + DataSN.Trim() + "'", txtConnection.Text.Trim());
                DataTable dt = newSQLite.ExecuteQueryDataTable("select *from bang1 where PN='" + DataSN.Trim() + "'");
                toolStripStatusLabel1.Text = "Connect Success!";
                txtPN.Text = dt.Rows[0][0].ToString();
                txtPicth.Text = dt.Rows[0][2].ToString();//isup to bang1 columns
                txtThick.Text = dt.Rows[0][1].ToString();//isup to bang1 columns
                if (Double.Parse(dt.Rows[0][1].ToString()) == 0 || txtThick.Text == "")
                {
                    SendToComMotor("STOP");
                    // this.btnRun.Enabled = true;
                    btnRun.BackColor = Color.ForestGreen;
                    btnRun.Text = "Ready";
                    MessageBox.Show("Hãy set thick cho mã PN này!");
                    return false;
                }
                if (Double.Parse(dt.Rows[0][2].ToString()) == 0 || txtPicth.Text == "")
                {
                    SendToComMotor("STOP");
                    // this.btnRun.Enabled = true;
                    btnRun.BackColor = Color.ForestGreen;
                    btnRun.Text = "Ready";
                    MessageBox.Show("Hãy set picth cho mã PN này!");
                    return false;
                }
                return true;
            }
            catch
            {
                SendToComMotor("STOP");
                //this.btnRun.Enabled = true;
                btnRun.BackColor = Color.ForestGreen;
                btnRun.Text = "Ready";
                MessageBox.Show("Get value fail!");
                return false;
            }
        }

        private void cbCOMport_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (P1.IsOpen)
            {
                P1.Close();
            }
            RegistryKey key = Registry.CurrentUser.CreateSubKey(@"HKEY_LOCAL_MACHINE\\SOFTWARE\\Ambit\\CMTestProgram\\");
            key.SetValue("ComPort1", cbCOMport.SelectedItem.ToString());
            key.Close();
            P1.PortName = cbCOMport.SelectedItem.ToString();
            P1.Open();
        }

        private void btnDB_Click_1(object sender, EventArgs e)
        {
            rtb_PP.Clear();
            DialogResult rv = MessageBox.Show("Hãy đưa đĩa liệu vào máy đo!", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (rv == DialogResult.No)
            {
                running = false;
                return;
            }
            if (!openCAM2())
            {
                running = false;
                return;
            }
            if (P1.IsOpen)
            {
                SendToComNew(P1, "START", "START_OK", 30);
            }
            else
            {
                running = false;
                MessageBox.Show("Check COM motor");
                return;
            }
            //DelayMs(0, 2000);
            SendToComMotor("RUN1");
            List<double> listA = new List<double>();
            for (int i = 0; i < 150; i++)
            {
                SendToCom("S");
                DelayMs(0, 100);
                if (DataThick.Contains("S"))
                {
                    DataThick = DataThick.Substring(DataThick.IndexOf("X") + 1, 13).Trim();
                    listA.Add(Double.Parse(DataThick));
                    txtDataReceive.Clear();
                }
                else continue;
                DataThick = "";
            }
            rtb_PP.Text += Environment.NewLine + "Get thick OK";
            double tempThick = 0;
            if (listA.Count != 0)
            {
                tempThick = Math.Abs(listA.Average()) + 6.5;
                tempThick = averageThick(tempThick, 30, Int32.Parse(cbbPicth.Text), Int32.Parse(DataMax));
                tempThick = Math.Round(tempThick, 4);
                rtb_PP.Text += Environment.NewLine + "Picth:" + cbbPicth.Text + "mm";
                rtb_PP.Text += Environment.NewLine + "Max:" + DataMax + "pcs";
                rtb_PP.Text += Environment.NewLine + "Thick:" + tempThick + "mm";
                SendToComNew(P1, "STOP", "RUN_OK", 2);
            }
            else
            {
                running = false;
                SendToComMotor("STOP");
                MessageBox.Show("Kiểm tra kết nối RS232 đến thước đo!");
                return;
            }
            try
            {
                if (newSQLite.checkPN(DataSN))
                {
                    //db.cmdDB1("update bang1 set THICK=" + tempThick + " where PN='" + DataSN + "'", txtConnection.Text.Trim());
                    newSQLite.ExecuteNonQuery("update bang1 set THICK=" + tempThick + " where PN='" + DataSN + "'");
                    rtb_PP.Text += Environment.NewLine + "update bang1 set THICK=" + tempThick + " where PN='" + DataSN + "'";
                    rtb_PP.Text += Environment.NewLine + "UPDATE OK";
                    SendToComNew(P1, "RST", "RST_OK", 30);
                }
                else
                {
                    //db.cmdDB1("insert into bang1 values('" + DataSN + "'," + Int32.Parse(cbbPicth.Text) + "," + tempThick + ")", txtConnection.Text.Trim());
                    newSQLite.ExecuteNonQuery("insert into bang1 values('" + DataSN + "'," + Int32.Parse(cbbPicth.Text) + "," + tempThick + ")");
                    rtb_PP.Text += Environment.NewLine + "insert into bang1 values('" + DataSN + "'," + cbbPicth.Text + "," + tempThick + ")";
                    rtb_PP.Text += Environment.NewLine + "INSERT OK";
                    SendToComNew(P1, "RST", "RST_OK", 30);
                }
                SendToComMotor("STOP");
                running = false;
            }
            catch
            {
                SendToComMotor("STOP");
                running = false;
                MessageBox.Show("Send Data Fail!");
            }
        }
        private void GetThick_noVITME()
        {
            rtb_PP.Clear();
            DialogResult rv = MessageBox.Show("Hãy đưa đĩa liệu vào máy đo!", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (rv == DialogResult.No)
            {
                running = false;
                return;
            }
            if (!openCAM2())
            {
                running = false;
                return;
            }
            //DelayMs(0, 2000);
            SendToComMotor("RUN1");
            List<double> listA = new List<double>();
            for (int i = 0; i < 150; i++)
            {
                SendToCom("S");
                DelayMs(0, 100);
                if (DataThick.Contains("S"))
                {
                    DataThick = DataThick.Substring(DataThick.IndexOf("X") + 1, 13).Trim();
                    listA.Add(Double.Parse(DataThick));
                    txtDataReceive.Clear();
                }
                else continue;
                DataThick = "";
            }
            double tempThick = 0;
            if (listA.Count != 0)
            {
                tempThick = listA.Min() + 6.5;
                tempThick = averageThick(tempThick, 30, double.Parse(txtPicthGet.Text), double.Parse(DataMax));
                tempThick = Math.Round(tempThick, 4);
                rtb_PP.Text += Environment.NewLine + "Min:" + (listA.Min() + 6.5).ToString() + "mm";
                rtb_PP.Text += Environment.NewLine + "Max:" + (listA.Max() + 6.5).ToString() + "mm";
                rtb_PP.Text += Environment.NewLine + "Thick:" + tempThick + "mm";
                SendToComNew(P1, "STOP", "RUN_OK", 2);
            }
            else
            {
                running = false;
                SendToComMotor("STOP");
                MessageBox.Show("Kiểm tra kết nối RS232 đến thước đo!");
                return;
            }
            try
            {
                if (newSQLite.checkPN(DataSN))
                {
                    newSQLite.ExecuteNonQuery("update bang1 set THICK=" + tempThick + " where PN='" + DataSN + "'");
                    rtb_PP.Text += Environment.NewLine + "update bang1 set THICK=" + tempThick + " where PN='" + DataSN + "'";
                    rtb_PP.Text += Environment.NewLine + "INSERT OK";
                }
                else MessageBox.Show("PN not exist please insert by hand.", "Warning");
                SendToComMotor("STOP");
                running = false;
            }
            catch
            {
                SendToComMotor("STOP");
                running = false;
                MessageBox.Show("Send Data Fail!");
            }
        }
        private void btnDefaultRuler_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            List<double> listB = new List<double>();
            for (int i = 0; i < 10; i++)
            {
                SendToCom("S");
                DelayMs(0, 100);
                if (DataThick.Contains("S"))
                {
                    DataThick = DataThick.Substring(DataThick.IndexOf("X") + 1, 13).Trim();
                    listB.Add(Double.Parse(DataThick));
                    txtDataReceive.Clear();
                }
                else continue;
                DataThick = "";
            }
            if (listB.Count != 0)
            {
                rtb_PP.Text += Environment.NewLine + "Value:" + listB.Max() + "mm";
                RegistryKey key = Registry.CurrentUser.CreateSubKey(@"HKEY_LOCAL_MACHINE\\SOFTWARE\\Ambit\\CMTestProgram\\");
                key.SetValue("RulerDefault", listB.Max().ToString());
                txtDefault.Text = listB.Max().ToString();
                key.Close();
            }
            this.Enabled = true;

        }
        private bool checkRuler()
        {           
                double tempR=0;
                List<double> listB = new List<double>();
                for (int i = 0; i < 3; i++)
                {
                    SendToCom("S");
                    DelayMs(0, 100);
                    if (DataThick.Contains("S"))
                    {
                        DataThick = DataThick.Substring(DataThick.IndexOf("X") + 1, 13).Trim();
                        listB.Add(Double.Parse(DataThick));
                        txtDataReceive.Clear();
                    }
                    else continue;
                    DataThick = "";
                }
                if (listB.Count != 0)
                {
                    tempR = listB.Max();
                    if (Math.Abs(Double.Parse(txtDefault.Text) - tempR) < 5)
                    {
                        rtb_Log.Text += Environment.NewLine + "Ruler Now:" + tempR + "mm";
                        rtb_Log.Text += Environment.NewLine + "Ruler Default:" + txtDefault.Text + "mm";
                        return true;
                    }
                    else
                    {
                        rtb_Log.Text += Environment.NewLine + "Ruler Now:" + tempR + "mm";
                        rtb_Log.Text += Environment.NewLine + "Ruler Default:" + txtDefault.Text + "mm";
                        return false;
                    }
                   
                }
            return false;
        }

        private void btnAddParts_Click(object sender, EventArgs e)
        {
            string GetPN;
            double GetPicth;
            double GetThick;
            this.Enabled = false;
            if (txtPNGet.Text == "")
            {
                MessageBox.Show("Please input PN", "Warning");
                txtPNGet.Focus();
                this.Enabled = true;
                return;
            }
            else GetPN = txtPNGet.Text;
            if (txtPicthGet.Text == "")
            {
                MessageBox.Show("Please input Picth", "Warning");
                txtPicthGet.Focus();
                this.Enabled = true;
                return;

            }
            else
            {
                try
                {
                    GetPicth = Double.Parse(txtPicthGet.Text);
                }
                catch
                {
                    MessageBox.Show("Wrong picth value!!");
                    this.Enabled = true;
                    return;
                }
            }

            if (txtThickGet.Text == "")
            {
                MessageBox.Show("Please input Thick", "Warning");
                txtThickGet.Focus();
                this.Enabled = true;
                return;
            }
            else
            {
                try
                {
                    GetThick = Double.Parse(txtThickGet.Text);
                }
                catch
                {
                    MessageBox.Show("Wrong Thick value!!");
                    this.Enabled = true;
                    return;
                }
            }
            if (newSQLite.checkPN(GetPN))
            {
                newSQLite.ExecuteNonQuery("update bang1 set PICTH=" + GetPicth + " ,THICK=" + GetThick + " where PN='" + GetPN + "'");
                MessageBox.Show("Update Success");
            }
            else
            {
                newSQLite.ExecuteNonQuery("insert into bang1 values('" + GetPN + "'," + GetPicth + "," + GetThick + ")");
                MessageBox.Show("Insert Success");

            }
            this.Enabled = true;
        }

        private void Setup_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Alt)
            {
                if (e.KeyCode == Keys.D1)
                {
                    tabControl1.SelectedIndex = 0;
                    e.Handled = true;
                }               
                else if (e.KeyCode == Keys.D2)
                {
                    tabControl1.SelectedIndex = 1;
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.D3)
                {
                    tabControl1.SelectedIndex = 2;
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.F1)
                {
                    txtPNGet.Enabled = true;
                    txtPicthGet.Enabled = true;
                    txtPicthGet.Visible = true;
                    txtThickGet.Enabled = true;
                    btnDefaultRuler.Visible = true;
                    txtConnection.Enabled = true;
                    txtCameraP.Enabled = true;
                    txtConst.Enabled = true;
                }
                else if (e.KeyCode == Keys.S)
                {

                    btnSTOPVM_Click(sender, e);
                    Set_Default();
                    running = false;
                }
                else if(e.KeyCode==Keys.C)
                {
                    checkCAM();
                }
                else if(e.KeyCode == Keys.M)
                {
                    DoMouseClick();
                }
            }
            if (e.KeyCode == Keys.Enter && running == false && tabControl1.SelectedIndex == 0)
            {
                running = true;
                if (boolcheckRuler)
                {
                    if (chBvitme.Checked)
                    {
                        Run_Count();
                    }
                    else
                        Run_Count2();
                    //Run_Count_ByPoint();
                }
            }
            if (e.KeyCode == Keys.Enter && running == false && tabControl1.SelectedIndex == 1)
            {
                running = true;
                if (chBvitme.Checked)
                {
                    btnDB_Click_1(sender, e);
                }
                else
                    GetThick_noVITME();
            }
            if (e.KeyCode == Keys.L)
            {
                Savelog();
            }
        }
        private double BuLieu(double R1, double thick, double qty)
        {
            double temp = 0;
            if (thick < 1.5)
            {
                if (R1 < 55)
                {
                    temp = qty * 0.93;
                }
                else if (R1 < 65 && R1 >= 55)
                {
                    temp = qty * 0.96;
                }
                else if (R1 <= 75 && R1 > 65)
                {
                    temp = qty * 0.98;
                }
                else
                    temp = qty;
            }
            else
            {
                temp = qty;
            }
            return temp;
        }
       
        private void PrintResult()
        {
            txtPrint.Text = "P" + txtPN.Text + "\r\n" + "Q" + txtQuantity.Text + "\r\n" + DateTime.Now.ToShortDateString();
            if (File.Exists(pathcurrentNow + @"\printFile.txt"))
            {
                File.Delete(pathcurrentNow + @"\printFile.txt");
                File.WriteAllText(pathcurrentNow + @"\printFile.txt", txtPrint.Text);
                string a = File.ReadAllText((pathcurrentNow + @"\printFile.txt"));
            }
            else
            {
                File.WriteAllText(pathcurrentNow + @"\printFile.txt", txtPrint.Text);
            }
            PrintDocument printDoc = new PrintDocument();
            printDoc.PrintPage += new PrintPageEventHandler(pd_PrintPage);
            printDoc.DocumentName = pathcurrentNow + @"\printFile.txt";
            File.SetAttributes(pathcurrentNow + @"\printFile.txt", FileAttributes.Hidden);
            printDoc.Print();
        }
        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            Font f = new Font("Arial", 10);
            string a = File.ReadAllText((pathcurrentNow + @"\printFile.txt"));
            ev.Graphics.DrawString(a, f, Brushes.Black, 15, 10, new StringFormat());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (txtPNGet.Text == "")
            {
                MessageBox.Show("Please input PN", "Warning");
                txtPNGet.Focus();
                this.Enabled = true;
                return;
            }
            txtPicthGet.Clear();
            txtThickGet.Clear();
            try
            {
                if (!newSQLite.checkPN(txtPNGet.Text))
                {
                    MessageBox.Show("NO data please add new");
                }
                else
                {
                    DataTable dt = newSQLite.ExecuteQueryDataTable("select *from bang1 where PN='" + txtPNGet.Text.Trim() + "'");
                    txtPNGet.Text = dt.Rows[0][0].ToString();
                    txtPicthGet.Text = dt.Rows[0][2].ToString();
                    txtThickGet.Text = dt.Rows[0][1].ToString();
                    if (Double.Parse(dt.Rows[0][1].ToString()) == 0)
                    {
                        MessageBox.Show("Please Set the thick!");
                    }
                }

            }
            catch
            {
                MessageBox.Show("Get value fail!");
            }
        }
        private double averageThick(double R1, double R2, double Picth, double Max)
        {
            return (R1 * R1 - R2 * R2) * pi / ((Max+Max/200) * Picth);
        }
        private void Run_Count_ByPoint()
        {
            //this.btnRun.Enabled = false;
            this.btnRun.Text = "Running";
            btnRun.BackColor = Color.Yellow;
            rtb_Log.Clear();
            piBoxPN.Image = null;
            if (!openCAM())
            {
                Set_Default();
                return;
            }
            piBoxPN.Image = Image.FromFile(pathcurrentNow + @"/cam2.PNG");
            if (!getDBValue())
            {
                MessageBox.Show("Check database");
                Set_Default();
                return;
            }
            List<double> listA = new List<double>();
            for (int v = 0; v < 12; v++)
            {
                SendToComNew(P1, "START", "START_OK", 30);
                for (int i = 0; i < 10; i++)
                {
                    SendToCom("S");
                    DelayMs(0, 100);
                    if (DataThick.Contains("S"))
                    {
                        DataThick = DataThick.Substring(DataThick.IndexOf("X") + 1, 13).Trim();
                        listA.Add(Double.Parse(DataThick));
                        txtDataReceive.Clear();
                    }
                    else continue;
                    DataThick = "";
                }
                SendToComNew(P1, "RST", "RST_OK", 20);
            }
            if (listA.Count != 0)
            {
                double temp2 = PartsCounter2(listA.Average() + 6.5, 30, double.Parse(txtThick.Text), double.Parse(txtPicth.Text));
                rtb_Log.Text += Environment.NewLine + "Min:" + (listA.Min() + 6.5).ToString() + "mm";
                rtb_Log.Text += Environment.NewLine + "Max:" + (listA.Max() + 6.5).ToString() + "mm";
                if (temp2 < 0)
                {
                    temp2 = 0;
                }            
                rtb_Log.Text += Environment.NewLine + "Số lượng liệu là:" + Convert.ToInt32(temp2).ToString();
                txtQuantity.Text = Convert.ToInt32(temp2).ToString();
                SendToComNew(P1, "STOP", "RUN_OK", 2);
            }
            else
            {
                SendToComNew(P1, "RST", "RST_OK", 20);
                running = false;
                MessageBox.Show("Kiểm tra kết nối RS232");
            }
            Savelog();
            //try
            //{
            //    Savelog();
            //    PrintResult();
            //}
            //catch
            //{
            //    MessageBox.Show("Please check printer!!");
            //}
            btnRun.BackColor = Color.ForestGreen;
            this.btnRun.Text = "Finish";
            running = false;
            SendToComNew(P1, "RST", "RST_OK", 20);
        }
        private void Get_Thick_byPoint()
        {
            rtb_PP.Clear();
            DialogResult rv = MessageBox.Show("Hãy đưa đĩa liệu vào máy đo!", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (rv == DialogResult.No)
            {
                running = false;
                return;
            }
            if (!openCAM2())
            {
                running = false;
                return;
            }
            List<double> listA = new List<double>();
            for (int v = 0; v < 8; v++)
            {
                SendToComNew(P1, "START", "START_OK", 30);
                for (int i = 0; i < 10; i++)
                {
                    SendToCom("S");
                    DelayMs(0, 100);
                    if (DataThick.Contains("S"))
                    {
                        DataThick = DataThick.Substring(DataThick.IndexOf("X") + 1, 13).Trim();
                        listA.Add(Double.Parse(DataThick));
                        txtDataReceive.Clear();
                    }
                    else continue;
                    DataThick = "";
                }
                SendToComNew(P1, "RST", "RST_OK", 20);
            }
            double tempThick = 0;
            if (listA.Count != 0)
            {
                tempThick = listA.Average() + 6.5;
                tempThick = averageThick(tempThick, 30, double.Parse(txtPicthGet.Text), double.Parse(DataMax));
                tempThick = Math.Round(tempThick, 4);
                rtb_PP.Text += Environment.NewLine + "Min:" + (listA.Min() + 6.5).ToString() + "mm";
                rtb_PP.Text += Environment.NewLine + "Max:" + (listA.Max() + 6.5).ToString() + "mm";
                rtb_PP.Text += Environment.NewLine + "Thick:" + tempThick + "mm";
                SendToComNew(P1, "STOP", "RUN_OK", 2);
            }
            else
            {
                running = false;
                SendToComMotor("STOP");
                MessageBox.Show("Kiểm tra kết nối RS232 đến thước đo!");
                return;
            }
            try
            {
                if (newSQLite.checkPN(DataSN))
                {
                    //MessageBox.Show((listA.Max() - Double.Parse(txtDefault.Text) - 0.05).ToString());
                    newSQLite.ExecuteNonQuery("update bang1 set THICK=" + tempThick + " where PN='" + DataSN + "'");
                    rtb_PP.Text += Environment.NewLine + "update bang1 set THICK=" + tempThick + " where PN='" + DataSN + "'";
                    MessageBox.Show("INSERT OK", "Success");
                    SendToComNew(P1, "RST", "RST_OK", 30);
                }
                else MessageBox.Show("PN not exist please insert by hand.", "Warning");
                SendToComMotor("STOP");
                running = false;
            }
            catch
            {
                SendToComMotor("STOP");
                running = false;
                MessageBox.Show("Send Data Fail!");
            }
        }

        private void btnSTOPVM_Click(object sender, EventArgs e)
        {
            SendToComMotor("STOP");
        }

        private void chBvitme_CheckedChanged(object sender, EventArgs e)
        {
            RegistryKey key = Registry.CurrentUser.CreateSubKey(@"HKEY_LOCAL_MACHINE\\SOFTWARE\\Ambit\\CMTestProgram\\");
            if (chBvitme.Checked)
            {
                key.SetValue("boolVitme", "1");
            }
            else
            {
                key.SetValue("boolVitme", "0");
            }
            key.Close();
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
                RegistryKey key = Registry.CurrentUser.CreateSubKey(@"HKEY_LOCAL_MACHINE\\SOFTWARE\\Ambit\\CMTestProgram\\");
                key.SetValue("strConnection", txtConnection.Text);
                key.SetValue("strCAMPro", txtCameraP.Text);
                key.SetValue("MichaelLeo", txtConst.Text);
                key.Close();
                txtCameraP.Enabled = false;
                txtConnection.Enabled = false;
                txtConst.Enabled = false;          
        }
        //use to find window handle
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string sClassName, String sAppName);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter,string sClassName,string lpsizeWindow);
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        /*private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll",ExactSpelling =true,CharSet =CharSet.Auto)]
        [return:MarshalAs(UnmanagedType.Bool)]*/
        //[DllImport("user32.dll", SetLastError = true)]
        //public static extern Int32 SendMesssage(int hwnd, int Msg, int wparam, StringBuilder lparam);
        // private const int WM_GETTEXT = 0x000D;
        private bool EnumWindow(IntPtr hwnd,IntPtr lParam)
        {
            GCHandle gcChildHandleList = GCHandle.FromIntPtr(lParam);
            if (gcChildHandleList == null || gcChildHandleList.Target == null)
            {
                return false;
            }
            List<IntPtr> childHandle = gcChildHandleList.Target as List<IntPtr>;
            childHandle.Add(hwnd);
            return true;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SendDataToITForm();
                /*txtPrint.Text = "P" + txtPN.Text + "\r\n" + "Q" + txtQuantity.Text + "\r\n" + DateTime.Now.ToShortDateString();
                if (File.Exists(pathcurrentNow + @"\printFile.txt"))
                {
                    File.Delete(pathcurrentNow + @"\printFile.txt");
                    File.WriteAllText(pathcurrentNow + @"\printFile.txt", txtPrint.Text);
                    string a = File.ReadAllText((pathcurrentNow + @"\printFile.txt"));
                }
                else
                {
                    File.WriteAllText(pathcurrentNow + @"\printFile.txt", txtPrint.Text);
                }
                PrintDocument printDoc = new PrintDocument();
                printDoc.PrintPage += new PrintPageEventHandler(pd_PrintPage);
                printDoc.DocumentName = pathcurrentNow + @"\printFile.txt";
                File.SetAttributes(pathcurrentNow + @"\printFile.txt", FileAttributes.Hidden);
                printDoc.Print();*/
        }
        static IntPtr FindWindowByIndex(IntPtr hwndParent,int index ,string sClass)
        {
            if (index == 0)
                return hwndParent;
            else
            {
                int ct = 0;
                IntPtr result = IntPtr.Zero;
                do
                {
                    result = FindWindowEx(hwndParent, result, sClass, null);
                    if (result != IntPtr.Zero)
                    {
                        ++ct;
                    }
                } while (ct < index && result != IntPtr.Zero);
                return result;
            }
        }

        private void chkBoxSendITForm_CheckedChanged(object sender, EventArgs e)
        {
            RegistryKey key = Registry.CurrentUser.CreateSubKey(@"HKEY_LOCAL_MACHINE\\SOFTWARE\\Ambit\\CMTestProgram\\");
            if (chkBoxSendITForm.Checked)
            {
                key.SetValue("boolSendData", "1");
            }
            else
            {
                key.SetValue("boolSendData", "0");
            }
            key.Close();
        }
        private void SendDataToITForm()
        {
            if (txtQuantity.Text.Trim() == "" || txtQuantity.Text.Trim() == "0")
            {
                MessageBox.Show("Null!");
                return;
            }
            IntPtr thisWindow = FindWindow(textBox1.Text.Trim(), textBox2.Text.Trim());
            
            if (!thisWindow.Equals(IntPtr.Zero))
            {               
                {              
                    //IntPtr edithWnd = FindWindowEx(thisWindow, IntPtr.Zero, "TPanel", null);
                    IntPtr edithWnd = FindWindowByIndex(thisWindow, 3, "ThunderRT6Frame");
                    if (!edithWnd.Equals(IntPtr.Zero))
                    {
                        DialogResult rv = MessageBox.Show("Send key to IT system!!", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        SetForegroundWindow(thisWindow);
                        IntPtr edithWnd1 = FindWindowByIndex(edithWnd, 3, "ThunderRT6TextBox");                        
                        if (!edithWnd1.Equals(IntPtr.Zero))
                        {
                            if (rv == DialogResult.Yes)
                            {                                                       
                                //SetForegroundWindow(edithWnd1);
                                //SendKeys.Send("{END}");
                                //for (int k = 1; k < 23; k++)
                                //{
                                //    SendKeys.Send("{BS}");
                                //}
                                //SendMessage(edithWnd1, WM_SETTEXT, IntPtr.Zero, new StringBuilder(lblmfg1.Text));
                                SendKeys.Send(lblmfg1.Text);// send SN PKG  lblmfg1.Text
                                //SendKeys.Send("VTW00110572112025344");
                                SendKeys.Send("{ENTER}");
                                //SendKeys.Send("{END}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Không tìm thấy phần mềm in label1!");
                        }
                        IntPtr edithWnd2 = FindWindowByIndex(edithWnd, 2, "ThunderRT6TextBox");
                        if (!edithWnd2.Equals(IntPtr.Zero))
                        {
                            if (rv == DialogResult.Yes)
                            {
                                //        for (int k = 1; k < 7; k++)
                                //        {
                                            SendKeys.Send("{BS}");
                                //        }
                                //        SetForegroundWindow(edithWnd2);
                                        SendKeys.Send(txtQuantity.Text.Trim());// send quantity txtQuantity.Text
                                //        //SendMessage(edithWnd2, WM_SETTEXT, IntPtr.Zero, new StringBuilder(txtQuantity.Text.Trim()));
                            }
                        }
                        else
                        {
                            MessageBox.Show("Không tìm thấy phần mềm in label2!");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Không tìm thấy phần mềm in label1!");
                    }
                }
            }
            else
            {
                MessageBox.Show("Không tìm thấy phần mềm in label!");
            }
        }
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx,uint dy, uint CButtons,uint dwExtraInfo);
        //Mouse actions
        private const int MOUSEEVENT_LEFTDOWN = 0x02;
        private const int MOUSEEVENT_LEFTUP = 0x04;
        private const int MOUSEEVENT_RIGHTDOWN = 0x08;
        private const int MOUSEEVENT_RIGHTUP = 0x10;
        [DllImport("user32.dll")]
        private static extern Int32 SendMessage(IntPtr hWnd,int Msg,IntPtr wParam,StringBuilder lParam);
        private const int WM_SETTEXT = 0x0C;
        private const int WM_KEYDOWN = 0x100;
        private const int WM_KEYUP = 0x101;
        private const int WM_CHAR = 0x102;
        private const int WM_ENTER = 0x0D;
        public void DoMouseClick()
        {
            IntPtr thisWindow = FindWindow("TAssignMent", "FOXCONN RMA CABLE MODEM PROGRAM V1.5.6 (Build Time: Feb 21 2022 16:54:49)");

            if (!thisWindow.Equals(IntPtr.Zero))
            {
                SetForegroundWindow(thisWindow);
                IntPtr edithWnd = FindWindowByIndex(thisWindow, 1, "TPanel");
                if (!edithWnd.Equals(IntPtr.Zero))
                {
                    IntPtr edithWnd1 = FindWindowByIndex(edithWnd, 1, "TGroupBox");
                    if (!edithWnd1.Equals(IntPtr.Zero))
                    {
                        IntPtr edithWnd2 = FindWindowByIndex(edithWnd1, 5, "TEdit");
                        if (!edithWnd2.Equals(IntPtr.Zero))
                        {
                            SetForegroundWindow(edithWnd2);
                            SendMessage(edithWnd2, WM_SETTEXT, IntPtr.Zero, new StringBuilder("1234"));                           
                            SendMessage(edithWnd2, WM_KEYDOWN, (IntPtr)Keys.End, null);
                            SendMessage(edithWnd2, WM_CHAR, (IntPtr)Keys.Enter, null);
                            IntPtr edithWnd3 = FindWindowByIndex(edithWnd1, 4, "TEdit");
                            SendMessage(edithWnd3, WM_SETTEXT, IntPtr.Zero, new StringBuilder("A"));
                            SendMessage(edithWnd3, WM_KEYDOWN, (IntPtr)Keys.End, null);
                            //SendMessage(edithWnd2, WM_KEYDOWN, (IntPtr)Keys.Delete, null);
                            //uint X = (uint)Cursor.Position.X;
                            //uint Y = (uint)Cursor.Position.Y;
                            //mouse_event(MOUSEEVENT_LEFTDOWN, 400, 200, 0, 0);
                            //mouse_event(MOUSEEVENT_LEFTDOWN, 400, 200, 0, 0);
                            //mouse_event(MOUSEEVENT_LEFTDOWN, 400, 200, 0, 0);
                            //mouse_event(MOUSEEVENT_LEFTDOWN, 400, 200, 0, 0);
                        }
                        else
                        {
                            MessageBox.Show("FAIL2");
                        }
                    }
                    else
                    {
                        MessageBox.Show("FAIL1");
                    }
                }
                else
                {
                    MessageBox.Show("FAIL");
                }
            }
            
        }


        //private void button2_Click1(object sender, EventArgs e)
        //{
        //    lblmfg1.Text = "";
        //    txtPrint.Text = "";

        //    //File.SetAttributes(pathcurrentNow + "\\Barcode.txt", FileAttributes.Hidden);
        //    StreamReader str = new StreamReader(pathcurrentNow + "\\Barcode.txt");
        //    string s = str.ReadLine();
        //    if (s != null)
        //    {
        //        str.Close();
        //        rtb_Log.Text += "Scan barcode= " + s;
        //        if (!s.Contains("1_"))
        //        {

        //            SendToComMotor("STOP");
        //            btnRun.BackColor = Color.ForestGreen;
        //            this.btnRun.Text = "Ready";
        //            MessageBox.Show("Sai mã PN !");
        //        }
        //        if (s.Contains("CJAPAN"))
        //        {
        //            s = s.Substring(0, s.LastIndexOf(","));
        //        }
        //        try
        //        {
        //            DataSN = s.Substring(s.IndexOf("P") + 1, s.IndexOf(",") - 3).Trim();
        //            if (DataSN.Contains(","))
        //            {
        //                DataSN = DataSN.Substring(0, DataSN.IndexOf(",") - 1).Trim();
        //            }
        //            DataMax = (s.Substring(s.IndexOf("Q")));
        //            DataMax = DataMax.Substring(1, DataMax.IndexOf(",") - 1).Trim();
        //            DataMax = DataMax.Substring(0, DataMax.LastIndexOf("0") + 1);
        //            if (DataMax.Contains(","))
        //            {
        //                DataMax = DataMax.Substring(0, DataSN.IndexOf(",") - 1).Trim();
        //            }
        //            if (DataMax.Contains(" "))
        //            {
        //                DataMax = DataMax.Replace(" ", "");
        //            }
        //            rtb_Log.Text += Environment.NewLine + "PN=" + DataSN;
        //            rtb_Log.Text += Environment.NewLine + "Limit=" + DataMax;
        //            lblPN1.Text = DataSN;
        //            lblQty1.Text = DataMax;
        //            string a = s.Substring(s.IndexOf(",S") + 2);
        //            if (a.Contains(","))
        //            {
        //                a = a.Substring(0, a.IndexOf(",") - 1).Trim();
        //            }
        //            //a = a.Substring(0, a.IndexOf(","));
        //            if (a.Contains("MURATA"))
        //            {
        //                a = a.Substring(0, a.Length - 2);
        //                txtPrint.Text = a.Trim();
        //            }
        //            lblmfg1.Text = a.Trim();
        //            DataPrintSN = s.Substring(s.LastIndexOf(",")).Trim();
        //            MessageBox.Show(lblmfg1.Text);
        //        }
        //        catch
        //        {
        //            SendToComMotor("STOP");
        //            btnRun.BackColor = Color.ForestGreen;
        //            this.btnRun.Text = "Ready";
        //            MessageBox.Show("Sai mã PN abc!");

        //        }

        //    }
        //    else
        //    {
        //        SendToComMotor("STOP");
        //        btnRun.BackColor = Color.ForestGreen;
        //        this.btnRun.Text = "Ready";
        //        MessageBox.Show("SCAN hỏng rồi!");

        //    }
        //}
    }
}
