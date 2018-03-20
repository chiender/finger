using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Data.OleDb;
using System.IO;
using System.Net.Mail;


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public int i, X, j = 0, node_nu = 1, bytes_count = 0, id_count = 0, flag = 0; //i:迴圈用,j:接收資料陣列用,node_nu:節點旗標,bytes_count:位元計數用
        public byte[] Rx_data = new byte[110]; //rs232接收,預設陣列
        OleDbConnection conn;  
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //MessageBox.Show("bye bye");
            serialPort1.Close();
            conn.Close();         //關閉資料庫
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            // TODO: 這行程式碼會將資料載入 'dBMSDataSet.Person' 資料表。您可以視需要進行移動或移除。
            //this.personTableAdapter.Fill(this.dBMSDataSet.Person);
            System.IO.Ports.SerialPort sport = new System.IO.Ports.SerialPort();//宣告連接埠
            foreach (string com in System.IO.Ports.SerialPort.GetPortNames())//取得所有可用的連接埠
            {
                cBox_port.Items.Add(com);
            }
            tabControl1.Enabled = false;
            
            //tabControl1.TabPages.Remove(tabPage4); //登入
            //tabControl1.TabPages.Remove(tabPage1); //出勤資料
            tabControl1.TabPages.Remove(tabPage2);  //人事資料
            tabControl1.TabPages.Remove(tabPage3);  //系統管理
            //tabControl1.TabPages.Remove(tabPage5);  //說明
            tabControl1.TabPages.Remove(tabPage6);  //測試
            btn_delall_arr.Enabled = false; 
            btn_delallPDB.Enabled = false;
            btn_delallFinger.Enabled = false;
            btn_search_arrivals.Enabled = false;
            btn_excel.Enabled = false;

            //請假管理
            btn_lea_sentmail.Visible= false; //寄信測試
            btn_lea_del.Enabled = false; //刪除
            btn_lea_delall.Enabled = false; //刪除所有
            btn_lea_excel.Enabled = false; //輸出EXCEL
            btn_lea_query.Enabled = false; //查詢

            button1.Visible = false;

            //string dbpath = "DBMS.mdb";    //宣告資料庫所在的路徑變數
            string Source;                 //宣告連線的字串
            //Source = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbpath;
            Source = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBMS.mdb";
            //OleDbConnection conn;          //宣告連線的物件
            conn = new OleDbConnection(Source);   //連線
            conn.Open();          //開啟資料庫
            //MessageBox.Show("成功連結到Access資料庫");
            //conn.Close();         //關閉資料庫
        }

        private void timer1_Tick(object sender, EventArgs e) //用計時器方式快確認指紋是否建檔
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            int ii;
            int jj = 0;
            //int id_count = 0;
            byte[] TX_com = new byte[12];
            //讀取開頭
            TX_com[0] = 85; //0x55
            TX_com[1] = 170; //0xAA

            TX_com[2] = 1;
            TX_com[3] = 0;
            //參數
            //TX_com[4] = byte.Parse(textBox41.Text);
            TX_com[5] = 0;
            TX_com[6] = 0;
            TX_com[7] = 0;

            //命令代碼
            TX_com[8] = 33;
            TX_com[9] = 0;


            TX_com[4] = byte.Parse((id_count +int.Parse(txt_initVaule.Text)).ToString());//設掃瞄起始值

            //check sum (從0累加至9) (jj:BD0A)
            for (ii = 0; ii <= 9; ii++)
            {
                jj = jj + TX_com[ii];

            }
            //將累加值,分別存入陣列
            TX_com[10] = byte.Parse(Convert.ToInt32(Convert.ToString(int.Parse(jj.ToString()), 16).PadLeft(4, '0').Substring(2, 2), 16).ToString());
            TX_com[11] = byte.Parse(Convert.ToInt32(Convert.ToString(int.Parse(jj.ToString()), 16).PadLeft(4, '0').Substring(0, 2), 16).ToString());

            txt_Rx6.Text = TX_com[10].ToString();
            txt_Rx7.Text = TX_com[11].ToString();
            serialPort1.Write(TX_com, 0, 11);
            //Thread.Sleep(500);
            ii = 0;
            jj = 0;
            id_count++;
            if (id_count == (20 + int.Parse(txt_initVaule.Text))) //起始值往後掃20枚
            { 
                timer1.Enabled = false;
            }
        }

        private void btn_connect_Click(object sender, EventArgs e)//RS232連線
        {
            if (serialPort1.IsOpen == false)
            {
                try
                {
                    serialPort1.PortName = cBox_port.Text;
                    serialPort1.BaudRate = Convert.ToUInt16(cBox_BRate.Text);
                    serialPort1.Open();
                    btn_connect.Text = "中斷";
                    btn_connect.BackColor = Color.Red;
                    //btn_clear.Enabled = true;
                    tabControl1.Enabled = true;
                    //Btn_string.Enabled = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else if (serialPort1.IsOpen == true)
            {
                try
                {
                    serialPort1.Close();
                    btn_connect.Text = "連線";
                    btn_connect.BackColor = Color.Lime ;
                    //btn_clear.Enabled = false;
                    btn_logout.PerformClick();
                    tabControl1.Enabled = false;
                    //Btn_string.Enabled = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private string ByteArrayToHexString(byte[] data)//bytes to Hex,無特別數值轉換時,可省略以下
        {
            StringBuilder sb = new StringBuilder(data.Length * 3);
            foreach (byte b in data)
                sb.Append(Convert.ToString(b, 16).PadLeft(2, '0').PadRight(3, ' '));
            return sb.ToString().ToUpper();
        }

        private byte[] HexStringToByteArray(string s)// Hex to bytes,無特別數值轉換時,可省略以下
        {
            s = s.Replace(" ", "");
            byte[] buffer = new byte[s.Length / 2];
            for (int i = 0; i < s.Length; i += 2)
                buffer[i / 2] = (byte)Convert.ToByte(s.Substring(i, 2), 16);
            return buffer;
        }
        
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)//RS232接收料資觸發
        {
            this.Invoke(new EventHandler(RS232Data));
        }

        private void RS232Data(object s, EventArgs e)   //rs232接收資料
        {
            int Rx_count = serialPort1.BytesToRead; //接收多少資料位元組
            byte[] RxBuff = new byte[Rx_count];

            switch (tabControl1.SelectedTab.Text)
            {
                case "主頁面":
                    break;
                case "出勤資料":
                    serialPort1.Read(RxBuff, 0, Rx_count);  //讀取資料

                    for (i = bytes_count; i < (Rx_count + bytes_count); i++)
                    {
                        if (j != Rx_count)
                        {
                            Rx_data[i] = RxBuff[j];
                        }
                        j++;
                    }
                    j = 0;

                    bytes_count = Rx_count + bytes_count;    //計算資料接收位元組

                    if (bytes_count % 12 == 0)//接收資料一行為12 bytes
                    {
                        if ((bytes_count / 12) == 1)//接收一行指令
                        {
                            for (X = 0; X < (bytes_count / 12); X++)
                            {
                                //txt_Rx4.Text = txt_Rx4.Text + Rx_data[8 + (12 * X)].ToString() + "\r\n"; //ACK 48 or NACK 49
                                //txt_Rx2.Text = txt_Rx2.Text + Rx_data[4 + (12 * X)].ToString() + "\r\n"; //接收參數
                                //txt_Rx3.Text = txt_Rx3.Text + Rx_data[5 + (12 * X)].ToString() + "\r\n"; //接收參數
                            }
                            
                            X = 0;
                        }

                        if ((bytes_count / 12) == 4)//接收4行指令,一對多比對,48 bytes
                        {
                            for (X = 0; X < (bytes_count / 12); X++)
                            {
                                //txt_Rx4.Text = txt_Rx4.Text + Rx_data[8 + (12 * X)].ToString() + "\r\n"; //ACK 48 or NACK 49
                                //txt_Rx2.Text = txt_Rx2.Text + Rx_data[4 + (12 * X)].ToString() + "\r\n"; //接收參數
                                //txt_Rx3.Text = txt_Rx3.Text + Rx_data[5 + (12 * X)].ToString() + "\r\n"; //接收參數

                                if (Rx_data[8 + (12 * X)] == 49)
                                {
                                    if (Rx_data[4 + (12 * X)] == 18)
                                    {
                                        //MessageBox.Show("手指沒有按壓");
                                    }
                                    if (Rx_data[4 + (12 * X)] == 8)
                                    {
                                        //MessageBox.Show("比對失敗");
                                    }
                                }
                                if (X == 2)
                                {
                                    if (Rx_data[8 + (12 * X)] == 48)
                                    {
                                        //MessageBox.Show("比對成功!\r\n指紋編號: " + Rx_data[4 + (12 * X)]);
                                        //比對成功,即執行寫入資料庫
                                        string SelectCmd = "INSERT INTO arrivals(編號,年月日,時分秒) VALUES('"
                                                    + Rx_data[4 + (12 * X)].ToString() + "','" + DateTime.Now.ToShortDateString() + "','" +
                                                    DateTime.Now.ToLongTimeString() + "')";
                                        OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
                                        Cmd.ExecuteNonQuery();
                                        btn_show_arrivals.PerformClick();
                                        
                                    }
                                }
                            }
                            X = 0;
                        }
                        
                    }
                    //txt_Rx5.Text = bytes_count.ToString(); //計算bytes
                    break;
                case "人事資料":
                    //int Rx_count = serialPort1.BytesToRead; //接收多少資料位元組
                    //byte[] RxBuff = new byte[Rx_count];

                    serialPort1.Read(RxBuff, 0, Rx_count);  //讀取資料

                    //txt_Rx1.Text = txt_Rx1.Text + ByteArrayToHexString(RxBuff) + "\r\n";   //顯示資料

                    //將片段接收的資料,存入預設的陣列Rx_data 
                    for (i = bytes_count; i < (Rx_count + bytes_count); i++)
                    {
                        if (j != Rx_count)
                        {
                            Rx_data[i] = RxBuff[j];
                        }
                        j++;
                    }
                    j = 0;

                    bytes_count = Rx_count + bytes_count;    //計算資料接收位元組

                    if (bytes_count % 12 == 0)//接收資料一行為12 bytes
                    {
                        if ((bytes_count / 12) == 1)//接收一行指令
                        {
                            for (X = 0; X < (bytes_count / 12); X++)
                            {
                                //txt_Rx4.Text = txt_Rx4.Text + Rx_data[8 + (12 * X)].ToString() + "\r\n"; //ACK 48 or NACK 49
                                //txt_Rx2.Text = txt_Rx2.Text + Rx_data[4 + (12 * X)].ToString() + "\r\n"; //接收參數
                                //txt_Rx3.Text = txt_Rx3.Text + Rx_data[5 + (12 * X)].ToString() + "\r\n"; //接收參數
                            }
                            if (flag == 1)//刪除指定ID 
                            { 
                                if (Rx_data[8] == 48)
                                {
                                    MessageBox.Show("刪除成功!","訊息",MessageBoxButtons .OK ,MessageBoxIcon.Information );
                                    //刪除成功,就刪除資料庫id
                                    string SelectCmd = "DELETE * FROM Person WHERE 編號='"
                                                        + txt_number.Text.Trim() + "'";
                                    OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
                                    Cmd.ExecuteNonQuery();
                                    btn_showAllDB.PerformClick();
                                    txt_number.Text = "";
                                }
                                else
                                {
                                    if (Rx_data[4] == 4)
                                    { //指定的ID沒有被使用
                                        MessageBox.Show("刪除失敗!" + "\r\n" + "指定的ID沒有被使用", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                    else
                                    {
                                        MessageBox.Show("刪除失敗!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }

                                }
                                flag = 0;//預設
                            }
                            if (flag == 2)
                            { //讀取指紋總數
                                if (Rx_data[8] == 48)
                                {
                                    MessageBox.Show("總指數: " + Rx_data[4], "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                flag = 0;//預設
                            }
                            if (flag == 3)
                            { //刪除所有指紋
                                if (Rx_data[8] == 48)
                                {
                                    MessageBox.Show("已全部刪除指紋庫", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information );
                                }
                                else
                                {
                                    MessageBox.Show("刪除失敗!", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                                flag = 0;//預設
                            }
                            X = 0;
                        }

                        if ((bytes_count / 12) == 4)//接收4行指令,一對多比對,48 bytes
                        {
                            for (X = 0; X < (bytes_count / 12); X++)
                            {
                                //txt_Rx4.Text = txt_Rx4.Text + Rx_data[8 + (12 * X)].ToString() + "\r\n"; //ACK 48 or NACK 49
                                //txt_Rx2.Text = txt_Rx2.Text + Rx_data[4 + (12 * X)].ToString() + "\r\n"; //接收參數
                                //txt_Rx3.Text = txt_Rx3.Text + Rx_data[5 + (12 * X)].ToString() + "\r\n"; //接收參數
                                
                                if (Rx_data[8 + (12 * X)] == 49)
                                {
                                    if (Rx_data[4 + (12 * X)] == 18)
                                    {
                                        MessageBox.Show("手指沒有按壓","警告",MessageBoxButtons.OK ,MessageBoxIcon.Warning);
                                    }
                                    if (Rx_data[4 + (12 * X)] == 8)
                                    {
                                        MessageBox.Show("比對失敗", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                }
                                if (X == 2)
                                {
                                    if (Rx_data[8 + (12 * X)] == 48)
                                    {
                                        MessageBox.Show("比對成功!\r\n指紋編號: " + Rx_data[4 + (12 * X)], "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        //比對成功,即執行資料庫查詢
                                        txt_number.Text = Rx_data[4 + (12 * X)].ToString();
                                        btn_searchDB.PerformClick();
                                        txt_number.Text = "";
                                    }
                                }
                            }
                            X = 0;
                        }
                        if ((bytes_count / 12) == 9)//接收9行指令,快速建檔, 108 bytes
                        {
                            int Ack_count = 0;
                            for (X = 0; X < (bytes_count / 12); X++)
                            {
                                //txt_Rx4.Text = txt_Rx4.Text + Rx_data[8 + (12 * X)].ToString() + "\r\n"; //ACK 48 or NACK 49
                                //txt_Rx2.Text = txt_Rx2.Text + Rx_data[4 + (12 * X)].ToString() + "\r\n"; //接收參數
                                //txt_Rx3.Text = txt_Rx3.Text + Rx_data[5 + (12 * X)].ToString() + "\r\n"; //接收參數

                                if (Rx_data[8 + (12 * X)] == 48)
                                {
                                    Ack_count++;
                                }
                            }
                            if (Ack_count == 9)
                            {
                                MessageBox.Show("指紋建檔成功","訊息",MessageBoxButtons.OK ,MessageBoxIcon.Information );
                                //建檔成功,將資料寫入資料庫
                                string SelectCmd = "INSERT INTO Person(編號,姓名,電話,手機,mail,地址,部門,職位,上班時間,下班時間,午休起,午休迄) VALUES('"
                                                    + txt_number.Text.Trim() + "','" + txt_name.Text.Trim() + "','" +
                                                    txt_tel.Text.Trim() + "','" + txt_phone.Text.Trim() + "','" +
                                                     txt_email.Text.Trim()+ "','"+ txt_address.Text.Trim() +"','"+
                                                     com_depart.Text.Trim()+"','"+ com_position.Text.Trim()+"','"+
                                                     com_onwork_h.Text.Trim()+":"+com_onwork_m.Text.Trim()+"','"+
                                                     com_offwork_h.Text.Trim()+":"+com_offwork_m.Text.Trim()+"','"+
                                                     com_break_h1.Text.Trim()+":"+com_break_m1.Text.Trim()+"','"+
                                                     com_break_h2.Text.Trim()+":"+com_break_m2.Text.Trim()+"')";
                                                     
                                OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
                                Cmd.ExecuteNonQuery();
                                btn_showAllDB.PerformClick();
                                txt_number.Text = "";
                                txt_name.Text = "";
                                txt_phone.Text = "";
                                txt_tel.Text = "";
                                txt_address.Text = "";
                                txt_email.Text = "";
                            }
                            else
                            {

                                MessageBox.Show("指紋建檔失敗", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                            }

                            X = 0;
                        }
                    }
                    //txt_Rx5.Text = bytes_count.ToString(); //計算bytes

                    break;
                case "系統管理":
                    break;
                case "說明":
                    break;
                case "測試"://測試資料接收,指紋辨識.
                    //int Rx_count = serialPort1.BytesToRead; //接收多少資料位元組
                    //byte[] RxBuff = new byte[Rx_count];

                    serialPort1.Read(RxBuff, 0, Rx_count);  //讀取資料

                    txt_Rx1.Text = txt_Rx1.Text + ByteArrayToHexString(RxBuff) + "\r\n";   //顯示資料

                    //將片段接收的資料,存入預設的陣列Rx_data 
                    for (i = bytes_count; i < (Rx_count + bytes_count); i++)
                    {
                        if (j != Rx_count)
                        {
                            Rx_data[i] = RxBuff[j];
                        }
                        j++;
                    }
                    j = 0;

                    bytes_count = Rx_count + bytes_count;    //計算資料接收位元組

                    if (bytes_count % 12 == 0)//接收資料一行為12 bytes
                    {
                        if ((bytes_count / 12) == 1)//接收一行指令
                        {
                            for (X = 0; X < (bytes_count / 12); X++)
                            {
                                txt_Rx4.Text = txt_Rx4.Text + Rx_data[8 + (12 * X)].ToString() + "\r\n"; //ACK 48 or NACK 49
                                txt_Rx2.Text = txt_Rx2.Text + Rx_data[4 + (12 * X)].ToString() + "\r\n"; //接收參數
                                txt_Rx3.Text = txt_Rx3.Text + Rx_data[5 + (12 * X)].ToString() + "\r\n"; //接收參數
                            }
                            if (flag == 1)
                            { //刪除指定ID 
                                if (Rx_data[8] == 48)
                                {
                                    MessageBox.Show("刪除成功!");
                                    //刪除成功,就刪除資料庫id
                                }
                                else
                                {
                                    if (Rx_data[4] == 4)
                                    { //指定的ID沒有被使用
                                        MessageBox.Show("刪除失敗!" + "\r\n" + "指定的ID沒有被使用");
                                    }
                                    else
                                    {
                                        MessageBox.Show("刪除失敗!");
                                    }

                                }
                                flag = 0;
                            }
                            if (flag == 2)
                            { //讀取指紋總數
                                if (Rx_data[8] == 48)
                                {
                                    MessageBox.Show("總指數: " + Rx_data[4]);
                                }
                                flag = 0;//預設
                            }
                            X = 0;
                        }

                        if ((bytes_count / 12) == 4)//接收4行指令,一對多比對,48 bytes
                        {
                            for (X = 0; X < (bytes_count / 12); X++)
                            {
                                txt_Rx4.Text = txt_Rx4.Text + Rx_data[8 + (12 * X)].ToString() + "\r\n"; //ACK 48 or NACK 49
                                txt_Rx2.Text = txt_Rx2.Text + Rx_data[4 + (12 * X)].ToString() + "\r\n"; //接收參數
                                txt_Rx3.Text = txt_Rx3.Text + Rx_data[5 + (12 * X)].ToString() + "\r\n"; //接收參數
                                if (Rx_data[8 + (12 * X)] == 49)
                                {
                                    if (Rx_data[4 + (12 * X)] == 18)
                                    {
                                        MessageBox.Show("手指沒有按壓");
                                    }
                                    if (Rx_data[4 + (12 * X)] == 8)
                                    {
                                        MessageBox.Show("比對失敗");
                                    }
                                }
                                if (X == 2)
                                {
                                    if (Rx_data[8 + (12 * X)] == 48)
                                    {
                                        MessageBox.Show("指紋編號: " + Rx_data[4 + (12 * X)]);
                                        //比對成功,即執行資料庫查詢
                                    }
                                }
                            }
                            X = 0;
                        }
                        if ((bytes_count / 12) == 9)//接收9行指令,快速建檔, 108 bytes
                        {
                            int Ack_count = 0;
                            for (X = 0; X < (bytes_count / 12); X++)
                            {
                                txt_Rx4.Text = txt_Rx4.Text + Rx_data[8 + (12 * X)].ToString() + "\r\n"; //ACK 48 or NACK 49
                                txt_Rx2.Text = txt_Rx2.Text + Rx_data[4 + (12 * X)].ToString() + "\r\n"; //接收參數
                                txt_Rx3.Text = txt_Rx3.Text + Rx_data[5 + (12 * X)].ToString() + "\r\n"; //接收參數

                                if (Rx_data[8 + (12 * X)] == 48)
                                {
                                    Ack_count++;
                                }
                            }
                            if (Ack_count == 9)
                            {
                                MessageBox.Show("指紋建檔成功");
                                //建檔成功,將資料寫入資料庫
                            }
                            else
                            {

                                MessageBox.Show("指紋建檔失敗");
                            }

                            X = 0;
                        }
                    }
                    txt_Rx5.Text = bytes_count.ToString(); //計算bytes

                    break;
                default:
                    break;
            }

        }

        private void btn_TX_LEDopen_Click(object sender, EventArgs e) //LED open
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;

            byte[] data = HexStringToByteArray("55 AA 01 00 01 00 00 00 12 00 13 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_LEDclose_Click(object sender, EventArgs e) //LED Close
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 00 00 00 00 12 00 12 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_RX_delall_Click(object sender, EventArgs e)//清空所有接收資料
        {
            txt_Rx1.Text = "";
            txt_Rx2.Text = "";
            txt_Rx3.Text = "";
            txt_Rx4.Text = "";
            txt_Rx5.Text = "";
            txt_Rx6.Text = "";
            txt_Rx7.Text = "";

        }

        private void btn_TX_test_Click(object sender, EventArgs e)//指紋辨識,測試
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 01 00 00 00 01 00 02 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_initi_Click(object sender, EventArgs e) //初始化
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            //改變baud rate 9600
            byte[] data = HexStringToByteArray("55 AA 01 00 80 25 00 00 04 00 A9 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(300);
            //初始化
            data = HexStringToByteArray("55 AA 01 00 01 00 00 00 01 00 02 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_IDnumber_Click(object sender, EventArgs e) //取得建檔枚數
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            flag = 2; //回傳一行接收資料,判斷,2為建檔枚數
            byte[] data = HexStringToByteArray("55 AA 01 00 00 00 00 00 20 00 20 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_impress_Click(object sender, EventArgs e) //確認手指是否有壓
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 00 00 00 00 26 00 26 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_identify_Click(object sender, EventArgs e) //identify 1:N
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 00 00 00 00 51 00 51 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_del2_Click(object sender, EventArgs e) //del id=2
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 02 00 00 00 40 00 42 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_crea2_Click(object sender, EventArgs e)//開始建檔 id=2
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 02 00 00 00 22 00 24 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_capFing_Click(object sender, EventArgs e) //採集指紋
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 00 00 00 00 60 00 60 01");
            serialPort1.Write(data, 0, data.Length);

        }

        private void btn_TX_create1_Click(object sender, EventArgs e) //第一次建檔
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 00 00 00 00 23 00 23 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_create2_Click(object sender, EventArgs e) //第二次建檔
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 00 00 00 00 24 00 24 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_create3_Click(object sender, EventArgs e) //第三次建檔
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 00 00 00 00 25 00 25 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_command_Click(object sender, EventArgs e) //指紋 指令輸出
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] TX_com = new byte[12];
            //讀取開頭
            TX_com[0] = 85; //0x55
            TX_com[1] = 170; //0xAA

            TX_com[2] = 1;
            TX_com[3] = 0;
            //參數
            TX_com[4] = byte.Parse(txt_Tx_var.Text);
            TX_com[5] = 0;
            TX_com[6] = 0;
            TX_com[7] = 0;

            //命令代碼
            TX_com[8] = byte.Parse(txt_Tx_command.Text);
            TX_com[9] = 0;

            //check sum (從0累加至9) (jj:BD0A)
            int ii;
            int jj = 0;
            for (ii = 0; ii <= 9; ii++)
            {
                jj = jj + TX_com[ii];

            }
            //將累加值,分別存入陣列
            TX_com[10] = byte.Parse(Convert.ToInt32(Convert.ToString(int.Parse(jj.ToString()), 16).PadLeft(4, '0').Substring(2, 2), 16).ToString());
            TX_com[11] = byte.Parse(Convert.ToInt32(Convert.ToString(int.Parse(jj.ToString()), 16).PadLeft(4, '0').Substring(0, 2), 16).ToString());

            txt_Rx6.Text = TX_com[10].ToString();
            txt_Rx7.Text = TX_com[11].ToString();

            serialPort1.Write(TX_com, 0, 11);
        }

        private void btn_TX_verify_Click(object sender, EventArgs e) //1對多比對
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            //LED開
            byte[] data = HexStringToByteArray("55 AA 01 00 01 00 00 00 12 00 13 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(200);
            //採集指紋

            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 60 00 60 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(200);
            //1對多比對
            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 51 00 51 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(1000);
            //LED關
            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 12 00 12 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_create_Click(object sender, EventArgs e) //快速建檔
        {
            if (txt_Tx_var.Text  == "") {
                MessageBox.Show("請輸入參數");
                return;
            }
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;

            byte[] TX_com = new byte[12];
            //讀取開頭
            TX_com[0] = 85; //0x55
            TX_com[1] = 170; //0xAA
            TX_com[2] = 1;
            TX_com[3] = 0;
            //參數
            TX_com[4] = byte.Parse(txt_Tx_var.Text);
            TX_com[5] = 0;
            TX_com[6] = 0;
            TX_com[7] = 0;
            //命令代碼,建檔
            TX_com[8] = 34; //byte.Parse(textBox42.Text);
            TX_com[9] = 0;
            //check sum (從0累加至9) (jj:BD0A)
            int ii;
            int jj = 0;
            for (ii = 0; ii <= 9; ii++)
            {
                jj = jj + TX_com[ii];
            }
            //將累加值,分別存入陣列
            TX_com[10] = byte.Parse(Convert.ToInt32(Convert.ToString(int.Parse(jj.ToString()), 16).PadLeft(4, '0').Substring(2, 2), 16).ToString());
            TX_com[11] = byte.Parse(Convert.ToInt32(Convert.ToString(int.Parse(jj.ToString()), 16).PadLeft(4, '0').Substring(0, 2), 16).ToString());
            txt_Rx6.Text = TX_com[10].ToString();
            txt_Rx7.Text = TX_com[11].ToString();


            //LED開
            byte[] data = HexStringToByteArray("55 AA 01 00 01 00 00 00 12 00 13 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(200);
            //開始建檔
            serialPort1.Write(TX_com, 0, 11);
            Thread.Sleep(1000);
            //採集指紋
            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 60 00 60 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(200);
            //第一次建檔
            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 23 00 23 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(1000);
            //採集指紋
            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 60 00 60 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(200);
            //第二次建檔
            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 24 00 24 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(1000);
            //採集指紋
            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 60 00 60 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(200);
            //第三次建檔
            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 25 00 25 01");
            serialPort1.Write(data, 0, data.Length);
            Thread.Sleep(1000);
            //LED關
            data = HexStringToByteArray("55 AA 01 00 00 00 00 00 12 00 12 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_TX_del_Click(object sender, EventArgs e) //刪除指定指紋
        {
            if (txt_Tx_var.Text == "")
            {
                MessageBox.Show("請輸入參數");
                return;
            }

            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            flag = 1; //回傳一行接收資料,判斷,1為刪除
            byte[] TX_com = new byte[12];
            //讀取開頭
            TX_com[0] = 85; //0x55
            TX_com[1] = 170; //0xAA

            TX_com[2] = 1;
            TX_com[3] = 0;
            //參數
            TX_com[4] = byte.Parse(txt_Tx_var.Text);
            TX_com[5] = 0;
            TX_com[6] = 0;
            TX_com[7] = 0;

            //命令代碼
            TX_com[8] = 64; //byte.Parse(textBox42.Text);
            TX_com[9] = 0;

            //check sum (從0累加至9) (jj:BD0A)
            int ii;
            int jj = 0;
            for (ii = 0; ii <= 9; ii++)
            {
                jj = jj + TX_com[ii];

            }
            //將累加值,分別存入陣列
            TX_com[10] = byte.Parse(Convert.ToInt32(Convert.ToString(int.Parse(jj.ToString()), 16).PadLeft(4, '0').Substring(2, 2), 16).ToString());
            TX_com[11] = byte.Parse(Convert.ToInt32(Convert.ToString(int.Parse(jj.ToString()), 16).PadLeft(4, '0').Substring(0, 2), 16).ToString());
            txt_Rx6.Text = TX_com[10].ToString();
            txt_Rx7.Text = TX_com[11].ToString();
            serialPort1.Write(TX_com, 0, 11);
        }

        private void btn_TX_IDuse_Click(object sender, EventArgs e) //用計時器方式快確認指紋是否建檔
        {
            id_count = 0;
            timer1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)//資料庫連結
        {
            //string dbpath = "DBMS.mdb";    //宣告資料庫所在的路徑變數
            string Source;                 //宣告連線的字串
            //Source = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbpath;
            Source = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBMS.mdb";
            //OleDbConnection conn;          //宣告連線的物件
            conn = new OleDbConnection(Source);   //連線
            conn.Open();          //開啟資料庫
            MessageBox.Show("成功連結到Access資料庫");
            //conn.Close();         //關閉資料庫
        }

        private void btn_addDB_Click(object sender, EventArgs e) //指紋建檔及新增資料庫
        {
            if (txt_number.Text == "")
            {
                MessageBox.Show("請輸入參數","警告", MessageBoxButtons.OK ,MessageBoxIcon.Warning );
                return;
            }
            
            //string SelectCmd = "INSERT INTO Person(編號,姓名,電話,手機,地址) VALUES('"
            //                    + txt_number.Text.Trim() + "','" + txt_name.Text.Trim() + "','" +
            //                    txt_tel.Text.Trim() + "','" + txt_phone.Text.Trim() + "','" +
            //                    txt_address.Text.Trim() + "')";
            //OleDbCommand Cmd=new OleDbCommand(SelectCmd,conn);
            //Cmd.ExecuteNonQuery();
            groupBox3.Enabled = false;
            txt_Tx_var.Text = txt_number.Text; //將編號傳至tab5的參數
            btn_TX_create_Click(sender, e); //指紋開始建檔 
            groupBox3.Enabled = true;

        }

        private void btn_delDB_Click(object sender, EventArgs e) //刪除指紋編號,指紋機+DB
        {
            //string SelectCmd = "DELETE * FROM Person WHERE 編號='"
            //                    + txt_number.Text.Trim() +"'";
            //OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
            //Cmd.ExecuteNonQuery();
            if (txt_number.Text == "")
            {
                MessageBox.Show("請輸入參數", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            txt_Tx_var.Text = txt_number.Text;
            btn_TX_del_Click(sender, e);
        }

        private void btn_fingerDB_Click(object sender, EventArgs e) //指紋比對DB
        {
            groupBox3.Enabled = false;
            btn_TX_verify_Click(sender,e);
            groupBox3.Enabled = true ;
        }

        private void btn_searchDB_Click(object sender, EventArgs e) //查詢DB
        {
            if (txt_number.Text == "" && txt_name.Text=="")
            {
                MessageBox.Show("請輸入參數", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            string SelectCmd = "select * from Person where 編號='"+ txt_number.Text.Trim()+"' or 姓名='" + txt_name.Text.Trim() + "'";

            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "Person");
            dataGridView1.DataSource = DtSet.Tables["Person"];
        }

        private void btn_showAllDB_Click(object sender, EventArgs e) //顯示人事全部資料
        {
            string SelectCmd = "select * from Person order by 編號 ASC";
            //OleDbCommand cmd = new OleDbCommand(SelectCmd, conn);
            //cmd.ExecuteReader();

            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "Person");
            dataGridView1.DataSource = DtSet.Tables["Person"];
        }

        private void btn_IDAllNumber_Click(object sender, EventArgs e) //取得總指紋數
        {
            btn_TX_IDnumber_Click(sender, e);
        }

        private void btn_show_arrivals_Click(object sender, EventArgs e) //顯示到班50筆資料
        {
            string SelectCmd = "select Top 50 arrivals.識別碼 ,arrivals.編號 ,Person.姓名 ,arrivals.年月日 ,arrivals.時分秒 from arrivals ,Person where arrivals.編號=Person.編號 order by arrivals.識別碼 desc";
            //OleDbCommand cmd = new OleDbCommand(SelectCmd, conn);
            //cmd.ExecuteReader();

            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "arrivals");
            dataGridView2.DataSource = DtSet.Tables["arrivals"];
        }

        private void button6_Click(object sender, EventArgs e) //單一辨識
        {
            
            btn_TX_verify_Click(sender, e);
            
        }

        private void timer2_Tick(object sender, EventArgs e) //自動辨識
        {
            button6.PerformClick();
        }

        private void btn_autoVery_Click(object sender, EventArgs e) //自動辨識
        {

            if (timer2.Enabled.ToString() == "False") {
                
                btn_autoVery.Text="停止辨識";
                btn_autoVery.BackColor = Color.Red;
                timer2.Enabled = true;
            }
            else 
            {
                btn_autoVery.Text = "自動辨識";
                btn_autoVery.BackColor = Color.Lime ;
                timer2.Enabled = false;
            }
               // timer2.Enabled = true;


                
                
            //MessageBox.Show(timer2.Enabled.ToString());
        }



        private void btn_search_arrivals_Click(object sender, EventArgs e) //到班查詢
        {
            string SelectCmd="";

            if (radioButton1.Checked.ToString() == "True")//用編號,姓名查詣
            {
                if (txt_arr_number.Text == "" && txt_arr_name.Text  == "")
                {
                    MessageBox.Show("請輸入參數", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                SelectCmd = "select arrivals.識別碼 ,arrivals.編號 ,Person.姓名 ,arrivals.年月日 ,arrivals.時分秒 from arrivals ,Person where arrivals.編號=Person.編號 and arrivals.編號='" + txt_arr_number.Text.Trim() + "' or Person.姓名='" + txt_arr_name.Text.Trim() + "'" + 
                            "order by arrivals.識別碼 desc";
            }

            if (radioButton2.Checked.ToString() == "True")//用時間查詢
            {
                //if (txt_arr_number.Text == "" && txt_arr_name.Text == "")
                //{
                //    MessageBox.Show("請輸入參數", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    return;
                //}
                if (txt_arr_name.Text == "")
                {
                    SelectCmd = "select arrivals.識別碼 ,arrivals.編號 ,Person.姓名 ,arrivals.年月日 ,arrivals.時分秒 from arrivals ,Person where arrivals.編號=Person.編號 and arrivals.年月日 between #" + dateTimePicker1.Value.ToShortDateString() + "# and #" +
                                 dateTimePicker2.Value.ToShortDateString() + "# order by arrivals.識別碼 desc";
                }
                else {
                    SelectCmd = "select arrivals.識別碼 ,arrivals.編號 ,Person.姓名 ,arrivals.年月日 ,arrivals.時分秒 from arrivals ,Person where arrivals.編號=Person.編號 and Person.姓名='"+ txt_arr_name.Text.Trim() + "' and arrivals.年月日 between #" + dateTimePicker1.Value.ToShortDateString() + "# and #" +
                                     dateTimePicker2.Value.ToShortDateString() + "# order by arrivals.識別碼 desc";
                
                }
            }

            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "arrivals");
            dataGridView2.DataSource = DtSet.Tables["arrivals"];
        }

        private void btn_TX_delall_Click(object sender, EventArgs e) //刪除所有指紋
        {
            serialPort1.DiscardInBuffer();
            serialPort1.DiscardOutBuffer();
            bytes_count = 0;
            j = 0;
            byte[] data = HexStringToByteArray("55 AA 01 00 02 00 00 00 41 00 43 01");
            serialPort1.Write(data, 0, data.Length);
        }

        private void btn_delallFinger_Click(object sender, EventArgs e) //刪除所有指紋
        {
            flag = 3;
            btn_TX_delall_Click(sender, e);
        }

        private void btn_delallPDB_Click(object sender, EventArgs e)//刪除所有人事資料
        {
            string SelectCmd = "DELETE * FROM Person ";                           
            OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
            Cmd.ExecuteNonQuery();
            btn_showAllDB.PerformClick();
            MessageBox.Show("已刪除所有人事資料庫", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }



        private void btn_show_admin_Click(object sender, EventArgs e) //顯示管理者資料
        {
            string SelectCmd = "select * from admin" ; // order by 編號 ASC";
            //OleDbCommand cmd = new OleDbCommand(SelectCmd, conn);
            //cmd.ExecuteReader();

            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "admin");
            dataGridView3.DataSource = DtSet.Tables["admin"];
        }

        private void btn_add_admin_Click(object sender, EventArgs e) //新增管理帳號
        {
            if (txt_account_admin.Text == "" || txt_name_admin.Text == "")
            {
                MessageBox.Show("請輸入參數", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string SelectCmd = "INSERT INTO admin(姓名,帳號,密碼,管理身份) VALUES('"
                                + txt_name_admin .Text.Trim() + "','" + txt_account_admin.Text.Trim() + "','" +
                                txt_passwd_admin.Text.Trim() + "','" + cob_admin_admin.Text.Trim() + "')";
            OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
            Cmd.ExecuteNonQuery();
            btn_show_admin.PerformClick();
            txt_name_admin.Text = "";
            txt_passwd_admin.Text = "";
            txt_account_admin.Text = "";

            MessageBox.Show("成功新增", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information );
        }

        private void btn_login_Click(object sender, EventArgs e) //登入
        {
            if (txt_loginID.Text  == "" || txt_loginPW.Text  == "")
            {
                MessageBox.Show("請輸入參數", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tabControl1.TabPages.Remove(tabPage2);  //人事資料
                tabControl1.TabPages.Remove(tabPage3);  //系統管理
                btn_delall_arr.Enabled = false;
                btn_delallPDB.Enabled = false;
                btn_delallFinger.Enabled = false;
                btn_search_arrivals.Enabled = false;

                btn_excel.Enabled = false;
                //請假管理
                btn_lea_del.Enabled = false; //刪除
                btn_lea_delall.Enabled = false; //刪除所有
                btn_lea_excel.Enabled = false; //輸出EXCEL
                btn_lea_query.Enabled = false; //查詢
                return;
            }
            
            string loginid="", loginpw="" ,loginadmin="";
            string SelectCmd = "select * from admin where 帳號='"+ txt_loginID.Text.Trim() +"'" ; // order by 編號 ASC";
            OleDbCommand cmd = new OleDbCommand(SelectCmd, conn);
            OleDbDataReader reader= cmd.ExecuteReader();

            while (reader.Read()) {
                loginid=reader["帳號"].ToString();
                loginpw=reader["密碼"].ToString();
                loginadmin = reader["管理身份"].ToString();
            }
            if (txt_loginID.Text == loginid.ToString() && txt_loginPW.Text == loginpw.ToString())
            {

                MessageBox.Show("登入成功","訊息",MessageBoxButtons .OK ,MessageBoxIcon.Information );
                txt_loginID.Text = "";
                txt_loginPW.Text = "";

                if (loginadmin == "高階管理者") { 
                    tabControl1.TabPages.Add(tabPage2);  //人事資料
                    tabControl1.TabPages.Add(tabPage3);  //系統管理
                    btn_delall_arr.Enabled = true;
                    btn_delallPDB.Enabled = true;
                    btn_delallFinger.Enabled = true;
                    btn_search_arrivals.Enabled = true;
                 
                    btn_excel.Enabled = true;
                    //請假管理
                    btn_lea_del.Enabled = true; //刪除
                    btn_lea_delall.Enabled = true; //刪除所有
                    btn_lea_excel.Enabled = true; //輸出EXCEL
                    btn_lea_query.Enabled = true; //查詢
                }
                if (loginadmin == "中階管理者")
                {
                    tabControl1.TabPages.Remove(tabPage2);  //人事資料
                    tabControl1.TabPages.Add(tabPage2);  //人事資料
                    tabControl1.TabPages.Remove(tabPage3);  //系統管理
                    btn_delall_arr.Enabled = false;
                    btn_delallPDB.Enabled = false;
                    btn_delallFinger.Enabled = false;
                    btn_search_arrivals.Enabled = true;
                    
                    btn_excel.Enabled = true;

                    //請假管理
                    btn_lea_del.Enabled = true; //刪除
                    btn_lea_delall.Enabled = false; //刪除所有
                    btn_lea_excel.Enabled = true; //輸出EXCEL
                    btn_lea_query.Enabled = true; //查詢
                    
                }
                if (loginadmin == "一般使用者")
                {
                    tabControl1.TabPages.Remove(tabPage2);  //人事資料
                    tabControl1.TabPages.Remove(tabPage3);  //系統管理
                    btn_delall_arr.Enabled = false;
                    btn_delallPDB.Enabled = false;
                    btn_delallFinger.Enabled = false;
                    btn_search_arrivals.Enabled = true;
                    
                    btn_excel.Enabled = true;
                    //請假管理
                    btn_lea_del.Enabled = false; //刪除
                    btn_lea_delall.Enabled = false; //刪除所有
                    btn_lea_excel.Enabled = true; //輸出EXCEL
                    btn_lea_query.Enabled = true; //查詢
                }
                
                
            }
            else {
                MessageBox.Show("登入失敗", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tabControl1.TabPages.Remove(tabPage2);  //人事資料
                tabControl1.TabPages.Remove(tabPage3);  //系統管理
                btn_delall_arr.Enabled = false;
                btn_delallPDB.Enabled = false;
                btn_delallFinger.Enabled = false;
                btn_search_arrivals.Enabled = false;
                
                btn_excel.Enabled = false;
                //請假管理
                btn_lea_del.Enabled = false; //刪除
                btn_lea_delall.Enabled = false; //刪除所有
                btn_lea_excel.Enabled = false; //輸出EXCEL
                btn_lea_query.Enabled = false; //查詢
            }
            


        }

        private void btn_delall_arr_Click(object sender, EventArgs e) //刪除所有出勤資料
        {
            string SelectCmd = "DELETE * FROM arrivals ";
            OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
            Cmd.ExecuteNonQuery();
            btn_showAllDB.PerformClick();
            MessageBox.Show("已刪除所有出勤資料", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void btn_logout_Click(object sender, EventArgs e) //登出
        {
            
            tabControl1.TabPages.Remove(tabPage2);  //人事資料
            tabControl1.TabPages.Remove(tabPage3);  //系統管理
            btn_delall_arr.Enabled = false;
            btn_delallPDB.Enabled = false;
            btn_delallFinger.Enabled = false;
            btn_search_arrivals.Enabled = false;
            btn_excel.Enabled = false;
            //請假管理
            btn_lea_del.Enabled = false; //刪除
            btn_lea_delall.Enabled = false; //刪除所有
            btn_lea_excel.Enabled = false; //輸出EXCEL
            btn_lea_query.Enabled = false; //查詢
            MessageBox.Show("成功登出", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information );
        }

        private void btn_del_admin_Click(object sender, EventArgs e) //刪除單一管理者
        {

            string SelectCmd = "DELETE * FROM admin WHERE 姓名='"
                                                        + txt_name_admin .Text.Trim() + "'";
            //OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
            
            //Cmd.ExecuteNonQuery();
            //btn_show_admin.PerformClick();
            //txt_name_admin.Text = "";
            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "admin");
            dataGridView3.DataSource = DtSet.Tables["admin"];
            txt_name_admin.Text = "";
            MessageBox.Show("已刪除資料", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void btn_query_admin_Click(object sender, EventArgs e) //查詢管理者
        {
            string SelectCmd = "select * FROM admin WHERE 姓名='"
                                                        + txt_name_admin.Text.Trim() + "'";
            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "admin");
            dataGridView3.DataSource = DtSet.Tables["admin"];
            txt_name_admin.Text = "";
        }

        private void btn_excel_Click(object sender, EventArgs e) //輸出Excel
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK) //開檔
            {
                FileInfo f = new FileInfo(saveFileDialog1.FileName );
                //TextWriter sw = new StreamWriter(@"F:\Text.txt");
                StreamWriter sw = f.CreateText();
                for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        sw.Write(dataGridView2.Rows[i].Cells[j].Value.ToString() + "\t");
                    }
                    sw.Write("\r\n");
                }
                sw.Close();
                MessageBox.Show("存檔成功","訊息",MessageBoxButtons.OK ,MessageBoxIcon.Information );
            }
        }

        

        //點休假管理,自動載入人事姓名
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (tabControl1.SelectedTab.Text.ToString() == "請假管理")
            {
                if (com_lea_person.Items.Count == 0)
                {
                    string SelectCmd = "select * from Person order by 姓名 ASC";
                    OleDbCommand cmd = new OleDbCommand(SelectCmd, conn);
                    OleDbDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        com_lea_name.Items.Add(reader["姓名"].ToString());
                        com_lea_person.Items.Add(reader["姓名"].ToString());
                        com_lea_manager.Items.Add(reader["姓名"].ToString());
                    }
                }

            }
        }


        //自動填入職代mail
        private void com_lea_person_SelectedIndexChanged(object sender, EventArgs e)
        {
            string SelectCmd = "select * from Person where 姓名='" +
                                  com_lea_person.Text.ToString() + "'";
            OleDbCommand cmd = new OleDbCommand(SelectCmd, conn);
            OleDbDataReader reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                if (txt_lea_mail.Text == "")
                {
                    txt_lea_mail.Text = reader["mail"].ToString();
                }
                else { 
                    txt_lea_mail.Text = reader["mail"].ToString() +","+ txt_lea_mail.Text;
                }
            }
        }
        
        //自動填入主管mail
        private void com_lea_manager_SelectedIndexChanged(object sender, EventArgs e)
        {
            string SelectCmd = "select * from Person where 姓名='" +
                                  com_lea_manager.Text.ToString() + "'";
            OleDbCommand cmd = new OleDbCommand(SelectCmd, conn);
            OleDbDataReader reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                if (txt_lea_mail.Text == "")
                {
                    txt_lea_mail.Text = reader["mail"].ToString();
                }
                else
                {
                    txt_lea_mail.Text = reader["mail"].ToString() + "," + txt_lea_mail.Text;
                }
            }
        }

        //點選姓名,自動填入編號
        private void com_lea_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            string SelectCmd = "select * from Person where 姓名='" +
                                              com_lea_name.Text.ToString() + "'";
            OleDbCommand cmd = new OleDbCommand(SelectCmd, conn);
            OleDbDataReader reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                txt_lea_num.Text = reader["編號"].ToString();

                if (txt_lea_mail.Text == "")
                {
                    txt_lea_mail.Text = reader["mail"].ToString();
                }
                else
                {
                    txt_lea_mail.Text = reader["mail"].ToString() + "," + txt_lea_mail.Text;
                }
            }

           
        }

        //休假管理,顯示前50筆
        private void btn_lea_show_Click(object sender, EventArgs e)
        {
            string SelectCmd = "select top 50 * from Leave order by 識別碼 desc";
            //OleDbCommand cmd = new OleDbCommand(SelectCmd, conn);
            //cmd.ExecuteReader();

            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "Leave");
            dataGridView4.DataSource = DtSet.Tables["Leave"];
        }

        //新增請假
        private void btn_lea_add_Click(object sender, EventArgs e)
        {
            if (com_lea_name.Text == "" || com_lea_day.Text=="" ||com_lea_manager.Text=="")
            {
                MessageBox.Show("請輸入參數", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btn_lea_calculate.PerformClick();

            //寄信
            if (chk_lea_sentmail.Checked == true)
            {
                //先寄信
                try
                {

                    //SMTP伺服器,port.google:smtp.gmail.com hotmail:smtp.live.com
                    SmtpClient SmtpServer = new SmtpClient("smtp.live.com", 587);
                    //寄信逾時設定
                    SmtpServer.Timeout = 10000;
                    //設定是否用SSL加密連線
                    SmtpServer.EnableSsl = true;
                    //使用預設SMTP郵件伺服器之認證
                    SmtpServer.UseDefaultCredentials = true;
                    //設定如何傳送郵件訊息
                    SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                    //輸入mail帳號,密碼
                    SmtpServer.Credentials = new System.Net.NetworkCredential("chander_service@hotmail.com", "971206@abcgod");
                    //信件
                    MailMessage mail = new MailMessage();
                    mail.From = new MailAddress("chander_service@hotmail.com", "差勤系統訊息");
                    mail.To.Add(txt_lea_mail.Text.ToString());
                    mail.Subject = com_lea_name.Text + "休假通知";
                    mail.Body = "編號：" + txt_lea_num.Text + "\r\n" +
                                "姓名：" + com_lea_name.Text + "\r\n" +
                                "職務代理人：" + com_lea_person.Text + "\r\n" +
                                "假別：" + com_lea_day.Text + "\r\n" +
                                "事由：" + txt_lea_reason.Text + "\r\n" +
                                "時間(起)：" + dateP_lea_start.Value.ToShortDateString() + "  " + com_lea_Sh.Text + ":" + com_lea_Sm.Text + "\r\n" +
                                "時間(迄)：" + date_lea_end.Value.ToShortDateString() + "  " + com_lea_Eh.Text + ":" + com_lea_Em.Text + "\r\n" +
                                "總時數：" + txt_lea_calcultime.Text + " 小時";
                    //夾檔
                    //Attachment data = new Attachment(@"D:\BT.txt");
                    //mail.Attachments.Add(data);

                    SmtpServer.Send(mail);
                    MessageBox.Show("郵件成功寄出", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    SmtpServer = null;
                    mail.Dispose();
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.ToString());
                    MessageBox.Show("信件無法寄出。\r\n請檢查網路及mail是否無誤", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

            }

            //新增請假記錄至資料庫
            string SelectCmd = "INSERT INTO Leave(編號,姓名,假別,事由,職代,主管,年月日起,時分起,年月日迄,時分迄,小計) VALUES('"
                                + txt_lea_num.Text.Trim() + "','" + com_lea_name.Text.Trim() + "','" +
                                com_lea_day.Text.Trim() + "','" + txt_lea_reason.Text.Trim() + "','" +
                                com_lea_person.Text.Trim() + "','" + com_lea_manager.Text.Trim() + "','" +
                                dateP_lea_start.Value.ToShortDateString()  + "','" + com_lea_Sh.Text.ToString() +":"+com_lea_Sm.Text.ToString() + "','" +
                                date_lea_end.Value.ToShortDateString() + "','" + com_lea_Eh.Text.ToString() + ":" + com_lea_Em.Text.ToString() + "','" +
                                txt_lea_calcultime.Text.Trim() + "')";
                
            OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
            Cmd.ExecuteNonQuery();
            btn_lea_show.PerformClick();
            
            com_lea_name.Text = "";
            com_lea_day.Text = "";
            com_lea_person.Text = "";
            com_lea_manager.Text = "";
            txt_lea_num.Text = "";
            txt_lea_reason.Text = "";
            txt_lea_mail.Text = "";

            MessageBox.Show("新增請假記錄","訊息",MessageBoxButtons.OK ,MessageBoxIcon.Information);

        }

       

        //計算休假時數
        private void btn_lea_calculate_Click(object sender, EventArgs e)
        {
            //計算天數
            //textBox5.Text=(DateTime.Parse(date_lea_end.Value.ToShortDateString())- DateTime.Parse(dateP_lea_start.Value.ToShortDateString())).TotalDays.ToString() ;

            float starttime, endtime;
            int break_day;

            break_day = System.Convert.ToInt16((DateTime.Parse(date_lea_end.Value.ToShortDateString()) - DateTime.Parse(dateP_lea_start.Value.ToShortDateString())).TotalDays);


            //計算-時
            if (chk_lea_break.Checked == true)
            { // 要扣午休時間
                //開始時間,時+分
                starttime = float.Parse(com_lea_Sh.Text.ToString()) + (float.Parse(com_lea_Sm.Text.ToString()) / 60);
                //結束時間,時+分
                endtime = float.Parse(com_lea_Eh.Text.ToString()) + (float.Parse(com_lea_Em.Text.ToString()) / 60);
                txt_lea_calcultime.Text = ((endtime - starttime - float.Parse(txt_lea_breadtime.Text.ToString())) + (break_day * 8)).ToString();
                MessageBox.Show("請假時數："+txt_lea_calcultime.Text +" 小時","訊息",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            else
            {    //不用扣午休 
                //開始時間,時+分
                starttime = float.Parse(com_lea_Sh.Text.ToString()) + (float.Parse(com_lea_Sm.Text.ToString()) / 60);
                //結束時間,時+分
                endtime = float.Parse(com_lea_Eh.Text.ToString()) + (float.Parse(com_lea_Em.Text.ToString()) / 60);
                txt_lea_calcultime.Text = ((endtime - starttime) + (break_day * 8)).ToString();
                MessageBox.Show("請假時數：" + txt_lea_calcultime.Text + " 小時", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //請假寄信
        private void btn_lea_sentmail_Click(object sender, EventArgs e)
        {
            btn_lea_calculate.PerformClick();
            try
            {

                //SMTP伺服器,port.google:smtp.gmail.com hotmail:smtp.live.com
                SmtpClient SmtpServer = new SmtpClient("smtp.live.com", 587);
                //寄信逾時設定
                SmtpServer.Timeout = 10000;
                //設定是否用SSL加密連線
                SmtpServer.EnableSsl = true;
                //使用預設SMTP郵件伺服器之認證
                SmtpServer.UseDefaultCredentials = true;
                //設定如何傳送郵件訊息
                SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                //輸入mail帳號,密碼
                SmtpServer.Credentials = new System.Net.NetworkCredential("chander_service@hotmail.com", "971206@abcgod");

                //信件
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress("chander_service@hotmail.com", "差勤系統訊息");

                mail.To.Add(txt_lea_mail.Text.ToString());
                //mail.To.Add("chander_lin@hotmail.com");
                //mail.Bcc.Add("chiender@iii.org.tw"); //密件

                //mail.To.Add("chiender@iii.org.tw,chander_lin@hotmail.com");
                //mail.To.Add("chander_lin@hotmail.com");
                //mail.BodyEncoding=System.Text.Encoding.Unicode;
                //mail.SubjectEncoding = System.Text.Encoding.Unicode;
                mail.Subject = com_lea_name.Text+"休假通知";
                mail.Body = "編號："+txt_lea_num.Text+"\r\n"+
                            "姓名："+com_lea_name.Text+"\r\n"+
                            "職務代理人："+com_lea_person.Text+"\r\n"+
                            "假別："+com_lea_day .Text +"\r\n"+
                            "事由："+txt_lea_reason.Text +"\r\n"+
                            "時間(起)：" + dateP_lea_start.Value.ToShortDateString() + "  " + com_lea_Sh.Text + ":" + com_lea_Sm.Text + "\r\n" +
                            "時間(迄)：" + date_lea_end.Value.ToShortDateString()+"  "+com_lea_Eh.Text+":"+com_lea_Em .Text+ "\r\n"+
                            "總時數："+txt_lea_calcultime.Text +" 小時";


                //夾檔
                //Attachment data = new Attachment(@"D:\BT.txt");
                //mail.Attachments.Add(data);

                SmtpServer.Send(mail);
                MessageBox.Show("郵件成功寄出", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                SmtpServer = null;
                mail.Dispose();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                MessageBox.Show("信件無法寄出。\r\n請檢查網路及mail是否無誤", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
            }
        }

        //刪除所有休假記錄
        private void btn_lea_delall_Click(object sender, EventArgs e)
        {
            string SelectCmd = "DELETE * FROM Leave ";
            OleDbCommand Cmd = new OleDbCommand(SelectCmd, conn);
            Cmd.ExecuteNonQuery();
            btn_lea_show.PerformClick();
            MessageBox.Show("已刪除所有出勤資料", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        //請假查詢
        private void btn_lea_query_Click(object sender, EventArgs e)
        {
            string SelectCmd = "";
            SelectCmd = "select * from Leave where 姓名='" + com_lea_name.Text.Trim() + "' and 年月日起 between #" + dateP_lea_start.Value.ToShortDateString()+ "# and #" +
                                     date_lea_end.Value.ToShortDateString() + "# order by 識別碼 desc";

            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "Leave");
            dataGridView4.DataSource = DtSet.Tables["Leave"];
        }

        //刪除請假記錄
        private void btn_lea_del_Click(object sender, EventArgs e)
        {
            string SelectCmd = "";
            SelectCmd = "delete * from Leave where 姓名='" + com_lea_name.Text.Trim() + "' and 年月日起 between #" + dateP_lea_start.Value.ToShortDateString() + "# and #" +
                                     date_lea_end.Value.ToShortDateString() + "# ";

            //宣告物件
            OleDbDataAdapter DtApter;
            DataSet DtSet;
            DtApter = new OleDbDataAdapter(SelectCmd, conn);
            DtSet = new DataSet();
            //讀取資料表
            DtApter.Fill(DtSet, "Leave");
            dataGridView4.DataSource = DtSet.Tables["Leave"];
            btn_lea_show.PerformClick();
            MessageBox.Show("已刪除請假記錄","訊息",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        //請假記錄輸出EXCEL
        private void btn_lea_excel_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK) //開檔
            {
                FileInfo f = new FileInfo(saveFileDialog1.FileName);
                //TextWriter sw = new StreamWriter(@"F:\Text.txt");
                StreamWriter sw = f.CreateText();
                for (int i = 0; i < dataGridView4.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView4.Columns.Count; j++)
                    {
                        sw.Write(dataGridView4.Rows[i].Cells[j].Value.ToString() + "\t");
                    }
                    sw.Write("\r\n");
                }
                sw.Close();
                MessageBox.Show("存檔成功", "訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //按Enter,登入
        private void tabControl1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode) { 
            
                case Keys.Enter :
                    btn_login.PerformClick();
                break ;

            }
        }


   }     
}
