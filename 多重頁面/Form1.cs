using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Net;
using Renci.SshNet;
using System.Threading;
using Microsoft.VisualBasic;

namespace 多重頁面
{
    public partial class Form1 : Form
    {
        string[] TD_area = { "北市區", "北北區", "北南區", "北西區" };
        public Form1()
        {
            InitializeComponent();

            comboBox1.Items.Add("自定義");
            comboBox1.Items.Add("MN (固定IP)");
            comboBox1.Items.Add("資訊局 (Routing)");
            comboBox1.Items.Add("資訊局 (Bridge)");
            comboBox1.Items.Add("圖書館 (Bridge)");
            comboBox1.Items.Add("民防 (Routing)");
            comboBox1.Items.Add("停管處 (Routing)");
            comboBox1.Items.Add("停管處 (Bridge)");

            // ComboBox　預設顯示值為 
            comboBox1.SelectedIndex = 0;
            panel1.Visible = false;
        }
        //========================================主頁面========================================
        private void AutoOpen_button_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = false;

        }
        private void TD_button_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = true;
        }





        //========================================自動開通========================================
        // 桌布更換設定
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam);




        //  OLT 選單
        string[] T100 = { "T100-011", "T100-021", "T100-031", "T100-041", "T100-061", "T100-071" };
        string[] T103 = { "T103-001", "T103-011", "T103-021", "T103-031", "T103-041", "T103-051" };
        string[] T104 = { "T104-001", "T104-011", "T104-031", "T104-041", "T104-051", "T104-061", "T104-071", "T104-081" };
        string[] T105 = { "T105-001", "T105-011", "T105-021", "T105-031", "T105-041", "T105-051", "T105-061", "T105-071", "T105-081" };
        string[] T106 = { "T106-001", "T106-011", "T106-021", "T106-031", "T106-041", "T106-051", "T106-061", "T106-071", "T106-081", "T106-101", "T106-121" };
        string[] T108 = { "T108-011", "T108-021", "T108-031", "T108-041", "T108-051", "T108-061" };
        string[] T110 = { "T110-001", "T110-021", "T110-031", "T110-041", "T110-051", "T110-061", "T110-081", "T110-091", "T110-101" };
        string[] T111 = { "T111-001", "T111-011", "T111-021", "T111-031", "T111-041", "T111-061", "T111-071", "T111-081", "T111-091", "T111-101", "T111-111" };
        string[] T112 = { "T112-001", "T112-011", "T112-031", "T112-041", "T112-051", "T112-061", "T112-071", "T112-081" };
        string[] T114 = { "T114-001", "T114-002", "T114-004", "T114-011", "T114-021", "T114-031", "T114-041", "T114-051", "T114-061", "T114-071", "T114-081", "T114-091", "T114-101" };
        string[] T115 = { "T115-011", "T115-021", "T115-031", "T115-041", "T115-051", "T115-061" };
        string[] T116 = { "T116-011", "T116-021", "T116-031", "T116-041", "T116-051", "T116-061", "T116-071", "T116-081", "T116-091", "T116-101" };



        // ====================================主要按鈕====================================
        private void grow_button_Click(object sender, EventArgs e)
        {
            auto_button.Enabled = true;

            ONT_Profile.Slot = Slot_Box.Text;
            ONT_Profile.Port = Port_Box.Text;
            ONT_Profile.Onuid = OnuID_Box.Text;
            ONT_Profile.Sn = SN_Box.Text;
            ONT_Profile.Des = Des_Box.Text;
            ONT_Profile.mode = Mode_Box.Text;
            ONT_Profile.Svlan = Svlan_Box.Text;
            ONT_Profile.Cvlan = Cvlan_Box.Text;
            ONT_Profile.Bw_Up = UP_Box.Text;
            ONT_Profile.Bw_Down = Down_Box.Text;
            ONT_Profile.Ip = IP_Box.Text;
            ONT_Profile.Gw = GW_Box.Text;
            ONT_Profile.Bw_Per = Percen_Box.Text;
            if (Des_Box.Text != "" && Des_Box.Text.Length == 10)
            {
                ONT_Profile.MV = Des_Box.Text.Substring(6, 4);
            }
            else
            {
                MessageBox.Show("你電路編號是不是填錯阿!");
            }

            ONT_Profile.Group7750();


            string total =
                "=====OLT部分：=====\r\n" +
                ONT_Profile.Dba_profile() +
                ONT_Profile.Vlan_profile() +
                ONT_Profile.Traffic_profile() +
                ONT_Profile.Onu_profile() +
                ONT_Profile.GPON_OLT() +
                ONT_Profile.OLT_Vlan() +
                "\r\n=====10K部分：=====\r\n" +
                ONT_Profile.K10() +
                "\r\n=====7750或7450部分：=====\r\n" +
                ONT_Profile.ALU7750()
                ;




            richTextBox1.Text = total;



        }







        // 行政區 → OLT 選單
        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox2.SelectedItem.ToString())
            {
                case @"T100 (中正區)":
                    ONT_Profile.Area = "100";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T100)
                    {

                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T103 (大同區)":
                    ONT_Profile.Area = "103";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T103)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T104 (中山區)":
                    ONT_Profile.Area = "104";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T104)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T105 (松山區)":
                    ONT_Profile.Area = "105";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T105)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T106 (大安區)":
                    ONT_Profile.Area = "106";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T106)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T108 (萬華區)":
                    ONT_Profile.Area = "108";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T108)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T110 (信義區)":
                    ONT_Profile.Area = "110";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T110)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T111 (士林區)":
                    ONT_Profile.Area = "111";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T111)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T112 (北投區)":
                    ONT_Profile.Area = "112";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T112)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T114 (內湖區)":
                    ONT_Profile.Area = "114";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T114)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T115 (南港區)":
                    ONT_Profile.Area = "115";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T115)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                case @"T116 (文山區)":
                    ONT_Profile.Area = "116";
                    comboBox3.Items.Clear();
                    comboBox3.ResetText();
                    if (comboBox1.Text == "MN (固定IP)")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                    }
                    foreach (var item in T116)
                    {
                        comboBox3.Items.Add(item);
                    }
                    break;
                default:
                    break;
            }

            if (comboBox1.Text == "MN (固定IP)" && comboBox3.Text != "")
            {
                Svlan_Box.Text = ((int.Parse(ONT_Profile.Area) - 100) * 100 + (((int.Parse(comboBox3.Text.Substring(5, 3)) - 1) / 10) * 4 + 1) + 1000).ToString();
            }
        }



        // ====================================專案選單====================================
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedItem.ToString())
            {
                case "自定義":
                    Slot_Box.Clear();
                    Port_Box.Clear();
                    OnuID_Box.Clear();
                    SN_Box.Clear();
                    Des_Box.Clear();
                    Mode_Box.Clear();
                    Svlan_Box.Clear();
                    Cvlan_Box.Clear();
                    UP_Box.Clear();
                    Down_Box.Clear();
                    IP_Box.Clear();
                    GW_Box.Clear();
                    textBox7.Clear();
                    textBox10.Clear();
                    textBox9.Clear();
                    textBox10.Clear();
                    textBox11.Clear();
                    textBox12.Clear();
                    Percen_Box.Enabled = true;
                    OnuID_Box.Enabled = true;
                    SN_Box.Enabled = true;
                    Des_Box.Enabled = true;
                    Mode_Box.Enabled = true;
                    Svlan_Box.Enabled = true;
                    Cvlan_Box.Enabled = true;
                    UP_Box.Enabled = true;
                    IP_Box.Enabled = true;
                    GW_Box.Enabled = true;
                    Percen_Box.Enabled = true;
                    textBox7.Enabled = true;
                    textBox8.Enabled = true;
                    textBox9.Enabled = true;
                    textBox10.Enabled = true;
                    textBox11.Enabled = true;
                    textBox12.Enabled = true;
                    break;

                case "資訊局 (Routing)":
                    Svlan_Box.Text = "602";
                    Svlan_Box.Enabled = false;
                    Mode_Box.Text = "4";
                    Mode_Box.Enabled = false;
                    textBox7.Enabled = false;
                    textBox8.Enabled = false;
                    textBox9.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    textBox12.Enabled = false;
                    IP_Box.Enabled = true;
                    GW_Box.Enabled = true;
                    Cvlan_Box.Enabled = true;
                    Cvlan_Box.Clear();
                    break;
                case "資訊局 (Bridge)":
                    Svlan_Box.Text = "602";
                    Svlan_Box.Enabled = false;
                    Mode_Box.Text = "0";
                    Mode_Box.Enabled = false;
                    textBox7.Enabled = false;
                    textBox8.Enabled = false;
                    textBox9.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    textBox12.Enabled = false;
                    Cvlan_Box.Clear();
                    IP_Box.Clear();
                    GW_Box.Clear();
                    IP_Box.Enabled = false;
                    GW_Box.Enabled = false;
                    Cvlan_Box.Enabled = true;
                    break;
                case "MN (固定IP)":
                    if (comboBox2.Text == "" || comboBox3.Text == "")
                    {
                        Svlan_Box.Text = "尚未選擇OLT";
                        Cvlan_Box.Text = "尚未填入完整Port位";
                        Svlan_Box.Enabled = false;
                        Cvlan_Box.Enabled = false;
                        Mode_Box.Text = "0";
                        Mode_Box.Enabled = false;
                        textBox7.Enabled = false;
                        textBox8.Enabled = false;
                        textBox9.Enabled = false;
                        textBox10.Enabled = false;
                        textBox11.Enabled = false;
                        textBox12.Enabled = false;
                        IP_Box.Clear();
                        GW_Box.Clear();
                        IP_Box.Enabled = false;
                        GW_Box.Enabled = false;
                        textBox7.Enabled = true;
                        textBox8.Enabled = true;
                        textBox9.Enabled = true;
                        textBox10.Enabled = true;
                        textBox11.Enabled = true;
                        textBox12.Enabled = true;
                    }
                    else
                    {
                        Svlan_Box.Text = ((int.Parse(ONT_Profile.Area) - 100) * 100 + (((int.Parse(comboBox3.Text.Substring(5, 3)) - 1) / 10) * 4 + 1) + 1000).ToString();
                        if (comboBox1.Text == "MN (固定IP)" && Port_Box.Text != "" && Slot_Box.Text != "" && OnuID_Box.Text != "")
                        {
                            Cvlan_Box.Text = ((((int.Parse(Slot_Box.Text) - 1) * 4 + int.Parse(Port_Box.Text)) - 1) * 32 + (int.Parse(OnuID_Box.Text) - 1) + 1000).ToString();
                        }
                        Svlan_Box.Enabled = false;
                        Cvlan_Box.Enabled = false;
                        Mode_Box.Text = "0";
                        Mode_Box.Enabled = false;
                        textBox7.Enabled = false;
                        textBox8.Enabled = false;
                        textBox9.Enabled = false;
                        textBox10.Enabled = false;
                        textBox11.Enabled = false;
                        textBox12.Enabled = false;
                        IP_Box.Clear();
                        GW_Box.Clear();
                        IP_Box.Enabled = false;
                        GW_Box.Enabled = false;
                        textBox7.Enabled = true;
                        textBox8.Enabled = true;
                        textBox9.Enabled = true;
                        textBox10.Enabled = true;
                        textBox11.Enabled = true;
                        textBox12.Enabled = true;
                    }

                    break;
                case "圖書館 (Bridge)":
                    Svlan_Box.Text = "605";
                    Svlan_Box.Enabled = false;
                    Mode_Box.Text = "0";
                    Mode_Box.Enabled = false;
                    textBox7.Enabled = false;
                    textBox8.Enabled = false;
                    textBox9.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    textBox12.Enabled = false;
                    Cvlan_Box.Clear();
                    IP_Box.Clear();
                    GW_Box.Clear();
                    IP_Box.Enabled = false;
                    GW_Box.Enabled = false;
                    Cvlan_Box.Enabled = true;
                    break;
                case "民防 (Routing)":
                    Svlan_Box.Text = "424";
                    Svlan_Box.Enabled = false;
                    Mode_Box.Text = "4";
                    Mode_Box.Enabled = false;
                    textBox7.Enabled = false;
                    textBox8.Enabled = false;
                    textBox9.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    textBox12.Enabled = false;
                    IP_Box.Enabled = true;
                    GW_Box.Enabled = true;
                    Cvlan_Box.Enabled = true;
                    Cvlan_Box.Clear();
                    break;
                case "停管處 (Routing)":
                    Svlan_Box.Text = "610";
                    Svlan_Box.Enabled = false;
                    Mode_Box.Text = "4";
                    Mode_Box.Enabled = false;
                    textBox7.Enabled = false;
                    textBox8.Enabled = false;
                    textBox9.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    textBox12.Enabled = false;
                    IP_Box.Enabled = true;
                    GW_Box.Enabled = true;
                    Cvlan_Box.Enabled = true;
                    Cvlan_Box.Clear();
                    break;
                case "停管處 (Bridge)":
                    Svlan_Box.Text = "610";
                    Svlan_Box.Enabled = false;
                    Mode_Box.Text = "0";
                    Mode_Box.Enabled = false;
                    textBox7.Enabled = false;
                    textBox8.Enabled = false;
                    textBox9.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    textBox12.Enabled = false;
                    Cvlan_Box.Clear();
                    IP_Box.Clear();
                    GW_Box.Clear();
                    IP_Box.Enabled = false;
                    GW_Box.Enabled = false;
                    Cvlan_Box.Enabled = true;
                    break;

                default:
                    break;
            }
        }










        // ====================================按鈕切換判斷====================================

        private void ComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "MN (固定IP)" && comboBox2.Text != "")
            {
                Svlan_Box.Text = ((int.Parse(ONT_Profile.Area) - 100) * 100 + (((int.Parse(comboBox3.Text.Substring(5, 3)) - 1) / 10) * 4 + 1) + 1000).ToString();
            }
        }

        private void Slot_Box_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "MN (固定IP)" && Port_Box.Text != "" && Slot_Box.Text != "" && OnuID_Box.Text != "")
            {
                Cvlan_Box.Text = ((((int.Parse(Slot_Box.Text) - 1) * 4 + int.Parse(Port_Box.Text)) - 1) * 32 + (int.Parse(OnuID_Box.Text) - 1) + 1000).ToString();
            }
            if (comboBox1.Text == "MN (固定IP)")
            {
                if (Port_Box.Text == "" || Slot_Box.Text == "" || OnuID_Box.Text == "")
                {
                    Cvlan_Box.Text = "尚未填入完整Port位";
                }

            }

        }

        private void Port_Box_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "MN (固定IP)" && Port_Box.Text != "" && Slot_Box.Text != "" && OnuID_Box.Text != "")
            {
                Cvlan_Box.Text = ((((int.Parse(Slot_Box.Text) - 1) * 4 + int.Parse(Port_Box.Text)) - 1) * 32 + (int.Parse(OnuID_Box.Text) - 1) + 1000).ToString();
            }
            if (comboBox1.Text == "MN (固定IP)")
            {
                if (Port_Box.Text == "" || Slot_Box.Text == "" || OnuID_Box.Text == "")
                {
                    Cvlan_Box.Text = "尚未填入完整Port位";
                }

            }

            if (Port_Box.Text != "")
            {
                if (int.Parse(Port_Box.Text) > 4)
                {
                    WMPLib.WindowsMediaPlayer wplayer = new WMPLib.WindowsMediaPlayer();
                    wplayer.URL = $@"{AppDomain.CurrentDomain.BaseDirectory}DATA/MIC02.mp3";
                    wplayer.controls.play();
                    MessageBox.Show("哇賽 你的OLT 超過4個Portㄝ");
                }
            }
        }

        private void OnuID_Box_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "MN (固定IP)" && Port_Box.Text != "" && Slot_Box.Text != "" && OnuID_Box.Text != "")
            {
                Cvlan_Box.Text = ((((int.Parse(Slot_Box.Text) - 1) * 4 + int.Parse(Port_Box.Text)) - 1) * 32 + (int.Parse(OnuID_Box.Text) - 1) + 1000).ToString();
            }
            if (comboBox1.Text == "MN (固定IP)")
            {
                if (Port_Box.Text == "" || Slot_Box.Text == "" || OnuID_Box.Text == "")
                {
                    Cvlan_Box.Text = "尚未填入完整Port位";
                }

            }
        }












        // ====================================VIP按鈕====================================

        private void VIP_button_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "zong")
            {
                MessageBox.Show("Right!!");

                int Desktop = 0;
                Desktop = SystemParametersInfo(20, 1, $@"{AppDomain.CurrentDomain.BaseDirectory}DATA/PIC01.jpg");






            }
            else if (textBox1.Text == "yuqin")
            {
                WMPLib.WindowsMediaPlayer wplayer = new WMPLib.WindowsMediaPlayer();
                wplayer.URL = $@"{AppDomain.CurrentDomain.BaseDirectory}DATA/MIC04.mp3";
                wplayer.controls.play();
                MessageBox.Show("恭喜開通：一鍵開通功能!");
                auto_button.Visible = true;
                richTextBox2.Visible = true;
            }


            else
            {
                WMPLib.WindowsMediaPlayer wplayer = new WMPLib.WindowsMediaPlayer();
                wplayer.URL = $@"{AppDomain.CurrentDomain.BaseDirectory}DATA/MIC01.mp3";
                wplayer.controls.play();
                MessageBox.Show("密碼提示：每個同事的英文名字都可能有不同效果");


            }
        }


        // ====================================輸出txt按鈕====================================
        private void txt_button_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            StreamWriter str1 = new StreamWriter(path.SelectedPath + "//Profile輸出.txt");
            str1.WriteLine(richTextBox1.Text);
            str1.Close();
            MessageBox.Show("輸出完畢!");

        }


        // ====================================輸入限制====================================
        private void Port_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void UP_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void Down_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void Percen_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void Svlan_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void Cvlan_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void Mode_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void Slot_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void OnuID_Box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }




        // ====================================開通按鈕====================================
        private void auto_button_Click(object sender, EventArgs e)
        {

            OLT_DIC OLTIP = new OLT_DIC();


            using (var client = new SshClient(OLTIP.Find_DIC(comboBox3.Text), 22, "admin", "123"))
            {
                // 建立連線
                client.Connect();

                // 連線參數
                var stream = client.CreateShellStream("", 0, 0, 0, 0, 0);

                string GP =
                                        $@"
en
config t
gp
"
;

                string GP_Port =
                    $@"
gp {ONT_Profile.Slot}/{ONT_Profile.Port}
"
;



                Thread.Sleep(5000);
                stream.WriteLine(GP);
                Thread.Sleep(2000);
                stream.WriteLine(ONT_Profile.Dba_profile());
                Thread.Sleep(2000);
                stream.WriteLine(ONT_Profile.Vlan_profile());
                Thread.Sleep(2000);
                stream.WriteLine(ONT_Profile.Traffic_profile());
                Thread.Sleep(2000);
                stream.WriteLine(ONT_Profile.Onu_profile());
                Thread.Sleep(2000);
                stream.WriteLine(GP_Port);
                Thread.Sleep(2000);
                stream.WriteLine(ONT_Profile.GPON_OLT());
                Thread.Sleep(2000);
                stream.WriteLine("exit");
                stream.WriteLine("bridge");
                Thread.Sleep(2000);
                stream.WriteLine(ONT_Profile.OLT_Vlan());


                // 輸出結果
                string line;
                while ((line = stream.ReadLine(TimeSpan.FromSeconds(2))) != null)
                {
                    Console.WriteLine(line);
                }
                // 結束連線
                stream.Close();
                client.Disconnect();
                WMPLib.WindowsMediaPlayer wplayer = new WMPLib.WindowsMediaPlayer();
                wplayer.URL = $@"{AppDomain.CurrentDomain.BaseDirectory}DATA/MIC03.mp3";
                wplayer.controls.play();
                MessageBox.Show("開通完畢!");


            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            TD_richTextBox.Clear();
            DateTime d = DateTime.Now;
            //string t = textBox1.Text.ToString();
            string t = Strings.StrConv(TD_textbox.Text.ToString(), VbStrConv.Wide);
            int count = 0;
            DateTime yday = d.AddDays(1);//明天
            string yyday = yday.ToString("yyyyMMdd");   //明天轉字串
            string day = d.ToString("yyyyMMdd");    //今天轉字串
            string line;
            string[] SecrH = new string[500];
            string[] text1 = new string[5000];
            int num = 0;
            int Dnum = 0;
            int cos = 0;
            int Tcos = 0;
            text1[0] = " ";
            string Sameday = d.GetDateTimeFormats('D')[1].ToString();
            foreach (var item in TD_area)
            {

                try
                {

                    StreamReader str = new StreamReader($@"{AppDomain.CurrentDomain.BaseDirectory}\{item}\{Sameday}-{item}.txt");
                    while ((line = str.ReadLine()) != null)
                    {
                        //  Console.WriteLine(line);
                        //text.Add(line);
                        if (num == 0 || num < 5000)
                        {
                            text1[num] = line;


                            int dayd = line.IndexOf(yyday); // 文本中搜尋明天

                            int ttd = line.IndexOf(day);//文本搜尋今天
                            int td = line.IndexOf("日期:");
                            int time = line.IndexOf("分 至");

                            int r = line.IndexOf(t);


                            if (ttd != -1)           //當搜尋到今天日期 count=1
                            {
                                count = 1;
                                //richTextBox2.Text += $"{text1[num]}" + Environment.NewLine;
                                if (Tcos == 0)
                                {
                                    TD_richTextBox.Text += $"{DateTime.Now.ToString("yyyy/MM/dd")}" + Environment.NewLine;
                                    Tcos = 1;
                                }
                            }
                            else if (td != -1)
                            {
                                count = 0;
                            }
                            // Console.WriteLine(line);

                            if (time != -1)
                            {
                                SecrH[Dnum] = text1[num];
                                Dnum++;
                                cos = 1;
                            }


                            switch (count)
                            {
                                case 1:
                                    if (r != -1)  //當搜尋到"日期"  "文字格中的字串" "自"  和count=1
                                    {                                             //輸出



                                        //richTextBox2.Text += $"{text1[num]}" + Environment.NewLine; 顯示區域
                                        if (cos == 1)
                                        {
                                            TD_richTextBox.Text += $"===============================" +
                                                    $"============================================" +
                                                    $"==========================================" +
                                                    $"\n{SecrH[Dnum - 1]}" + Environment.NewLine;

                                        }

                                        TD_richTextBox.Text += $"{text1[num]}" + Environment.NewLine;


                                        cos = 0;
                                    }

                                    break;
                                default:
                                    break;
                            }

                            num++;
                        }

                    }
                    str.Close();

                }
                catch (Exception)
                {

                    MessageBox.Show("檔案遺失，程式將重啟");
                    Application.Restart();
                }



            }
        }







        //========================================自動開通========================================






















    }
}
