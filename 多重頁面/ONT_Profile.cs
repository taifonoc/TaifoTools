using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 多重頁面
{
    class ONT_Profile
    {
        public static string Area, Slot, Port, Onuid, Sn, Des, Mode, Svlan, Cvlan, Bw_Up, Bw_Down, Ip, Gw, Bw_Per, MV;

        public static string P1, P2, P3, P4, OutArea;

        public static int Group;

        public static void Group7750()
        {
            if (Area == "104" || Area == "105" || Area == "110" || Area == "114" || Area == "115" || Area == "116")
            {
                Group = 1;
            }
            else
            {
                Group = 2;
            }
        }







        public static string mode
        {
            get { return Mode; }
            set
            {
                switch (value)
                {
                    case "0":
                        Mode = "0n4b";
                        P1 = "mgmt-mode uni eth 1 omci";
                        P2 = "mgmt-mode uni eth 2 omci";
                        P3 = "mgmt-mode uni eth 3 omci";
                        P4 = "mgmt-mode uni eth 4 omci";
                        break;
                    case "1":
                        Mode = "1n3b";
                        P1 = "mgmt-mode uni eth 1 non-omci link virtual-eth 1";
                        P2 = "mgmt-mode uni eth 2 omci";
                        P3 = "mgmt-mode uni eth 3 omci";
                        P4 = "mgmt-mode uni eth 4 omci";
                        break;
                    case "2":
                        Mode = "2n2b";
                        P1 = "mgmt-mode uni eth 1 non-omci link virtual-eth 1";
                        P2 = "mgmt-mode uni eth 2 non-omci link virtual-eth 1";
                        P3 = "mgmt-mode uni eth 3 omci";
                        P4 = "mgmt-mode uni eth 4 omci";
                        break;
                    case "3":
                        Mode = "3n1b";
                        P1 = "mgmt-mode uni eth 1 non-omci link virtual-eth 1";
                        P2 = "mgmt-mode uni eth 2 non-omci link virtual-eth 1";
                        P3 = "mgmt-mode uni eth 3 non-omci link virtual-eth 1";
                        P4 = "mgmt-mode uni eth 4 omci";
                        break;
                    case "4":
                        Mode = "4n0b";
                        P1 = "mgmt-mode uni eth 1 non-omci link virtual-eth 1";
                        P2 = "mgmt-mode uni eth 2 non-omci link virtual-eth 1";
                        P3 = "mgmt-mode uni eth 3 non-omci link virtual-eth 1";
                        P4 = "mgmt-mode uni eth 4 non-omci link virtual-eth 1";
                        break;
                }
            }
        }
        //mgmt-mode uni eth 1 non-omci link virtual-eth 1
        //mgmt-mode uni eth 3 omci

        public static string Dba_profile()
        {
            string dba;

            if (Bw_Up == "" || Bw_Per == "")
            {
                dba = "dab-profile創建失敗：頻寬欄位-少配置\r\n";
                MessageBox.Show(dba);
            }
            else
            {
                string Bw_end = (((int.Parse(Bw_Up) * (1 + double.Parse(Bw_Per) / 100)) * 1024) / 64).ToString("0");
                dba =
    $@"
dba-profile m{Bw_Up}_{Bw_Per}% create
    mode sr 
    sla maximum {int.Parse(Bw_end) * 64}
    apply
 !
";
            }
            return dba;

        }



        public static string Vlan_profile()
        {
            string vlan_pro;

            if (Svlan == "")
            {
                vlan_pro = "extended-vlan創建失敗：Svlan 欄位-少配置\r\n";
                MessageBox.Show(vlan_pro);
            }
            else
            {
                vlan_pro =
    $@"
extended-vlan-tagging-operation {Svlan}-1 create
    downstream-mode enable
    untagged-frame 1 
        treat outer vid {Svlan} cos 0 tpid 0x8100 
        treat inner vid 1 cos 0 tpid 0x8100 
    apply 
!
";
            }
            return vlan_pro;
        }



        public static string Traffic_profile()
        {
            string Sname;



            string traffic;
            if (
                Svlan == "" ||
                Mode == "" ||
                Bw_Up == "" ||
                Bw_Per == "" ||
                Bw_Down == "" ||
                P1 == "" ||
                P2 == "" ||
                P3 == "" ||
                P4 == ""
                )
            {
                traffic = "traffic-profile創建失敗：Svlan / Port Mode / 頻寬 其中有欄位-少配置\r\n";
                MessageBox.Show(traffic);
            }

            else
            {
                if (int.Parse(Svlan) >= 1000)
                {
                    Sname = "MN";
                }
                else
                {
                    Sname = "S" + Svlan;
                }
                string Bw_end1 = (((int.Parse(Bw_Up) * (1 + double.Parse(Bw_Per) / 100)) * 1024) / 64).ToString("0");
                string Bw_end2 = (((int.Parse(Bw_Down) * (1 + double.Parse(Bw_Per) / 100)) * 1024) / 64).ToString("0");
                traffic =
$@"
traffic-profile {Sname}_{Mode}_{Bw_Up} create
 {P1}
 {P2}
 {P3}
 {P4}
  tcont 1 
   gemport 1/1 
   dba-profile m{Bw_Up}_{Bw_Per}%
  tcont 2 
   gemport 2/1 
   dba-profile GVm512k
  mapper 1 
   gemport count 1 
   gemport 1 rate-limit upstream {int.Parse(Bw_end1) * 64} {int.Parse(Bw_end1) * 64}
   gemport 1 rate-limit downstream {int.Parse(Bw_end2) * 64} {int.Parse(Bw_end2) * 64}
  mapper 2 
   gemport count 1 
  bridge 1 
   ani mapper 1 
   uni eth 1 
    extended-vlan-tagging-operation {Svlan}-1 
   uni eth 2 
    extended-vlan-tagging-operation {Svlan}-1 
   uni eth 3 
    extended-vlan-tagging-operation {Svlan}-1  
   uni eth 4 
    extended-vlan-tagging-operation {Svlan}-1 
  bridge 2 
   ani mapper 2 
   link ip-host-config 1
  ip-host-config 1
   ip address static
   extended-vlan-tagging-operation 11 
  ip-host-config 2
   ip address static
  apply
 !

";
            }
            return traffic;
        }


        public static string Onu_profile()
        {
            string onu_pro;
            if (
               Svlan == "" ||
               Mode == "" ||
               Bw_Up == ""
               )
            {
                onu_pro = "onu-profile創建失敗：Svlan / Port Mode / 頻寬 其中有欄位-少配置\r\n";
                MessageBox.Show(onu_pro);
            }
            else
            {

                onu_pro =
    $@"
 onu-profile S{Svlan}_{Mode}_{Bw_Up} create
  traffic-profile S{Svlan}_{Mode}_{Bw_Up}
  loop-detect enable
  loop-detect block
  apply
 !

";
            }
            return onu_pro;
        }




















        public static string GPON_OLT()
        {
            string gpon_olt;

            if (
            Slot == "" ||
             Port == "" ||
             Onuid == "" ||
             Sn == "" ||
             Svlan == "" ||
             Mode == "" ||
             Bw_Up == "" ||
             Bw_Per == "" ||
             Bw_Down == "" ||
             Cvlan == ""


             )
            {
                gpon_olt = "GPON-OLT創建失敗：你幾乎都不填，是想怎樣?!\r\n";
                MessageBox.Show(gpon_olt);
            }
            else
            {


                if (Mode == "4n0b")
                {
                    if (Ip == "" || Gw == "")
                    {
                        gpon_olt = "GPON-OLT創建失敗：IP / GW 其中有欄位-少配置\r\n";
                        MessageBox.Show(gpon_olt);
                    }
                    else
                    {
                        gpon_olt =
        $@"
onu add {Onuid} {Sn} auto-learning
onu description {Onuid} {Des}
onu-profile {Onuid} S{Svlan}_0n4b_{Bw_Up}
{"onu static-ip " + Onuid + " " + "ip-host 1 10.255." + ((int.Parse(Slot) - 1) * 4 + (1 * (int.Parse(Port) - 1))).ToString("0") + "." + Onuid + "/22 gw 10.255." + (int.Parse(Slot) - 1) * 4 + "." + "254"}
onu static-ip {Onuid} ip-host 2 {Ip} gw {Gw}
onu extended-vlan {Onuid} {Svlan}-1 untagged-frame 1 treat inner vid {Cvlan} cos 0
{"onu extended-vlan" + " " + Onuid + " " + "11" + " " + "untagged-frame 1 treat inner vid " + 10 + int.Parse(Slot) + " cos 0"}
";
                    }
                }
                else
                {
                    gpon_olt =
    $@"
onu add {Onuid} {Sn} auto-learning
onu description {Onuid} {Des}
{"onu static-ip " + Onuid + " " + "ip-host 1 10.255." + ((int.Parse(Slot) - 1) * 4 + (1 * (int.Parse(Port) - 1))).ToString("0") + "." + Onuid + "/22 gw 10.255." + (int.Parse(Slot) - 1) * 4 + "." + "254"}
onu-profile {Onuid} S{Svlan}_0n4b_{Bw_Up}
onu extended-vlan {Onuid} {Svlan}-1 untagged-frame 1 treat inner vid {Cvlan} cos 0
{"onu extended-vlan" + " " + Onuid + " " + "11" + " " + "untagged-frame 1 treat inner vid " + 10 + int.Parse(Slot) + " cos 0"}
";

                }
            }
            return gpon_olt;
        }



        public static string OLT_Vlan()
        {
            string olt_vlan;

            if (Svlan == "" || Slot == "" || Port == "")
            {
                olt_vlan = "OLT Tag創建失敗：Slot / Port /Svlan 欄位-少配置\r\n";
                MessageBox.Show(olt_vlan);
            }
            else
            {
                olt_vlan =
    $@"
vlan add {Svlan} {Slot}/{Port} tagged 
vlan add {Svlan} t/1 tagged
";
            }
            return olt_vlan;




        }



        public static string K10()
        {
            string k10;

            if (Svlan == "")
            {
                k10 = "10K創建失敗：Svlan 欄位-少配置\r\n";
                MessageBox.Show(k10);
            }
            else
            {
                k10 =
    $@"
ethernet-service svlan {Svlan}

";
            }
            return k10;
        }




        public static string ALU7750()
        {
            string alu7750 = "";

            if (Svlan == "" || Cvlan == "" || Area == "")
            {
                alu7750 = "Alu7750 創建失敗：Svlan / Cvlan / OLT 欄位-少配置\r\n";
                MessageBox.Show(alu7750);
            }
            else if (Svlan == "602" || Svlan == "610")
            {
                if (Group == 1)
                {
                    switch (Svlan)
                    {
                        case "602":
                            OutArea = "110";
                            break;
                        case "610":
                            OutArea = "110";
                            break;
                        default:
                            break;
                    }

                    alu7750 =
$@"
內湖7750-01：
vpls {Svlan}{Cvlan} customer 1 create

description ""MV{MV}-{Svlan}{Cvlan}""
service-mtu 9194
stp
    no shutdown
exit
service-name ""MV{MV}-{Svlan}{Cvlan}""
sap lag-{Area}:{Svlan}.{Cvlan} create
exit
sap lag-{OutArea}:{Svlan}.{Cvlan} create
exit
spoke-sdp 10801:{Svlan}{Cvlan} create
      no shutdown
  exit
no shutdown

萬華7750-01：
vpls {Svlan}{Cvlan} customer 1 create

description ""MV{MV}-{Svlan}{Cvlan}""
service-mtu 9194
stp
    no shutdown
exit
service-name MV{MV}-{Svlan}{Cvlan}
sap lag-{Area}:{Svlan}.{Cvlan} create
exit
sap lag-{OutArea}:{Svlan}.{Cvlan} create
exit
spoke-sdp 10801:{Svlan}{Cvlan} create
      no shutdown
  exit
no shutdown



";

                }
                else
                {
                    alu7750 =
$@"
內湖7750-02：
vpls {Svlan}{Cvlan} customer 1 create

description ""MV{MV}-{Svlan}{Cvlan}""
service-mtu 9194
stp
    no shutdown
exit
service-name ""MV{MV}-{Svlan}{Cvlan}""
sap lag-{Area}:{Svlan}.{Cvlan} create
exit
spoke-sdp 10802:{Svlan}{Cvlan} create
exit
spoke-sdp 11401:{Svlan}{Cvlan} create
      no shutdown
  exit
no shutdown

萬華7750-02：
vpls {Svlan}{Cvlan} customer 1 create

description ""MV{MV}-{Svlan}{Cvlan}""
service-mtu 9194
stp
    no shutdown
exit
service-name ""MV{MV}-{Svlan}{Cvlan}""
sap lag-{Area}:{Svlan}.{Cvlan} create
exit
spoke-sdp 10801:{Svlan}{Cvlan} create
exit
spoke-sdp 11402:{Svlan}{Cvlan} create
      no shutdown
  exit
no shutdown

內湖7750-01：
vpls {Svlan}{Cvlan} customer 1 create

description ""MV{MV}-{Svlan}{Cvlan}""
service - mtu 9194
stp
    no shutdown
exit
service-name ""MV{MV}-{Svlan}{Cvlan}""
sap lag-{OutArea}:{Svlan}.{Cvlan} create
exit
spoke-sdp 10801:{Svlan}{Cvlan} create
exit
spoke-sdp 11402:{Svlan}{Cvlan} create
      no shutdown
  exit
no shutdown

萬華7750-01：
vpls {Svlan}{Cvlan} customer 1 create

description ""MV{MV}-{Svlan}{Cvlan}""
service-mtu 9194
stp
    no shutdown
exit
service-name ""MV{MV}-{Svlan}{Cvlan}""
sap lag-{OutArea}:{Svlan}.{Cvlan} create
exit
spoke-sdp 10802:{Svlan}{Cvlan} create
exit
spoke-sdp 11401:{Svlan}{Cvlan} create
      no shutdown
  exit
no shutdown


";







                }



            }




            return alu7750;
        }


    }
}
