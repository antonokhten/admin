using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Security.AccessControl;
//using Ionic.Zip;
using MetroFramework.Components;
using MetroFramework.Forms;
using MetroFramework.Animation;
using MetroFramework.Controls;
using MetroFramework.Drawing;
using MetroFramework.Fonts;
using Sipek.Common;
using Sipek.Common.CallControl;
using Sipek.Sip;
using System.Management.Automation;

//using System.Management.automation;

namespace admin
{
    public partial class Form1 : MetroForm
    {
        #region properties

        CCallManager CallManager
        {
            get
            {
                return CCallManager.Instance;
            }
        }

        private rc_PhoneCfg v_hPhoneCfg = new rc_PhoneCfg();
        internal rc_PhoneCfg Config
        {
            get
            {
                return v_hPhoneCfg;
            }
        }

        private IStateMachine v_hCall = null;
        private IStateMachine v_hIncomingCall = null;
        #endregion
        //  string dmwrPath = @"C:/Program Files (x86)/SolarWinds/DameWare Remote Support";
        string dmwrPath = Properties.Settings.Default.dmwrPath;
        string file_msc = Properties.Settings.Default.mmc;
        string msinfo_xml = Properties.Settings.Default.msinfo_xml;
        string acronis_xml = Properties.Settings.Default.acronis_xml;
        string PsLogon = Properties.Settings.Default.PsLogon;
        int call_status = Properties.Settings.Default.call_status;
        int servise_int = Properties.Settings.Default.servise_int;
        int check_roll_dw = Properties.Settings.Default.check_roll_dw;
        // system32\msinfo32.exe  check_roll_dw
       

        System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"c:\mywavfile.wav");


        public void ps_user_info()
        {
            string s;



            //Clipboard.SetText();
            //Company,CanonicalName,telephoneNumber,whenCreated,LastLogonDate,HomeDirectory
            //Company,CanonicalName,telephoneNumber,whenCreated,LastLogonDate,HomeDirectory
            //Company,CanonicalName,telephoneNumber,whenCreated,LastLogonDate,HomeDirectory
            try
            {
                s = listBox1.SelectedItem.ToString();
                Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe ", @"-noexit -command import-module ActiveDirectory; Get-ADUser -identity " + s + " -properties * ; Get-ADUser " + s + " -Properties Memberof | Select -ExpandProperty memberOf;" + @" Read-Host -Prompt" + @"""Press Enter to exit""");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ps err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ps_logon_user()
        {
            string s;
            s = PsLogon;
            try
            {

                //  Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe ", @"-noexit " + s );
                Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe", s);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ps err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void ps_run(string commmand)
        {
            string s;



            //Clipboard.SetText();
            try
            {
                s = listBox1.SelectedItem.ToString();
                if (commmand == "addscan")
                {
                    Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe ", @" -command import-module ActiveDirectory; Add-ADGroupMember SCAN " + s);
                }
                if (commmand == "addpu")
                {
                    Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe ", @" -command import-module ActiveDirectory; Add-ADGroupMember PU " + s);
                }

            }

            catch (Exception)
            {
                MessageBox.Show(@"error", "ps err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string Tr2(string s)
        {
            StringBuilder ret = new StringBuilder();
            string[] rus = { " ", "А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й",
          "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц",
          "Ч", "Ш", "Щ", "Ъ", "Ы", "Ь", "Э", "Ю", "Я" };
            string[] eng = { " ", "A", "B", "V", "G", "D", "E", "E", "ZH", "Z", "I", "I",
          "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "KH", "TS",
          "CH", "SH", "SHCH", "IE", "Y", null, "E", "IU", "IA" };

            for (int j = 0; j < s.Length; j++)
                for (int i = 0; i < rus.Length; i++)
                    if (s.Substring(j, 1) == rus[i]) ret.Append(eng[i]);
            return ret.ToString();
        }


        public Form1()
        {
            InitializeComponent();
            CallManager.CallStateRefresh += new DCallStateRefresh(CallManager_CallStateRefresh);
            CallManager.IncomingCallNotification += new DIncomingCallNotification(CallManager_IncomingCallNotification);
            pjsipRegistrar.Instance.AccountStateChanged += new DAccountStateChanged(Instance_AccountStateChanged);

            CallManager.StackProxy = pjsipStackProxy.Instance;

            CallManager.Config = Config;
            pjsipStackProxy.Instance.Config = Config;
            pjsipRegistrar.Instance.Config = Config;

            CallManager.Initialize();

            pjsipRegistrar.Instance.registerAccounts();
            RTB_J83.Text = Properties.Settings.Default.j83_label;


            metroTile58.Text = Properties.Settings.Default.phone1_label;
            metroTile59.Text = Properties.Settings.Default.phone2_label;
            metroTile60.Text = Properties.Settings.Default.phone3_label;
            metroTile97.Text = Properties.Settings.Default.phone4_label;

            metroTile94.Text = Properties.Settings.Default.phone5_label;
            metroTile95.Text = Properties.Settings.Default.phone6_label;
            metroTile96.Text = Properties.Settings.Default.phone7_label;
            metroTile100.Text = Properties.Settings.Default.phone8_label;
            metroTile102.Text = Properties.Settings.Default.phone9_label;
            metroTile99.Text = Properties.Settings.Default.phone10_label;
        }

        #region callbacks
        void Instance_AccountStateChanged(Int32 iAccountId, Int32 iAccState)
        {
            if (InvokeRequired)
                this.BeginInvoke(new DAccountStateChanged(OnRegistrationUpdate), new Object[] { iAccountId, iAccState });
            else
                OnRegistrationUpdate(iAccountId, iAccState);
        }

        void CallManager_CallStateRefresh(Int32 iSessionId)
        {
            if (InvokeRequired)
                this.BeginInvoke(new DCallStateRefresh(OnStateUpdate), new Object[] { iSessionId });
            else
                OnStateUpdate(iSessionId);
        }

        void CallManager_IncomingCallNotification(Int32 iSessionId, String szNumber, String szInfo)
        {
            if (InvokeRequired)
                this.BeginInvoke(new DIncomingCallNotification(OnIncomingCallNotification), new Object[] { iSessionId, szNumber, szInfo });
            else
                OnIncomingCallNotification(iSessionId, szNumber, szInfo);
        }
        #endregion

        #region synchronized callbacks
        private void OnRegistrationUpdate(Int32 iAccountId, Int32 iAccState)
        {
            regstate.Text = "regstate:" +iAccState.ToString();
        }

        private void OnStateUpdate(Int32 iSessionId)
        {
            cs_CallState.Text = CallManager.getCall(iSessionId).StateId.ToString();
            
           
        }

        private void OnIncomingCallNotification(Int32 iSessionId, String szNumber, String szInfo)
        {

            v_hIncomingCall = CallManager.getCall(iSessionId);
            cs_CallState.Text = v_hIncomingCall.StateId.ToString();
            //-----
            cs_Callnomber.Text = v_hIncomingCall.CallingNumber.ToString();
            callstate.Text = "callstate:" + v_hIncomingCall.StateId.ToString();
            //------
            metroTile85.Text = "ОТВЕТИТЬ...   " + v_hIncomingCall.CallingNumber.ToString();
            textBox3.Text = v_hIncomingCall.CallingName;

            textBox3.Text = szInfo;


        }
        #endregion

        public void download_log()
        {
            label2.Text = "";

            listBox2.Items.Clear();
            int i;
            string fio;

            if (listBox1.SelectedItems.Count == 1)
            {

                try
                {
                    i = listBox1.SelectedIndex;
                    fio = listBox1.SelectedItem.ToString();
                    string path = @"\\guo.local\dfssvc\_loginlogNSK\userlog\" + fio + ".txt";

                    string[] arStr = File.ReadAllLines(path, Encoding.GetEncoding(1251));
                    listBox2.Items.AddRange(arStr);
                    listBox2.SelectedIndex = listBox2.Items.Count - 1;


                }
                catch (Exception)
                {
                    MessageBox.Show(@"error, wrong path \\guo.local\dfssvc\_loginlogNSK\userlog\", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            else
            {
                MessageBox.Show(@"выбрано больше одного", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
        public void dmware_start()
        {
            string ncomp, s, ss;
            char ck;
            int i_s;
            ncomp = "";
            try
            {
                ncomp = listBox2.SelectedItem.ToString();
                char c = ncomp[0];
                if (c == 'В')
                {
                    s = ncomp.Remove(0, 5);
                }
                else
                {
                    s = ncomp.Remove(0, 10);
                }

                for (int i = 0; i < s.Length; i++)
                {
                    ck = s[i];
                    if (ck == ',')
                    {
                        i_s = i;
                        ss = s.Remove(i, s.Length - i);
                        label2.Text = ss;
                        Process.Start(dmwrPath, "-c: -h -m:" + ss + " -a:1 -x:");
                        break;
                    }




                    //Form1.ActiveForm.WindowState = FormWindowState.Minimized;

                }

            }
            catch (Exception)
            {
                MessageBox.Show(@"error selected ", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
        

        /*  public void AddUserToGroup(string userId, string groupName)
          {
              try
              {
                  using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, "COMPANY"))
                  {
                      GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, groupName);
                      group.Members.Add(pc, IdentityType.UserPrincipalName, userId);
                      group.Save();
                  }
              }
              catch (System.DirectoryServices.DirectoryServicesCOMException E)
              {
                  //doSomething with E.Message.ToString(); 

              }
          }

          public void RemoveUserFromGroup(string userId, string groupName)
          {
              try
              {
                  using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, "COMPANY"))
                  {
                      GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, groupName);
                      group.Members.Remove(pc, IdentityType.UserPrincipalName, userId);
                      group.Save();
                  }
              }
              catch (System.DirectoryServices.DirectoryServicesCOMException E)
              {
                  //doSomething with E.Message.ToString(); 

              }
          } */
        public static void AddDomainUserToDomainGroup(string userName, string groupName)
        {
            //  string domainDNS = Domain.GetComputerDomain().ToString();
            // var root = Domain.GetComputerDomain().GetDirectoryEntry();
            //  root = root.Children.Find("OU=GROUPS");
            //  var group = root.Children.Find(string.Format("CN={0}","CN = SCAN, OU = GROUPS, DC = sib, DC = local"), "group");  //groupName
            // root = root.Children.Find("OU=SIB");
            //var user = root.Children.Find(string.Format("CN={0}", @"LDAP://DC=sib, DC=local,CN=Пертулисов Данил Викторович,OU=PTO,OU=DOM_TELEKOM,OU=SIB"), "user");


            //      DirectoryEntry entry = new DirectoryEntry(@"LDAP://sib.local/cn=SCAN, OU=GROUPS, DC=SIB, DC=local");
            //    //LDAP://child.domain.com/cn=group,ou=sample,dc=child,dc=domain,dc=com
            // entry.Path = @"LDAP://DC=sib, DC=local, OU=GROUPS, CN=SCAN";
            //      entry.AuthenticationType = AuthenticationTypes.Secure;

            //DirectorySearcher search_group = new DirectorySearcher(entry);
            //  search_group.Filter = (@"CN=SCAN");
            //  DirectorySearcher search_user = new DirectorySearcher(entry);
            //  search_user.Filter = (@"SAMAccountName=pertulisov");

            //  search_group.p
            //DC=sib, DC=local, OU=SIB, OU=DOM_TELEKOM, OU=PTO,

            //  entry.Properties["member"].Add(@"LDAP://DC=sib, DC=local, CN=Пертулисов Данил Викторович, OU=SIB, OU=DOM_TELEKOM, OU=PTO, ");
            // entry.Properties["member"].Add(@"LDAP://sib.local/SAMAccountName=pertulisov, OU=SIB, OU=DOM_TELEKOM, OU=PTO,  DC=SIB, DC=local");
            //        entry.CommitChanges();
            //  entry.Close();

            // group.Properties["member"].Add("LDAP://DC=sib, DC=local,CN=Пертулисов Данил Викторович,OU=PTO,OU=DOM_TELEKOM,OU=SIB");
            //    group.CommitChanges();
        }
        public static string GetUserProperty(string accountName, string propertyName)
        {
            DirectoryEntry entry = new DirectoryEntry();

            // "LDAP://CN=<group name>, CN =<Users>, DC=<domain component>, DC=<domain component>,..."
            entry.Path = @"LDAP://DC=sib, DC=local";
            entry.AuthenticationType = AuthenticationTypes.Secure;

            DirectorySearcher search = new DirectorySearcher(entry);
            search.Filter = "(SAMAccountName=" + accountName + ")";
            search.PropertiesToLoad.Add(propertyName);
            try
            {
                SearchResultCollection results = search.FindAll();
                if (results != null && results.Count > 0)
                {
                    return results[0].Properties[propertyName][0].ToString();


                }
                else
                {
                    return "Unknown";
                }
            }
            catch (Exception)
            {
                return "Unknown";
            }
        }

        void get_user_info()
        {
            string item1_user, s_tel, s_mail, s_pach, s_cn, s_comp, s_dol;
            try
            {
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                s_tel = "-";
                s_mail = "-";
                s_pach = "-";
                s_cn = "-";
                s_comp = "-";
                s_dol = "-";
                item1_user = listBox1.SelectedItem.ToString();
                s_tel = GetUserProperty(item1_user, "telephoneNumber");
                s_mail = GetUserProperty(item1_user, "mail");
                s_pach = GetUserProperty(item1_user, "HomeDirectory");
                s_cn = GetUserProperty(item1_user, "CanonicalName");
                s_comp = GetUserProperty(item1_user, "Company");
                s_dol = GetUserProperty(item1_user, "Title");
                textBox4.Text = s_tel;
                textBox5.Text = s_mail;
                textBox6.Text = s_pach;
                textBox7.Text = s_cn;
                textBox8.Text = s_comp;
                textBox9.Text = s_dol;
                
            }
            catch (Exception)
            {
                MessageBox.Show(@"не выбран юзер", "get_user_info", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        


        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            //  Process.Start("cmd.exe", "/C " + "ping 8.8.8.8");
            // Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:"+s+" -a:1 -x:");
            // Form1.ActiveForm.WindowState = FormWindowState.Minimized;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-dzr-j83k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j76k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-dzr-j31k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-dzr-j82k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-dzr-j22k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-dzr-j111k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-dzr-j2k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j6k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button9_Click(object sender, EventArgs e)
        {

            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j1k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j1k3 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j10vk1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-j10k-nar -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j4k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-dzr-j44k2 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j7k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j8k -a:1 -x:");
        }

        private void button25_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-reu2-k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j9k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j35k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j16k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j3k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-do-kas -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button27_Click(object sender, EventArgs e)
        {

        }

        private void button24_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-pers-j25k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-perj-j73k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j6k2 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j7k4 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button29_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-c: -h -m:sib-mks-j35k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button31_Click(object sender, EventArgs e)
        {

            // string compPatch = @"\\sib-dzr-j111k1\c$";
            Process.Start("explorer.exe ", @" \\sib-dzr-j111k1\c$");



        }

        private void button38_Click(object sender, EventArgs e)
        {

        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {

        }

        private void button45_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath, "-h -m:10.10.0. -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button47_Click(object sender, EventArgs e)
        {

        }

        private void button48_Click(object sender, EventArgs e)
        {

        }

        private void button49_Click(object sender, EventArgs e)
        {
            // string s;

            //   s = listBox1.SelectedValue.ToString();

            Process.Start("cmd.exe", @"-ping -t 8.8.8.8");
            // получаем id выделенного объекта
            // int id = (int)listBox1.SelectedValue;
            //   string phone;
            // получаем весь выделенный объект
            //  phone = listBox1.SelectedItem.ToString();
            //  MessageBox.Show(id.ToString() + ". " + phone);
            //label8.Text = listBox1.SelectedValue.ToString();
        }

        /*  private void Form1_Load(object sender, EventArgs e)
          {
              // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet3.Table1". При необходимости она может быть перемещена или удалена.

              this.table1TableAdapter.Fill(this.db_adminkaDataSet3.Table1);
              // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet1.Table1". При необходимости она может быть перемещена или удалена.
              //this.table1TableAdapter1.Fill(this.db_adminkaDataSet1.Table1);
              // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet3.Table1". При необходимости она может быть перемещена или удалена.
              this.table1TableAdapter.Fill(this.db_adminkaDataSet3.Table1);
              // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet3.unit1". При необходимости она может быть перемещена или удалена.
              // this.unit1TableAdapter.Fill(this.db_adminkaDataSet3.unit1);
              // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet2.T_remoute_office". При необходимости она может быть перемещена или удалена.
              //this.t_remoute_officeTableAdapter1.Fill(this.db_adminkaDataSet2.T_remoute_office);
              // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet11.T_remoute_office". При необходимости она может быть перемещена или удалена.
              // this.t_remoute_officeTableAdapter.Fill(this.db_adminkaDataSet11.T_remoute_office);
              // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet.Table1". При необходимости она может быть перемещена или удалена.
              //this.table1TableAdapter.Fill(this.db_adminkaDataSet.Table1);

              //cn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\AsusK450c\documents\visual studio 2010\Projects\ADD\ADD\testing.accdb;Persist Security Info=True";
              ///cmd.Connection = cn;
              ///loaddata();
              //  //listBox2.DataSource = db_adminkaDataSet2TableAdapters.;
              //  listBox2.DisplayMember = "rem_of_name";
              //  listBox2.ValueMember = "rem_of_ID";

              //   UNIT1   load
              //label4.Text = Properties.Settings.Default.UNIT1_name;
              try
              {
                  /// label1.Text = UserDomainName; \\sibsup\_LoginLog        @"X:\userlog"
                  string[] files = System.IO.Directory.GetFiles(@"\\sibsup\_LoginLog\userlog", "*.txt", System.IO.SearchOption.AllDirectories);

                  for (int i = 0; i < files.Length; i++)
                  {
                      files[i] = files[i].Replace(@"\\sibsup\_LoginLog\userlog\", "");
                      files[i] = files[i].Replace(@".txt", "");
                  }

                  this.listBox1.Items.AddRange(files);
                  textBox2.Focus();
              }
              catch (Exception)
              {
                  MessageBox.Show(@"error, wrong path \\sibsup\_LoginLog\userlog", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
              }


          }*/

        private void listBox2_DoubleClick(object sender, EventArgs e)
        {
            //  listBox3.DataSource = dbadminkaDataSet1BindingSource1;
        }

        private void button47_Click_1(object sender, EventArgs e)
        {

        }

        private void metrobutton48_Click_1(object sender, EventArgs e)
        {

        }

        private void button46_Click(object sender, EventArgs e)
        {

        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void button54_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            // string s;

            //s = listBox1.SelectedValue.ToString();
            //  Process.Start("cmd.exe", "/C " + "ping 8.8.8.8");
            //Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:" + s + " -a:1 -x:");
            // Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button44_Click(object sender, EventArgs e)
        {

        }

        private void table1BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.table1BindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.db_adminkaDataSet3);
            //this.tableAdapterManager.
            table1TableAdapter.Connection.Close();
            table1TableAdapter.Connection.Close();

        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click_1(object sender, EventArgs e)
        {

        }

        private void button7_Click_1(object sender, EventArgs e)
        {

        }

        private void button9_Click_1(object sender, EventArgs e)
        {

        }

        private void button31_Click_1(object sender, EventArgs e)
        {

        }

        private void button32_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            int j = listBox1.Items.Count;

            listBox1.SelectionMode = SelectionMode.MultiExtended;
            listBox1.SelectedItems.Clear();
            if (textBox2.Text.Length > 1)
            {
                for (int i = 0; i < j; i++)
                {
                    if (listBox1.Items[i].ToString().ToLower().Contains(textBox2.Text.ToLower()))
                    {
                        listBox1.SetSelected(i, true);
                        // label2.Text = i.ToString();

                    }
                }
                string v;
                int iv;
                v = listBox1.SelectedIndices.Count.ToString();
                label1.Text = v;
                iv = listBox1.SelectedIndices.Count;
                if (v == "1")
                {
                    // textBox2.Text = "";
                    //listBox2.ClearSelected();
                    listBox2.Items.Clear();

                    download_log();
                    get_user_info();
                }

            }
            else
            {
                listBox1.SelectedItems.Clear();
                listBox2.Items.Clear();
                label1.Text = "";
                label2.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
            }
        }



        private void button1_Click_2(object sender, EventArgs e)
        {


        }

        private void button33_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e)
        {
        }

        private void button45_Click_1(object sender, EventArgs e)
        {

        }

        private void listBox2_CursorChanged(object sender, EventArgs e)
        {





            //Form1.ActiveForm.WindowState = FormWindowState.Minimized;

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
           
                string ncomp, s, ss;
                char ck;
                int i_s;
                ncomp = "";
                ncomp = listBox2.SelectedItem.ToString();
                char c = ncomp[0];
                if (c == 'В')
                {
                    s = ncomp.Remove(0, 5);
                }
                else
                {
                    s = ncomp.Remove(0, 10);
                }

                for (int i = 0; i < s.Length; i++)
                {
                    ck = s[i];
                    if (ck == ',')
                    {
                        i_s = i;
                        ss = s.Remove(i, s.Length - i);
                        label2.Text = ss;
                        // Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:" + ss + " -a:1 -x:");
                        break;
                    }
                }
            
        }

        private void button3_Click_1(object sender, EventArgs e)
        {

        }

        private void listBox2_DoubleClick_1(object sender, EventArgs e)
        {
            string ncomp, s, ss;
            char ck;
            int i_s;
            ncomp = "";
            try
            {
                ncomp = listBox2.SelectedItem.ToString();
                char c = ncomp[0];
                if (c == 'В')
                {
                    s = ncomp.Remove(0, 5);
                }
                else
                {
                    s = ncomp.Remove(0, 10);
                }

                for (int i = 0; i < s.Length; i++)
                {
                    ck = s[i];
                    if (ck == ',')
                    {
                        i_s = i;
                        ss = s.Remove(i, s.Length - i);
                        label2.Text = ss;
                        Process.Start(dmwrPath, "-c: -h -m:" + ss + " -a:1 -x:");
                        break;
                    }
                }
            }

            catch (Exception)
            {
                MessageBox.Show(@"error selected ", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {

        }

        private void button8_Click_1(object sender, EventArgs e)
        {

        }

        private void button10_Click_1(object sender, EventArgs e)
        {

        }

        private void button11_Click_1(object sender, EventArgs e)
        {


        }

        private void button12_Click_1(object sender, EventArgs e)
        {

        }

        private void button5_Click_2(object sender, EventArgs e)
        {
            //C: \Users\okhten\Desktop\putty.exe

        }

        private void button13_Click_1(object sender, EventArgs e)
        {

        }

        private void button14_Click_1(object sender, EventArgs e)
        {


        }

        private void button15_Click_1(object sender, EventArgs e)
        {

        }

        private void button16_Click_1(object sender, EventArgs e)
        {

        }

        private void button17_Click_1(object sender, EventArgs e)
        {
            //%windir%\system32\dhcpmgmt.msc
            //Process.Start(@"");

        }

        private void button18_Click_1(object sender, EventArgs e)
        {
            // Z:\Сертификаты\Дзержинец\Рабочие места_пользователи_Дзержинец.xlsx
            Process.Start("excel.exe", @"Z:\Сертификаты\Дзержинец\Рабочие_места_пользователи_Дзержинец.xlsx");
        }

        private void button19_Click_1(object sender, EventArgs e)
        {
            Process.Start("excel.exe", @"Z:\Сертификаты\МКС\Рабочие_места_пользователи_МКС.xlsx");
        }

        private void button20_Click_1(object sender, EventArgs e)
        {
            //"C:\Program Files (x86)\VMware\Infrastructure\Virtual Infrastructure Client\Launcher\VpxClient.exe"


        }

        private void button21_Click_1(object sender, EventArgs e)
        {



        }

        private void button22_Click_1(object sender, EventArgs e)
        {

        }

        private void button23_Click_1(object sender, EventArgs e)
        {

        }

        private void button24_Click_1(object sender, EventArgs e)
        {


        }

        private void button25_Click_1(object sender, EventArgs e)
        {

        }

        private void button28_Click_1(object sender, EventArgs e)
        {


        }

        private void button26_Click_1(object sender, EventArgs e)
        {

        }

        private void button29_Click_1(object sender, EventArgs e)
        {

        }

        private void button34_Click(object sender, EventArgs e)
        {


        }

        private void button35_Click(object sender, EventArgs e)
        {



        }

        private void button36_Click(object sender, EventArgs e)
        {



        }

        private void button37_Click(object sender, EventArgs e)
        {

        }

        private void button39_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedValueChanged(object sender, EventArgs e)
        {




        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            string v;
            int iv;
            v = listBox1.SelectedIndices.Count.ToString();
            label1.Text = v;
            iv = listBox1.SelectedIndices.Count;

            if (v == "1")
            {

                // textBox2.Text = "";
                //listBox2.ClearSelected();
                listBox2.Items.Clear();
                //   label2.Text = "";
                download_log();
                get_user_info();
            }

            else
            {
                listBox1.SelectedItems.Clear();
                listBox2.Items.Clear();
                label1.Text = "";

            }
        }

        private void button40_Click(object sender, EventArgs e)
        {

        }

        private void button30_Click_1(object sender, EventArgs e)
        {

        }

        private void button42_Click(object sender, EventArgs e)
        {
            //C:\Program Files\Internet Explorer\iexplore.exe

        }

        private void button41_Click(object sender, EventArgs e)
        {

        }

        private void button41_Click_1(object sender, EventArgs e)
        {
            //% SystemRoot %\system32\WindowsPowerShell\v1.0\powershell.exe

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet4.T_company". При необходимости она может быть перемещена или удалена.
            this.t_companyTableAdapter1.Fill(this.db_adminkaDataSet4.T_company);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet5.T_company". При необходимости она может быть перемещена или удалена.

            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet4.T_company_ad_info". При необходимости она может быть перемещена или удалена.
            //this.t_company_ad_infoTableAdapter.Fill(this.db_adminkaDataSet4.T_company_ad_info);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet4.T_company". При необходимости она может быть перемещена или удалена.
            // this.t_companyTableAdapter1.Fill(this.db_adminkaDataSet4.T_company);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet2.T_company". При необходимости она может быть перемещена или удалена.
            this.t_companyTableAdapter.Fill(this.db_adminkaDataSet2.T_company);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet3.Table1". При необходимости она может быть перемещена или удалена.

            this.table1TableAdapter.Fill(this.db_adminkaDataSet3.Table1);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet1.Table1". При необходимости она может быть перемещена или удалена.
            //this.table1TableAdapter1.Fill(this.db_adminkaDataSet1.Table1);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet3.Table1". При необходимости она может быть перемещена или удалена.
            this.table1TableAdapter.Fill(this.db_adminkaDataSet3.Table1);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet3.unit1". При необходимости она может быть перемещена или удалена.
            // this.unit1TableAdapter.Fill(this.db_adminkaDataSet3.unit1);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet2.T_remoute_office". При необходимости она может быть перемещена или удалена.
            //this.t_remoute_officeTableAdapter1.Fill(this.db_adminkaDataSet2.T_remoute_office);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet11.T_remoute_office". При необходимости она может быть перемещена или удалена.
            // this.t_remoute_officeTableAdapter.Fill(this.db_adminkaDataSet11.T_remoute_office);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "db_adminkaDataSet.Table1". При необходимости она может быть перемещена или удалена.
            //this.table1TableAdapter.Fill(this.db_adminkaDataSet.Table1);

            //cn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\AsusK450c\documents\visual studio 2010\Projects\ADD\ADD\testing.accdb;Persist Security Info=True";
            ///cmd.Connection = cn;
            ///loaddata();
            //  //listBox2.DataSource = db_adminkaDataSet2TableAdapters.;
            //  listBox2.DisplayMember = "rem_of_name";
            //  listBox2.ValueMember = "rem_of_ID";
            webBrowser1.ScriptErrorsSuppressed = true;
            //   UNIT1   load
            //label4.Text = Properties.Settings.Default.UNIT1_name;
            try
            {
                /// label1.Text = UserDomainName; \\sibsup\_LoginLog        @"X:\userlog"
                string[] files = System.IO.Directory.GetFiles(@"\\guo.local\dfssvc\_loginlogNSK\userlog\", "*.txt", System.IO.SearchOption.AllDirectories);

                for (int i = 0; i < files.Length; i++)
                {
                    files[i] = files[i].Replace(@"\\guo.local\dfssvc\_loginlogNSK\userlog\", "");
                    files[i] = files[i].Replace(@".txt", "");
                }

                this.listBox1.Items.AddRange(files);
                textBox2.Focus();
            }
            catch (Exception)
            {
                MessageBox.Show(@"error, wrong path \\guo.local\dfssvc\_loginlogNSK\userlog", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            // webBrowser2.Navigate("https://mail.regenergy.ru/owa");
            //  webBrowser2.ScriptErrorsSuppressed = true;
            //webBrowser3.Navigate("http://dcnag.guo.local/nagios/");
            // webBrowser3.ScriptErrorsSuppressed = true;

            //webBrowser4.Navigate("http://dcrmng.guo.local/projects/task_for_admin_nsk/issues");
            // webBrowser4.ScriptErrorsSuppressed = true;
            if (check_roll_dw == 1)
            {
                checkBox1.CheckState=CheckState.Checked;
            }
            if (check_roll_dw == 0)
            {
                checkBox1.CheckState = CheckState.Unchecked;
            }


        }

        private void button43_Click(object sender, EventArgs e)
        {

        }

        private void button44_Click_1(object sender, EventArgs e)
        {
            //  string zipPath = @"C:\Program Files\7-Zip\7z.exe";
            //  string startPath = @"h:\4334.txt";
            //  string extractPath = @"h:\vm\4334.txt";
            // ZipFile.CreateFromDirectory(startPath, zipPath);

            //ZipFile.ExtractToDirectory(zipPath, extractPath);
            //  GZipStream.

            // FolderBrowserDialog bd = new FolderBrowserDialog();

        }

        private void button50_Click(object sender, EventArgs e)
        {

        }

        private void button51_Click(object sender, EventArgs e)
        {


        }

        private void button52_Click(object sender, EventArgs e)
        {
        }

        private void button53_Click(object sender, EventArgs e)
        {

        }

        private void button55_Click(object sender, EventArgs e)
        {

        }

        private void button35_Click_1(object sender, EventArgs e)
        {


        }

        private void button36_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(@"dcxenapp.guo.local");
        }

        private void button37_Click_1(object sender, EventArgs e)
        {
            //  Process.Start("printmanagement.msc");
            // printmanagement.msc
        }

        private void button37_Click_2(object sender, EventArgs e)
        {

        }

        private void button39_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(@"printmanagement.msc");
        }

        private void button56_Click(object sender, EventArgs e)
        {

        }

        private void button57_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://files.regenergy.ru/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button58_Click(object sender, EventArgs e)
        {

        }

        private void button49_Click_1(object sender, EventArgs e)
        {

        }

        private void button59_Click(object sender, EventArgs e)
        {

        }

        private void button60_Click(object sender, EventArgs e)
        {

        }

        private void button49_Click_2(object sender, EventArgs e)
        {

            try
            {
                //textBox6.Text;

                String dir_work = textBox6.Text;
                String dir_arch = @"\\10.10.0.27\HOME_disable_user_SIBSUP\" + listBox1.SelectedItem.ToString() + ".zip";
                if (!System.IO.File.Exists(dir_arch))
                {
                    MessageBox.Show(@"archive Exist " + dir_arch, "del profile", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (System.IO.File.Exists(dir_arch))
                {

                    // Directory.Delete(dir_work, true);
                    string message = "archive is present , profile remove ?";
                    string caption = "profile remove";
                    MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                    DialogResult result;

                    // Displays the MessageBox.

                    result = MessageBox.Show(message, caption, buttons);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        // MessageBox.Show(@"archive is present , profile removed", "del profile", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        // Directory.Delete(@"\\sibsup\e$\"+ listBox1.SelectedItem.ToString(), true);
                        //Directory.Delete(@"\\sibstore\e$\" + listBox1.SelectedItem.ToString(), true);
                        Directory.Delete(dir_work, true);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "del profile", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button61_Click(object sender, EventArgs e)
        {

        }

        private void button62_Click(object sender, EventArgs e)
        {

        }

        private void button63_Click(object sender, EventArgs e)
        {

        }

        private void button64_Click(object sender, EventArgs e)
        {


        }

        private void button66_Click(object sender, EventArgs e)
        {

        }

        private void button65_Click(object sender, EventArgs e)
        {

        }

        private void button67_Click(object sender, EventArgs e)
        {

        }

        private void button68_Click(object sender, EventArgs e)
        {

        }

        private void button57_Click_1(object sender, EventArgs e)
        {

        }

        private void button69_Click(object sender, EventArgs e)
        {

        }

        private void button70_Click(object sender, EventArgs e)
        {

        }

        private void button71_Click(object sender, EventArgs e)
        {

        }

        private void button72_Click(object sender, EventArgs e)
        {

        }

        private void button73_Click(object sender, EventArgs e)
        {

        }

        private void button75_Click(object sender, EventArgs e)
        {

        }

        private void button74_Click(object sender, EventArgs e)
        {

        }

        private void button72_Click_1(object sender, EventArgs e)
        {

        }

        private void button74_Click_1(object sender, EventArgs e)
        {

        }

        private void button75_Click_1(object sender, EventArgs e)
        {

        }

        private void button76_Click(object sender, EventArgs e)
        {
            Process.Start(@"\\sibsup\DISTR\okhten\Zuma Deluxe\Zuma.exe");
        }

        private void button77_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button78_Click(object sender, EventArgs e)
        {

        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @" \\sibstore\");
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @" \\sibsup\");
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @"\\guo.local\DFSFILES\NSK\OBMEN");
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @"\\guo.local\DFSFILES\IT\10.0_SIB");
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @"\\guo.local\dfssvc\_loginlogNSK");
        }

        private void metroButton7_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @"\\dcfilesrv");
        }

        private void metroButton6_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("explorer.exe ", @"\\sibsup\distr");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "explorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroButton8_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @" C:\");
        }

        private void metroButton9_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @" H:\");
        }

        private void metroButton10_Click(object sender, EventArgs e)
        {
            Form1.ActiveForm.Close();
        }

        private void metroButton12_Click(object sender, EventArgs e)
        {

        }

        private void metroButton11_Click(object sender, EventArgs e)
        {

        }

        private void metroButton13_Click(object sender, EventArgs e)
        {

        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            Process.Start("cmd.exe");
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            Process.Start("printmanagement.msc");
        }

        private void metroTile3_Click(object sender, EventArgs e)
        {
            Process.Start(file_msc);
        }

        private void metroTile6_Click(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:sibstore");
        }

        private void metroTile5_Click(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:sibsup");
        }

        private void metroTile4_Click(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:sibdc");
        }

        private void metroTile9_Click(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:10.10.0.164");
        }

        private void metroTile10_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "powershell", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile11_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://rm.guo.local/redmine/projects/nskadm/wiki/Info_kass_istall");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile12_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://rm.guo.local/redmine/projects/task-for-admin/issues");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile13_Click(object sender, EventArgs e)
        {
          
            string s;
            s = textBox1.Text;
            s = " -c: -h -m:" + s + "  -a:1 -x:";
            // textBox10.Text = dmwrPath + "DWRCC.exe"+ s;
            //Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:" + s + "  -a:1 -x:");
            Process.Start(dmwrPath, s);



            if (check_roll_dw == 1)
            {
                this.WindowState = FormWindowState.Minimized;
            }
        }

        private void metroTile14_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("explorer.exe ", @" \\" + textBox1.Text + @"\c$");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error pach ", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void metroTile15_Click(object sender, EventArgs e)
        {
            textBox1.Text = "10.10.";
        }

        private void metroTile16_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("cmd.exe", "/K ping -t " + textBox1.Text);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile17_Click(object sender, EventArgs e)
        {
            Process.Start("excel.exe", @"""\\dcfilesrv\DOC\телефонная книга.xlsx""");
        }

        private void metroTile18_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile19_Click(object sender, EventArgs e)
        {
            Process.Start("C:/Users/okhten/Desktop/putty.exe");
        }

        private void metroTile20_Click(object sender, EventArgs e)
        {
            Process.Start(@"C:\Program Files (x86)\VMware\Infrastructure\Virtual Infrastructure Client\Launcher\VpxClient.exe");
        }

        private void metroTile21_Click(object sender, EventArgs e)
        {
            Process.Start("C:/totalcmd/TOTALCMD64.EXE");
        }

        private void metroTile22_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Notepad++\notepad++.exe");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile23_Click(object sender, EventArgs e)
        {
            Process.Start(acronis_xml);
        }

        private void metroTile24_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "http://dcnag.guo.local/nagios/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile25_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "https://mail.regenergy.ru");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroButton11_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start("explorer.exe ", @"\\sibsup\distr\Фискальники");
            }
            catch (Exception)
            {
                MessageBox.Show(@"не могу открыть", "explorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroButton12_Click_1(object sender, EventArgs e)
        {
            
        }

        private void metroTile27_Click(object sender, EventArgs e)
        {
            
            try
            {
                String dir = @"\\" + label2.Text + @"\c$\Program Files\UK";
                if (Directory.Exists(dir))
                {
                   Directory.Delete(dir, true);
                   
                }


                if (!Directory.Exists(dir))
                {

                    Directory.CreateDirectory(dir);
                    FileInfo f = new FileInfo(@"\\sibsup\UK\pu.exe");
                    f.CopyTo(dir + @"\pu.exe");
                   
                    FileInfo d = new FileInfo(@"\\sibsup\UK\del.exe");
                    d.CopyTo(dir + @"\del.exe");
                    
                    FileInfo p = new FileInfo(@"\\sibsup\UK\path.txt");
                    p.CopyTo(dir + @"\path.txt");
                   
                    FileInfo l = new FileInfo(@"\\sibsup\UK\log.txt");
                    l.CopyTo(dir + @"\log.txt");
                    
                    FileInfo n = new FileInfo(@"\\sibsup\UK\nircmd.exe");
                    n.CopyTo(dir + @"\nircmd.exe");

                    FileInfo user_profile = new FileInfo(@"\\sibsup\UK\pu.lnk");
                    user_profile.CopyTo(textBox6.Text + @"\desktop\pu.lnk");

                    //FileInfo lnk = new FileInfo(@"\\sibsup\UK\pu.lnk");
                    //lnk.CopyTo(@"\\sibsup\ + @"\pu.lnk");

                    DirectoryInfo directory = new DirectoryInfo(dir);
                   
                    FileSystemAccessRule rule = new FileSystemAccessRule(@"sib\PU", FileSystemRights.FullControl,  AccessControlType.Allow );
                   
                    DirectorySecurity security = new DirectorySecurity();
                 
                    security.AddAccessRule(rule);
                    directory.SetAccessControl(security);
                    //InheritanceFlags.ContainerInherit, PropagationFlags.None,

                    DirectoryInfo dir_modify = new DirectoryInfo(dir);
                    FileSystemAccessRule rule_modify = new FileSystemAccessRule(@"sib\PU", FileSystemRights.FullControl, InheritanceFlags.ObjectInherit, PropagationFlags.InheritOnly, AccessControlType.Allow);
                    DirectorySecurity sec_mod = new DirectorySecurity();
                    sec_mod.AddAccessRule(rule_modify);
                    directory.SetAccessControl(sec_mod);

                    MessageBox.Show(@" copy - ok ", "copy", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception)
            {

                MessageBox.Show(@"error copy " , "copy", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void metroTile26_Click(object sender, EventArgs e)
        {
            try
            {
                String dir = @"\\" + label2.Text + @"\c$\Program Files\install_comf";

                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);

                    FileInfo f = new FileInfo(@"\\sibsup\distr\Комфорт\ComfortClient.EXE");
                    f.CopyTo(dir + @"\ComfortClient.EXE");
                    FileInfo fn = new FileInfo(@"\\sibsup\distr\Комфорт\HCS_Client.exe");
                    fn.CopyTo(dir + @"\HCS_Client.exe");
                    FileInfo fnn = new FileInfo(@"\\sibsup\distr\Комфорт\servers.txt");
                    fnn.CopyTo(dir + @"\servers.txt");
                    FileInfo fnnx = new FileInfo(@"\\sibsup\distr\Комфорт\Инструкция по установке.doc");
                    fnnx.CopyTo(dir + @"\Инструкция по установке.doc");
                    MessageBox.Show(@" copy - ok ", "copy", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }


                //  FileInfo n = new FileInfo(@"\\sibsup\UK\nircmd.exe");
                // n.CopyTo(dir + @"\nircmd.exe");


            }
            catch (Exception)
            {

                MessageBox.Show(@"error copy ", "copy", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void metroTile28_Click(object sender, EventArgs e)
        {
            ps_user_info();
        }

        private void metroTile29_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("cmd.exe", @" psexec \\" + label2.Text + @" gpupdate /force");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile30_Click(object sender, EventArgs e)
        {

            Process.Start(msinfo_xml, @" /computer " + label2.Text);
        }

        private void metroTile31_Click(object sender, EventArgs e)
        {
            try
            {
                get_user_info();
            }
            catch (Exception)
            {
                MessageBox.Show(@"не выбран юзер", "тел", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile32_Click(object sender, EventArgs e)
        {
            label2.Text = "";

            listBox2.Items.Clear();
            int i;
            string fio;

            if (listBox1.SelectedItems.Count == 1)
            {

                try
                {
                    i = listBox1.SelectedIndex;
                    fio = listBox1.SelectedItem.ToString();
                    string path = @"\\guo.local\dfssvc\_loginlogNSK\userlog" + fio + ".txt";

                    string[] arStr = File.ReadAllLines(path, Encoding.GetEncoding(1251));
                    listBox2.Items.AddRange(arStr);
                    listBox2.SelectedIndex = listBox2.Items.Count - 1;


                }
                catch (Exception)
                {
                    MessageBox.Show(@"error, wrong path \\guo.local\dfssvc\_loginlogNSK\userlog", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            else
            {
                MessageBox.Show(@"выбрано больше одного", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void metroTile33_Click(object sender, EventArgs e)
        {
            try
            {
                string in_p, out_p, user_n, cmd_com;
                user_n = listBox1.SelectedItem.ToString() + ".7z";
                in_p = textBox6.Text;
                out_p = @"\\10.10.0.27\f$\HOME_disable_user_SIBSUP\" + listBox1.SelectedItem.ToString() + ".7z";

                // out_p = @"H:\" + listBox1.SelectedItem.ToString() + ".zip";
                //     ZipFile zf = new ZipFile(out_p);
                //      zf.AddDirectory(in_p);
                //        zf.Save();
                //      MessageBox.Show("Архивация прошла успешно.", "Выполнено");
                //  Process.Start(@"C:\Program Files\7-Zip\7z.exe" , " a -tzip -ssw -mx9 -r0 " + out_p+"   "+ in_p);
                //  21.03.2019      cmd_com = (@"""C:\Program Files\7-Zip\7z.exe""  a -tzip -ssw -mx9 -r0 " + out_p + "   " + in_p + " && pause");
                cmd_com = (@"""C:\Program Files\7-Zip\7z.exe""  a -t7z -ssw -mx5 -r0 " + out_p + "   " + in_p + " && pause");
                Process.Start("cmd.exe", "/K " + cmd_com);
                //  MessageBox.Show(cmd_com, "profile save zip", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error ", "profile save zip", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile34_Click(object sender, EventArgs e)
        {
            string s;

            try
            {
                s = textBox1.Text;
                Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe ", @"Get-WMIObject -Class Win32_ComputerSystem -Computer " + s + "|Select-Object Username"+ @"""Press Enter to exit""");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "comp sel err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile35_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("cmd.exe", "/K ping -t " + label2.Text);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile36_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(@"sib\" + listBox1.SelectedItem.ToString());
        }

        private void metroTile37_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(listBox1.SelectedItem.ToString());
        }

        private void metroTile38_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(label2.Text);
        }

        private void metroTile41_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
            //listBox2.ClearSelected();
            listBox2.Items.Clear();
            label2.Text = "";
            textBox2.Focus();
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            label55.Text = "";
        }

        private void metroTile39_Click(object sender, EventArgs e)
        {
            if (check_roll_dw == 1)
            {
                this.WindowState = FormWindowState.Minimized;
            }

            string ncomp, s, ss;
            char ck;
            int i_s;
            ncomp = "";
            try
            {
                ncomp = listBox2.SelectedItem.ToString();
                char c = ncomp[0];
                if (c == 'В')
                {
                    s = ncomp.Remove(0, 5);
                }
                else
                {
                    s = ncomp.Remove(0, 10);
                }

                for (int i = 0; i < s.Length; i++)
                {
                    ck = s[i];
                    if (ck == ',')
                    {
                        i_s = i;
                        ss = s.Remove(i, s.Length - i);
                        label2.Text = ss;
                        //ps_user_info();
                        Process.Start(dmwrPath, "-c: -h -m:" + ss + " -a:1 -x:");
                        break;
                    }




                    //Form1.ActiveForm.WindowState = FormWindowState.Minimized;

                }

            }
            catch (Exception)
            {
                MessageBox.Show(@"error selected ", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

        }

        private void metroTile40_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("explorer.exe ", @" \\" + label2.Text + @"\c$");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error pach ", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void metroTile42_Click(object sender, EventArgs e)
        {
            try
            {

                String user = listBox1.SelectedItem.ToString();
                String dir = @"\\guo.local\dfsfiles\nsk\scan\" + user;
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                    DirectoryInfo directory = new DirectoryInfo(dir);
                    FileSystemAccessRule rule_Allow = new FileSystemAccessRule(@"sib\" + user, FileSystemRights.Read | FileSystemRights.ListDirectory | FileSystemRights.WriteData | FileSystemRights.ExecuteFile | FileSystemRights.Delete, InheritanceFlags.ObjectInherit,
            PropagationFlags.None, AccessControlType.Allow);
                    FileSystemAccessRule rule_Deny = new FileSystemAccessRule(@"sib\" + user, FileSystemRights.Delete, AccessControlType.Deny);

                    DirectorySecurity security = new DirectorySecurity();
                    security.AddAccessRule(rule_Allow);
                    security.AddAccessRule(rule_Deny);
                    directory.SetAccessControl(security);


                    ps_run("addscan");
                    MessageBox.Show(@"scan add - ok добавь пользователя в группу scan....", "add scan rule", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(@"уже есть папка", "add scan rule", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception)
            {
                MessageBox.Show(@"error add scan rule", "add scan", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile43_Click(object sender, EventArgs e)
        {
            String user = listBox1.SelectedItem.ToString();
            String dir = @"\\guo.local\dfsfiles\nsk\SCAN\" + user;

            if (Directory.Exists(dir))
            {
                Clipboard.SetText(dir);
                MessageBox.Show(dir, "scan folder", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            {
                MessageBox.Show(@"нет папки", "scan folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile44_Click(object sender, EventArgs e)
        {
            
            
            Clipboard.SetText(@"sib\scansib");
            MessageBox.Show("", @"sib\scansib", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile45_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("E;fcysqgfhjkm");
            MessageBox.Show("", @"pass in buff", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile46_Click(object sender, EventArgs e)
        {
            String user = listBox1.SelectedItem.ToString();
            String dir = @"\\sibstore\SCAN\" + user;

            Process.Start("explorer.exe ", dir);
        }

        private void metroTile48_Click(object sender, EventArgs e)
        {

            if (textBox6.Text != "")
            {
                Process.Start("explorer.exe ", textBox6.Text);
            }
        }

        private void metroTile50_Click(object sender, EventArgs e)
        {

            if (check_roll_dw == 1)
            {
                this.WindowState = FormWindowState.Minimized;
            }
            string s;
            // s = label9.Text;
            s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);
            if (s != "")
            {
                Process.Start(dmwrPath, "-c: -h -m:" + s + " -a:1 -x:");

            }
        }

        private void metroTile49_Click(object sender, EventArgs e)
        {
            string s;

            //  s = label9.Text;
            s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);
            //table1DataGridView

            if (s != "")
            {
                Process.Start("explorer.exe ", @" \\" + s + @"\c$");
            }
        }

        private void metroTile51_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "https://ess-p.ru/login.htm");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile52_Click(object sender, EventArgs e)
        {
            String dir = @"https://ess-p.ru";
            Clipboard.SetText(dir);
        }

        private void metroTile53_Click(object sender, EventArgs e)
        {
            string s;

            //  s = label9.Text;
            s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);
            try
            {
                Process.Start("cmd.exe", "/K ping -t " + s);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile54_Click(object sender, EventArgs e)
        {
            string s;

            //  s = label9.Text;
            s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);
            try
            {
                Process.Start("cmd.exe", @"/K tracert " + s);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile55_Click(object sender, EventArgs e)
        {
            string s;

            //  s = label9.Text;
            s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);
            try
            {
                Process.Start("cmd.exe", @"/K psexec \\" + s + @" cmd.exe");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile56_Click(object sender, EventArgs e)
        {
            string s;

            //  s = label9.Text;
            s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);
            try
            {
                Process.Start("cmd.exe", @"/K psexec \\" + s + @"  gpupdate /force");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile57_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "https://10.10.0.7/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button73_Click_1(object sender, EventArgs e)
        {

        }

        private void metroTile59_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "https://10.10.1.20");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile60_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "https://10.10.1.21/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile47_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"\\sibsup\DISTR\AA_v3.exe");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail / no file", "aa", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile61_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "https://mail.yandex.ru/?uid=1130000013527135&login=anton#inbox");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile62_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "https://translit.net/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void metroTile63_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "https://novosibirsk.e2e4online.ru/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile64_Click(object sender, EventArgs e)
        {
            Process.Start("excel.exe", @"""\\sibsup\distr\okhten\Сертификаты\МКС\Рабочие_места_пользователи_МКС.xlsx""");
        }

        private void metroTile65_Click(object sender, EventArgs e)
        {
            Process.Start("excel.exe", @"""\\sibsup\distr\okhten\Сертификаты\Дзержинец\Рабочие_места_пользователи_Дзержинец.xlsx""");
        }

        private void metroTile66_Click(object sender, EventArgs e)
        {

            Process.Start("explorer.exe ", @"\\sibsup\distr\kassa\kas_cert");
        }

        private void metroTile67_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "https://10.10.1.7/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile68_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://10.222.0.3");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile70_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://cdr.guo.local/index.php/index/index");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile69_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://files.regenergy.ru/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile71_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe ", @"Get-ADUser  -filter {Enabled -eq $True} -properties Description,LastLogonDate | where {$_.lastlogondate -lt (get-date).addmonths(-2)} | sort LastLogonDate | FT Name, LastLogonDate,Description;" + @" Read-Host -Prompt" + @"""Press Enter to exit""");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile72_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "https://www.dmosk.ru/miniinstruktions.php?mini=powershell-adgroup#get");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile73_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "https://vk.com/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile74_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "https://pikabu.ru/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile75_Click(object sender, EventArgs e)
        {
           
        }

        private void metroTile76_Click(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "https://www.avito.ru/novosibirsk");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile77_Click(object sender, EventArgs e)
        {

            try
            {
                //String dir = @"\\" + label2.Text + @"\c$\Program Files\UK";
                //FileInfo d = new FileInfo(dir + @"\path.txt");
                //FileInfo p = new FileInfo(@"\\sibsup\UK\path.txt");
                //d.Delete();
                //p.CopyTo(dir + @"\path.txt");
                //MessageBox.Show(@" copy - path.txt ok ", "copy", MessageBoxButtons.OK, MessageBoxIcon.Information);
       
                    ps_run("addpu");
                    MessageBox.Show(@"PU add - ok добавь пользователя в группу PU....", "add PU rule", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            catch (Exception)
            {
                MessageBox.Show(@" pu add err ", "add right err", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


         }

        private void metroTile78_Click(object sender, EventArgs e)
        {
           
        }

        private void metroTile79_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("services.msc");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "services", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile80_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\LanHelper\LanHelper.exe");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "prog exist", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile81_Click(object sender, EventArgs e)
        {
            ps_logon_user();
        }

        private void cs_MakeCall_Click(object sender, EventArgs e)
        {

        }

        private void cs_Release_Click(object sender, EventArgs e)
        {
            
        }

        private void cs_Answer_Click(object sender, EventArgs e)
        {

        }

        private void metroTile82_Click(object sender, EventArgs e)
        {
           
        }

        private void metroTile83_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            metroTile83.Style = MetroFramework.MetroColorStyle.Green;
            metroTile85.Style = MetroFramework.MetroColorStyle.Green;
            if ((call_status == 3) | (call_status == 0))
            {
                cs_Phone.Clear();

                CallManager.onUserRelease(v_hCall.Session);
            }

            else if ((call_status == 2))
            {
                CallManager.onUserRelease(v_hIncomingCall.Session);
                cs_Phone.Clear();
            }
            else if (call_status == 1)
            {
                CallManager.onUserAnswer(v_hIncomingCall.Session);
                CallManager.onUserRelease(v_hIncomingCall.Session);
                cs_Phone.Clear();
            }
            metroTile85.Text = "О Т В Е Т И Т Ь";
            metroTile84.Text = "В Ы З О В";
        }

        private void metroTile84_Click(object sender, EventArgs e)
        {
            if (cs_Phone.Text != "")
            {
                v_hCall = CallManager.createOutboundCall(cs_Phone.Text);

                //  metroTile84.Enabled = false;
                // System.Threading.Thread.Sleep(500);
                //  metroTile84.Enabled = true;
            }
            call_status = 3;
        }

        private void metroTile85_Click(object sender, EventArgs e)
        {
            CallManager.onUserAnswer(v_hIncomingCall.Session);
            textBox3.Text = v_hIncomingCall.CallingName;
            player.Stop();
            // timer1.Enabled = false;
            timer1.Stop();
            metroTile85.Style = MetroFramework.MetroColorStyle.Green;
            call_status = 2;
            metroTile83.Style = MetroFramework.MetroColorStyle.Red;
        }

        private void cs_CallState_TextChanged(object sender, EventArgs e)
        {
            
            if (cs_CallState.Text == "INCOMING")
            {
                this.WindowState = FormWindowState.Maximized;
                this.TopMost = true;
                // System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"c:\mywavfile.wav");
                player.Play();
                this.TopMost = false;
                // metroTile85.Style =MetroFramework.MetroColorStyle.Red;
                call_status = 1;

                timer1.Interval = 500;
                timer1.Enabled = true;

                timer1.Start();
                
                //  i = System.Convert.ToInt32(cs_Callnomber.Text);
                //metroTile85.TileCount = i;


            }
            else if  (cs_CallState.Text == "NULL")
            {
                timer1.Stop();
                metroTile85.Style = MetroFramework.MetroColorStyle.Green;
                metroTile84.Text = "В Ы З О В";
                callstate.Text = "callstate:NULL";
            }
            else if (cs_CallState.Text == "ACTIVE")
            {
                timer1.Stop();
                metroTile85.Style = MetroFramework.MetroColorStyle.Green;
                callstate.Text = "callstate:ACTIVE";
            }




        }

        private void metroTile86_Click(object sender, EventArgs e)
        {
           
        }

        private void metroTile87_Click(object sender, EventArgs e)
        {
            


        }



        private void metroTile84_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile84.Style = MetroFramework.MetroColorStyle.Teal;
            metroTile84.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }



        private void metroTile84_MouseLeave(object sender, EventArgs e)
        {
            metroTile84.Style = MetroFramework.MetroColorStyle.Green;
            metroTile84.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }



        private void metroTile84_MouseDown(object sender, MouseEventArgs e)
        {
            metroTile84.Style = MetroFramework.MetroColorStyle.Red;
            metroTile84.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile84_MouseMove_1(object sender, MouseEventArgs e)
        {
            metroTile84.Style = MetroFramework.MetroColorStyle.Teal;
            metroTile84.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;

        }

        private void metroTile83_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile83.Style = MetroFramework.MetroColorStyle.Red;
            metroTile83.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile83_MouseLeave(object sender, EventArgs e)
        {
            metroTile83.Style = MetroFramework.MetroColorStyle.Green;
            metroTile83.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile85_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile85.Style = MetroFramework.MetroColorStyle.Teal;
            metroTile85.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            metroTile85.Text = "О Т В Е Т И Т Ь";
        }

        private void metroTile85_MouseLeave(object sender, EventArgs e)
        {
            metroTile85.Style = MetroFramework.MetroColorStyle.Green;
            metroTile85.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        

        private void metroTabPage1_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
           

           
            
                
                {
                    if (metroTile85.Style == MetroFramework.MetroColorStyle.Green)
                    {
                       
                        metroTile85.Style = MetroFramework.MetroColorStyle.Red;
                    }
                    else
                    {
                        
                        metroTile85.Style = MetroFramework.MetroColorStyle.Green;
                    }
                }
            
           
        }

        private void Form1_Leave(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            servise_int = 1;
        }

        private void metroTile62_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile62.Style = MetroFramework.MetroColorStyle.Green;
            metroTile62.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile62_MouseLeave(object sender, EventArgs e)
        {
            metroTile62.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile62.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void cs_Phone_DoubleClick(object sender, EventArgs e)
        {
            cs_Phone.Text = "";
        }

        private void textBox4_DoubleClick(object sender, EventArgs e)
        {
            if (textBox4.Text != "")
            {
                if (cs_CallState.Text != "ACTIVE")
                {
                    cs_Phone.Text = textBox4.Text;
                    v_hCall = CallManager.createOutboundCall(cs_Phone.Text);
                }

            }
             metroTile84.Text = "В Ы З О В  "+textBox4.Text;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
           {
                check_roll_dw = 1;
           }
            else
            {
                check_roll_dw = 0;
            }


        }

        private void metroTile82_Click_1(object sender, EventArgs e)
        {
            string s;
            s = Convert.ToString(table1DataGridView[3, table1DataGridView.CurrentRow.Index].Value);
            //metroTile87.Text = s;
            cs_Phone.Text = s;

            if (cs_CallState.Text != "ACTIVE")
            {

                v_hCall = CallManager.createOutboundCall(cs_Phone.Text);
            }
            metroTabControl1.SelectedIndex = 0;
        }

        private void metroTile82_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile82.Style = MetroFramework.MetroColorStyle.Teal;
            metroTile82.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile82_MouseLeave(object sender, EventArgs e)
        {
            metroTile82.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile82.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroButton13_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(@"setvar=RECORD=1
setvar=REC_PROJECT=zao_mks");
        }

        private void metroTile86_Click_1(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox1.Text = Clipboard.GetText();

            try
            {
                Process.Start("cmd.exe", "/K ping -t " + textBox1.Text);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile87_Click_1(object sender, EventArgs e)
        {

            try
            {
              
                Process.Start("compmgmt.msc", " /computer=" + textBox1.Text);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "compmgmt.msc", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile88_Click(object sender, EventArgs e)
        {
            try
            {
               // Process.Start("сompmgmt.msc", " /computer=" + label2.Text);
                Process.Start("compmgmt.msc", " /computer=" + label2.Text);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error" + label2.Text, "compmgmt.msc " , MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile8_Click(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:sib-remapp");
        }

        private void metroTile7_Click(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:dcregterm");
        }

        private void metroTile90_Click(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe");
        }

        private void metroTile89_Click(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:sib-term");
        }

        private void metroTile3_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile3.Style = MetroFramework.MetroColorStyle.Green;
            metroTile3.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile3_MouseLeave(object sender, EventArgs e)
        {
            metroTile3.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile3.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile2_MouseLeave(object sender, EventArgs e)
        {

            metroTile2.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile2.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile1_MouseLeave(object sender, EventArgs e)
        {

            metroTile1.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile1.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile6_MouseLeave(object sender, EventArgs e)
        {

            metroTile6.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile6.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile5_MouseLeave(object sender, EventArgs e)
        {

            metroTile5.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile5.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile4_MouseLeave(object sender, EventArgs e)
        {

            metroTile4.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile4.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile9_MouseLeave(object sender, EventArgs e)
        {

            metroTile9.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile9.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile8_MouseLeave(object sender, EventArgs e)
        {

            metroTile8.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile8.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile7_MouseLeave(object sender, EventArgs e)
        {

            metroTile7.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile7.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile21_MouseLeave(object sender, EventArgs e)
        {

            metroTile21.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile21.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile90_MouseLeave(object sender, EventArgs e)
        {

            metroTile90.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile90.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile89_MouseLeave(object sender, EventArgs e)
        {

            metroTile89.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile89.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile22_MouseLeave(object sender, EventArgs e)
        {

            metroTile22.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile22.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile20_MouseLeave(object sender, EventArgs e)
        {

            metroTile20.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile20.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile17_MouseLeave(object sender, EventArgs e)
        {

            metroTile17.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile17.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile23_MouseLeave(object sender, EventArgs e)
        {

            metroTile23.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile23.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile10_MouseLeave(object sender, EventArgs e)
        {

            metroTile10.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile10.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile80_MouseLeave(object sender, EventArgs e)
        {

            metroTile80.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile80.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile19_MouseLeave(object sender, EventArgs e)
        {

            metroTile19.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile19.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile79_MouseLeave(object sender, EventArgs e)
        {

            metroTile79.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile79.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile2_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile2.Style = MetroFramework.MetroColorStyle.Green;
            metroTile2.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile1_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile1.Style = MetroFramework.MetroColorStyle.Green;
            metroTile1.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile6_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile6.Style = MetroFramework.MetroColorStyle.Green;
            metroTile6.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile5_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile5.Style = MetroFramework.MetroColorStyle.Green;
            metroTile5.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile4_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile4.Style = MetroFramework.MetroColorStyle.Green;
            metroTile4.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile9_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile9.Style = MetroFramework.MetroColorStyle.Green;
            metroTile9.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile8_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile8.Style = MetroFramework.MetroColorStyle.Green;
            metroTile8.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile7_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile7.Style = MetroFramework.MetroColorStyle.Green;
            metroTile7.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile21_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile21.Style = MetroFramework.MetroColorStyle.Green;
            metroTile21.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile90_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile90.Style = MetroFramework.MetroColorStyle.Green;
            metroTile90.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile89_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile89.Style = MetroFramework.MetroColorStyle.Green;
            metroTile89.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile22_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile22.Style = MetroFramework.MetroColorStyle.Green;
            metroTile22.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile20_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile20.Style = MetroFramework.MetroColorStyle.Green;
            metroTile20.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile17_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile17.Style = MetroFramework.MetroColorStyle.Green;
            metroTile17.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile23_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile23.Style = MetroFramework.MetroColorStyle.Green;
            metroTile23.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile10_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile10.Style = MetroFramework.MetroColorStyle.Green;
            metroTile10.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile80_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile80.Style = MetroFramework.MetroColorStyle.Green;
            metroTile80.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile19_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile19.Style = MetroFramework.MetroColorStyle.Green;
            metroTile19.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile79_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile79.Style = MetroFramework.MetroColorStyle.Green;
            metroTile79.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile39_MouseLeave(object sender, EventArgs e)
        {
            metroTile39.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile39.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile40_MouseLeave(object sender, EventArgs e)
        {
            metroTile40.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile40.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile48_MouseLeave(object sender, EventArgs e)
        {
            metroTile48.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile48.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile35_MouseLeave(object sender, EventArgs e)
        {
            metroTile35.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile35.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile88_MouseLeave(object sender, EventArgs e)
        {
            metroTile88.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile88.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile41_MouseLeave(object sender, EventArgs e)
        {
            metroTile41.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile41.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile37_MouseLeave(object sender, EventArgs e)
        {
            metroTile37.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile37.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile36_MouseLeave(object sender, EventArgs e)
        {
            metroTile36.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile36.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile38_MouseLeave(object sender, EventArgs e)
        {
            metroTile38.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile38.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile39_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile39.Style = MetroFramework.MetroColorStyle.Green;
            metroTile39.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile40_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile40.Style = MetroFramework.MetroColorStyle.Green;
            metroTile40.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile48_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile48.Style = MetroFramework.MetroColorStyle.Green;
            metroTile48.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile35_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile35.Style = MetroFramework.MetroColorStyle.Green;
            metroTile35.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile88_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile88.Style = MetroFramework.MetroColorStyle.Green;
            metroTile88.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile41_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile41.Style = MetroFramework.MetroColorStyle.Green;
            metroTile41.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile37_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile37.Style = MetroFramework.MetroColorStyle.Green;
            metroTile37.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile36_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile36.Style = MetroFramework.MetroColorStyle.Green;
            metroTile36.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile38_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile38.Style = MetroFramework.MetroColorStyle.Green;
            metroTile38.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile91_Click(object sender, EventArgs e)
        {
            try
            {
                // Process.Start("сompmgmt.msc", " /computer=" + label2.Text);
                Process.Start("services.msc", " /computer=" + label2.Text);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error" + label2.Text, "compmgmt.msc ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile91_MouseMove(object sender, MouseEventArgs e)
        {

            metroTile91.Style = MetroFramework.MetroColorStyle.Green;
            metroTile91.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile91_MouseUp(object sender, MouseEventArgs e)
        {
            
        }

        private void metroTile91_MouseLeave(object sender, EventArgs e)
        {
            metroTile91.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile91.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile93_Click(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall("*8");
        }

        private void metroTile93_MouseLeave(object sender, EventArgs e)
        {
            metroTile93.Style = MetroFramework.MetroColorStyle.Green;
            metroTile93.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
          
        }

        private void metroTile93_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile93.Style = MetroFramework.MetroColorStyle.Teal;
            metroTile93.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void label6_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(label6.Text);
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(label12.Text);
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.11");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.207");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        

        private void panel4_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.10");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {
           
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.11");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel3_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", label10.Text);
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel5_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.12");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel6_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.13");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel7_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.14");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel9_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.15");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel8_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.18");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel11_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.8");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel12_Click(object sender, EventArgs e)
        {
            Process.Start(@"C:\Program Files (x86)\VMware\Infrastructure\Virtual Infrastructure Client\Launcher\VpxClient.exe");
        }

        private void metroTile57_Click_1(object sender, EventArgs e)
        {

            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "http://10.10.0.24");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel11_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label48_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.2.10");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void metroTile61_MouseLeave(object sender, EventArgs e)
        {

            metroTile61.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile61.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile61_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile61.Style = MetroFramework.MetroColorStyle.Green;
            metroTile61.TileTextFontSize = MetroFramework.MetroTileTextSize.Medium;
        }

        private void metroTile92_MouseLeave(object sender, EventArgs e)
        {
            metroTile92.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile92.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile87_MouseLeave(object sender, EventArgs e)
        {
            metroTile87.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile87.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile86_MouseLeave(object sender, EventArgs e)
        {
            metroTile86.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile86.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile16_MouseLeave(object sender, EventArgs e)
        {
            metroTile16.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile16.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile14_MouseLeave(object sender, EventArgs e)
        {
            metroTile14.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile14.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile15_MouseLeave(object sender, EventArgs e)
        {
            metroTile15.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile15.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile13_MouseLeave(object sender, EventArgs e)
        {
            metroTile13.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile13.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile42_MouseLeave(object sender, EventArgs e)
        {
            metroTile42.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile42.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile46_MouseLeave(object sender, EventArgs e)
        {
            metroTile46.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile46.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile43_MouseLeave(object sender, EventArgs e)
        {
            metroTile43.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile43.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile44_MouseLeave(object sender, EventArgs e)
        {
            metroTile44.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile44.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile45_MouseLeave(object sender, EventArgs e)
        {
            metroTile45.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile45.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile32_MouseLeave(object sender, EventArgs e)
        {
            metroTile32.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile32.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile31_MouseLeave(object sender, EventArgs e)
        {
            metroTile31.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile31.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile30_MouseLeave(object sender, EventArgs e)
        {
            metroTile30.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile30.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile29_MouseLeave(object sender, EventArgs e)
        {
            metroTile29.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile29.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile92_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile92.Style = MetroFramework.MetroColorStyle.Green;
            metroTile92.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile87_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile87.Style = MetroFramework.MetroColorStyle.Green;
            metroTile87.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile86_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile86.Style = MetroFramework.MetroColorStyle.Green;
            metroTile86.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile16_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile16.Style = MetroFramework.MetroColorStyle.Green;
            metroTile16.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile14_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile14.Style = MetroFramework.MetroColorStyle.Green;
            metroTile14.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile15_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile15.Style = MetroFramework.MetroColorStyle.Green;
            metroTile15.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile13_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile13.Style = MetroFramework.MetroColorStyle.Green;
            metroTile13.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile42_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile42.Style = MetroFramework.MetroColorStyle.Green;
            metroTile42.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile46_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile46.Style = MetroFramework.MetroColorStyle.Green;
            metroTile46.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile43_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile43.Style = MetroFramework.MetroColorStyle.Green;
            metroTile43.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile44_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile44.Style = MetroFramework.MetroColorStyle.Green;
            metroTile44.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile45_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile45.Style = MetroFramework.MetroColorStyle.Green;
            metroTile45.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile30_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile30.Style = MetroFramework.MetroColorStyle.Green;
            metroTile30.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile31_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile31.Style = MetroFramework.MetroColorStyle.Green;
            metroTile31.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile29_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile29.Style = MetroFramework.MetroColorStyle.Green;
            metroTile29.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile58_Click(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone1_tel);
            call_status = 3;
        }

        private void metroTile59_Click_1(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone2_tel);
            call_status = 3;
        }

        private void metroTile60_Click_1(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone3_tel);
            call_status = 3;
        }

        private void metroTile97_Click(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone4_tel);
            call_status = 3;
        }

        private void metroTile58_MouseLeave(object sender, EventArgs e)
        {
            metroTile58.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile58.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile59_MouseLeave(object sender, EventArgs e)
        {
            metroTile59.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile59.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile60_MouseLeave(object sender, EventArgs e)
        {
            metroTile60.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile60.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile97_MouseLeave(object sender, EventArgs e)
        {
            metroTile97.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile97.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile58_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile58.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile58.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile59_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile59.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile59.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile60_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile60.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile60.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile97_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile97.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile97.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroButton12_Click_2(object sender, EventArgs e)
        {
            try
            {
                Process.Start("explorer.exe ", @"\\guo.local\DFSFILES");
            }
            catch (Exception)
            {
                MessageBox.Show(@"не могу открыть", "explorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroButton15_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("explorer.exe ", @"C:\Users\okhten\Desktop\screen");
            }
            catch (Exception)
            {
                MessageBox.Show(@"не могу открыть", "explorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile94_Click(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone5_tel);
            call_status = 3;
        }

        private void metroTile95_Click(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone6_tel);
            call_status = 3;
        }

        private void metroTile96_Click(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone7_tel);
            call_status = 3;
        }

        private void metroTile100_Click(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone8_tel);
            call_status = 3;
        }

        private void metroTile102_Click(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone9_tel);
            call_status = 3;
        }

        private void metroTile99_Click(object sender, EventArgs e)
        {
            v_hCall = CallManager.createOutboundCall(Properties.Settings.Default.phone10_tel);
            call_status = 3;
        }

        private void metroTile18_MouseLeave(object sender, EventArgs e)
        {
            metroTile18.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile18.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile11_MouseLeave(object sender, EventArgs e)
        {
            metroTile11.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile11.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile12_MouseLeave(object sender, EventArgs e)
        {
            metroTile12.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile12.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile18_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile18.Style = MetroFramework.MetroColorStyle.Green;
            metroTile18.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile11_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile11.Style = MetroFramework.MetroColorStyle.Green;
            metroTile11.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile12_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile12.Style = MetroFramework.MetroColorStyle.Green;
            metroTile12.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile33_MouseLeave(object sender, EventArgs e)
        {
            metroTile33.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile33.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile27_MouseLeave(object sender, EventArgs e)
        {
            metroTile27.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile27.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile26_MouseLeave(object sender, EventArgs e)
        {
            metroTile26.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile26.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile28_MouseLeave(object sender, EventArgs e)
        {
            metroTile28.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile28.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile81_MouseLeave(object sender, EventArgs e)
        {
            metroTile81.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile81.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile34_MouseLeave(object sender, EventArgs e)
        {
            metroTile34.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile34.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile34_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile34.Style = MetroFramework.MetroColorStyle.Green;
            metroTile34.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile81_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile81.Style = MetroFramework.MetroColorStyle.Green;
            metroTile81.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile28_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile28.Style = MetroFramework.MetroColorStyle.Green;
            metroTile28.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile26_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile26.Style = MetroFramework.MetroColorStyle.Green;
            metroTile26.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile27_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile27.Style = MetroFramework.MetroColorStyle.Green;
            metroTile27.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile33_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile33.Style = MetroFramework.MetroColorStyle.Green;
            metroTile33.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        
        

       

        private void metroTile50_MouseLeave(object sender, EventArgs e)
        {
            metroTile50.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile50.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile49_MouseLeave(object sender, EventArgs e)
        {
            metroTile49.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile49.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile53_MouseLeave(object sender, EventArgs e)
        {
            metroTile53.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile53.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile64_MouseLeave(object sender, EventArgs e)
        {
            metroTile64.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile64.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile65_MouseLeave(object sender, EventArgs e)
        {
            metroTile65.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile65.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile66_MouseLeave(object sender, EventArgs e)
        {
            metroTile66.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile66.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile50_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile50.Style = MetroFramework.MetroColorStyle.Green;
            metroTile50.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile49_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile49.Style = MetroFramework.MetroColorStyle.Green;
            metroTile49.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile53_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile53.Style = MetroFramework.MetroColorStyle.Orange;
            metroTile53.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile64_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile64.Style = MetroFramework.MetroColorStyle.Green;
            metroTile64.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile65_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile65.Style = MetroFramework.MetroColorStyle.Green;
            metroTile65.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile66_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile66.Style = MetroFramework.MetroColorStyle.Green;
            metroTile66.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile94_MouseLeave(object sender, EventArgs e)
        {
            metroTile94.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile94.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile95_MouseLeave(object sender, EventArgs e)
        {
            metroTile95.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile95.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile96_MouseLeave(object sender, EventArgs e)
        {
            metroTile96.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile96.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile100_MouseLeave(object sender, EventArgs e)
        {
            metroTile100.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile100.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile102_MouseLeave(object sender, EventArgs e)
        {
            metroTile102.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile102.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile99_MouseLeave(object sender, EventArgs e)
        {
            metroTile99.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile99.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile101_MouseLeave(object sender, EventArgs e)
        {
            metroTile101.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile101.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile98_MouseLeave(object sender, EventArgs e)
        {
            metroTile98.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile98.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile94_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile94.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile94.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile95_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile95.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile95.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile96_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile96.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile96.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile100_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile100.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile100.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile102_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile102.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile102.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile99_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile99.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile99.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile101_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile101.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile101.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile98_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile98.Style = MetroFramework.MetroColorStyle.Blue;
            metroTile98.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("en-US"));
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("en-US"));
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("ru-RU"));
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("ru-RU"));
        }

        private void metroTile77_MouseLeave(object sender, EventArgs e)
        {
            metroTile77.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile77.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile77_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile77.Style = MetroFramework.MetroColorStyle.Green;
            metroTile77.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void metroTile103_Click(object sender, EventArgs e)
        {
            string st, s1, s2, s3; ;
            st = textBox10.Text.ToUpper();
            textBox11.Text = Tr2(st);
            textBox11.Text = textBox11.Text.ToLower();

            
            s1 = textBox10.Text;
            s2 = textBox11.Text;
            s3 = textBox12.Text;

            //for (int i = 0; i < s2.Length; i++)
            //   {
            //    s3 = s3 + s2;
            string[] mystring = s2.Split();
            // s3 = mystring[0].ToString();

            textBox13.Text = mystring[0].ToString();
            textBox12.Text = mystring[1].ToString();
            textBox14.Text = mystring[2].ToString();
            textBox15.Text = textBox13.Text + "." + textBox12.Text[0] + textBox14.Text[0];


        }

        private void metroTile103_MouseLeave(object sender, EventArgs e)
        {
            metroTile103.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile103.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile103_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile103.Style = MetroFramework.MetroColorStyle.Green;
            metroTile103.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile104_MouseLeave(object sender, EventArgs e)
        {
            metroTile104.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile104.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile104_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile104.Style = MetroFramework.MetroColorStyle.Green;
            metroTile104.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile104_Click(object sender, EventArgs e)
        {
            string fio, s11, s12, log_in, tire, prob,ou, mail, grup, result, dolzhnost, otdel , tel;
            fio = textBox10.Text+ "\r\n"; 
            ou = textBox20.Text+ "\r\n";
            mail = "Почта - "+ label70.Text+"\r\n"; 
            grup = label75.Text + "\r\n";
            tire = "---------------------------------------------------------- \r\n";
            prob = " \r\n";
            log_in = textBox15.Text+ "\r\n"; 
            result = richTextBox1.Text + "\r\n";
            dolzhnost = textBox21.Text + "\r\n";
            otdel = textBox22.Text + "\r\n";
            tel = textBox23.Text + "\r\n";
            richTextBox1.Text = "";
            richTextBox1.AppendText("Добрый день, \r\n");
            richTextBox1.AppendText("Необходимо создать пользователя: :\r\n");
            richTextBox1.AppendText(tire);
            richTextBox1.AppendText(prob);
            richTextBox1.AppendText("ФИО - "+fio);
            richTextBox1.AppendText(@"Логин - sib\" + log_in);
            richTextBox1.AppendText(prob);
            richTextBox1.AppendText(tire);
            richTextBox1.AppendText(prob);
            richTextBox1.AppendText(ou);
            richTextBox1.AppendText(mail);
            richTextBox1.AppendText(prob);
            richTextBox1.AppendText(tire);

            richTextBox1.AppendText("Компания - " + comboBox1.Text+ "\r\n");
            richTextBox1.AppendText("Должность - "+dolzhnost);
            richTextBox1.AppendText("Отдел - " + otdel);
            richTextBox1.AppendText("тел - " + tel);

            richTextBox1.AppendText("Группы:  \r\n");
            richTextBox1.AppendText(grup);

        }

        private void metroTile105_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox10.Text);
            MessageBox.Show("в буфер " + textBox10.Text, "ФИО", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile106_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox11.Text);
            MessageBox.Show("в буфер " + textBox11.Text, "FIO"+ textBox11.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile107_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox12.Text);
            MessageBox.Show("в буфер " + textBox12.Text, "Имя ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile108_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox13.Text);
            MessageBox.Show("в буфер " + textBox13.Text, "фамилия ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile109_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox14.Text);
            MessageBox.Show("в буфер " + textBox14.Text, "отче-о ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile111_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox15.Text);
            MessageBox.Show("в буфер " + textBox15.Text, "login ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile112_MouseLeave(object sender, EventArgs e)
        {

            metroTile112.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile112.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile112_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile112.Style = MetroFramework.MetroColorStyle.Red;
            metroTile112.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile112_Click(object sender, EventArgs e)
        {
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            richTextBox1.Text = "";
            textBox20.Text = "";
            textBox21.Text = "";
            textBox22.Text = "";
            textBox23.Text = "";
        }

        private void metroTile113_Click(object sender, EventArgs e)
        {
            try
            {
                String dirPR_86 = @"\\" + label2.Text + @"\c$\Program Files (x86)";
                String dirPR = @"\\" + label2.Text + @"\c$\Program Files";

                String dirComfPR = @"\\" + label2.Text + @"\c$\Program Files\Comfort";
                String dirComfPR86 = @"\\" + label2.Text + @"\c$\Program Files (x86)\Comfort";

                if (!Directory.Exists(dirPR_86))
                {
                    Directory.Delete(dirComfPR,true);
                    FileInfo user_profile = new FileInfo(@"\\" + label2.Text + @"\c$\Users\Public\Desktop\Комфорт.lnk");
                    user_profile.Delete();

                    MessageBox.Show(@"delete Comfort from X86 System ", "del comf", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                else
                {
                    Directory.Delete(dirComfPR86, true);
                    FileInfo user_profile = new FileInfo(@"\\" + label2.Text + @"\c$\Users\Public\Desktop\Комфорт.lnk");
                    user_profile.Delete();

                    MessageBox.Show(@"delete Comfort from X64 System ", "del comf", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

        }
             catch (Exception)
               {

                  MessageBox.Show(@"err 4404 ", "copy", MessageBoxButtons.OK, MessageBoxIcon.Error);

               }
}

        private void metroTile114_Click(object sender, EventArgs e)
        {
            // /U sib\okhten
            try
            {
                //Process.Start("cmd.exe", "/K ping -t " + label2.Text);
                Process.Start("cmd.exe", "/K taskkill /F /S " + label2.Text + " /IM " + textBox16.Text);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile114_MouseLeave(object sender, EventArgs e)
        {
            metroTile114.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile114.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile114_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile114.Style = MetroFramework.MetroColorStyle.Green;
            metroTile114.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile78_MouseLeave(object sender, EventArgs e)
        {
            metroTile78.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile78.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile78_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile45.Style = MetroFramework.MetroColorStyle.Green;
            metroTile45.TileTextFontSize = MetroFramework.MetroTileTextSize.Medium;
        }

        private void metroTile75_Click_1(object sender, EventArgs e)
        {
            metroTile75.Enabled = false;
            string s1,s2,startPath = Application.StartupPath;
            label54.Text = "";
            try
            {
                Process proc = Process.Start(new ProcessStartInfo 
                {
                   FileName = Environment.ExpandEnvironmentVariables("%windir%\\sysnative\\cmd.exe"),
                    //FileName = "cmd",
                    Arguments = "/c query session /server:dctempSIBterm",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    Verb = "runas",
                    StandardOutputEncoding = Encoding.GetEncoding(866)
                });
                s1 = proc.StandardOutput.ReadToEnd();

                Process proc2 = Process.Start(new ProcessStartInfo
                {
                    FileName = Environment.ExpandEnvironmentVariables("%windir%\\sysnative\\cmd.exe"),
                    Arguments = "/c query session /server:dctempSIB2term",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    Verb = "runas",
                    StandardOutputEncoding = Encoding.GetEncoding(866)

                });
                s2 = proc2.StandardOutput.ReadToEnd();
                label54.Text = "dctempSIBterm"+s1 + "dctempSIB2term" + s2;
            }
            catch (Exception)
            {
                MessageBox.Show(@"query", "query err", MessageBoxButtons.OK, MessageBoxIcon.Error);
                metroTile75.Enabled = true;
            }
            metroTile75.Enabled = true;
        }

        private void metroTile75_MouseLeave(object sender, EventArgs e)
        {
            metroTile75.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile75.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile75_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile75.Style = MetroFramework.MetroColorStyle.Green;
            metroTile75.TileTextFontSize = MetroFramework.MetroTileTextSize.Medium;
        }

        private void metroTile115_Click(object sender, EventArgs e)
        {
            string s;

            if (checkBox2.Checked==true)
            {
                s = " /v:DCTEMPSIBTERM /control";
                try
                {
                    Process.Start("Mstsc", " /shadow:" + textBox17.Text + s);
                }
                catch (Exception)
                {
                    MessageBox.Show(@"error", "mstsc err", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (checkBox3.Checked == true)
            {
                s = " /v:DCTEMPSIB2TERM /control";
                try
                {
                    Process.Start("Mstsc", " /shadow:" + textBox17.Text + s);
                }
                catch (Exception)
                {
                    MessageBox.Show(@"error", "mstsc err", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            //Mstsc /shadow:24 /v:DCTEMPSIBTERM /control

                   
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void metroTile115_MouseLeave(object sender, EventArgs e)
        {
            metroTile115.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile115.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile115_MouseMove(object sender, MouseEventArgs e)
        {
            metroTile115.Style = MetroFramework.MetroColorStyle.Green;
            metroTile115.TileTextFontSize = MetroFramework.MetroTileTextSize.Medium;
        }

        private void checkBox4_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked==true)
            {
               // webBrowser1.
            }
        }

        private void metroTile116_Click(object sender, EventArgs e)
        {
            label54.Text = "";
            textBox17.Text = "";
        }

        private void metroTile78_Click_1(object sender, EventArgs e)
        {
            string ncomp, output, startPath = Application.StartupPath;
            label55.Text = "";
            ncomp = label2.Text;
            
            Process proc = Process.Start(new ProcessStartInfo
            {
               // FileName = Environment.ExpandEnvironmentVariables("%windir%\\sysnative\\cmd.exe"),
                FileName = "cmd",
                Arguments = "/c ping " + ncomp,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                Verb = "runas",
                StandardOutputEncoding = Encoding.GetEncoding(866)
            });
            //label55.Text = proc.StandardOutput.ReadToEnd();
            //proc.BeginErrorReadLine();

             proc.OutputDataReceived += (s, e1) => { label55.Text += e1.Data+"\n"; };
            proc.BeginOutputReadLine();
            proc.Close();
        }

        private void metroTile78_MouseLeave_1(object sender, EventArgs e)
        {
            metroTile78.Style = MetroFramework.MetroColorStyle.Silver;
            metroTile78.TileTextFontSize = MetroFramework.MetroTileTextSize.Small;
        }

        private void metroTile78_MouseMove_1(object sender, MouseEventArgs e)
        {
            metroTile78.Style = MetroFramework.MetroColorStyle.Green;
            metroTile78.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
        }

        private void metroTile118_Click(object sender, EventArgs e)
        {
            string ou, user_fio,user_company; 
            int ou_count, lb2_count;
               string path = @"C:\report\";
                DirectoryInfo dirInfo = new DirectoryInfo(path);
            int y,z;
                z = 0;
            
            if (textBox18.Text != "")
            {
                z = Convert.ToInt16(textBox18.Text);
            }
            try
            {
                // listBox1.SelectedIndex = listBox1.Items.Count - 1; //колличество элементов (последний)
                //listBox1.SelectedIndex = 1;
                
                int lb1_count = listBox1.Items.Count; //колличкство строк listbox1

                 for (y=z; y<=lb1_count;y++) // цикл перебора пользователей
                 {
                    listBox1.SelectedIndex = y;
                    ou = textBox7.Text;
                    user_fio = "";
                    lb2_count = 0;
                    /// начало цикла отчет по пользователю
                    if (ou != "Unknown")
                {
                    ou_count = ou.Length;
                    //  textBox18.Text = ou_count.ToString();
                    for (int i = ou_count - 1; i != 0; i--)
                    {
                        if (ou[i].ToString() == @"/") // поиск слеша с конца
                        {
                            i = i + 1;
                            for (int j = i; j != (ou_count); j++)
                            {
                                user_fio = user_fio + ou[j].ToString();
                            }
                            break;
                        }
                        //  else
                        //    {
                        //          textBox18.Text = textBox18.Text + ou[i].ToString();
                        //        break;
                        //    }
                    }// конец цикла вырезания фио

                    //начало поиск дат с 23.03.2020 по 30.04.2020
                    listBox2.SelectedIndex = 0; //переход на верхнюю строку
                    lb2_count = listBox2.Items.Count; //колличкство строк listbox2

                    //
                    //------------------------------Создание файла------------------------------


                    user_company = textBox8.Text;  /// 
                    user_company = user_company.Replace(@"""", "");
                    StreamWriter file = new StreamWriter(path + user_fio + " " + user_company + ".txt"); // Создание файла
                    file.WriteLine(user_fio + " " + user_company); // первая строка файла
                    file.WriteLine(" ");
                    file.WriteLine("------------------------------------------");
                    file.WriteLine("воход за период с 23.03.2020 по 31.03.2020");
                    DateTime dubele_date; /// переменная убрать даты двойники
                    dubele_date = new DateTime(2000, 01, 01);
                    int count_vhod = 0; //колличество полезных строк. входов(логинов)
                    //-----------------ПЕРВЫЙ ПЕРИОД
                    for (int i = 0; i < lb2_count; i++)
                    {

                        string str_lb2 = listBox2.Items[i].ToString(); //строка в переменную
                        str_lb2 = str_lb2.Substring(str_lb2.IndexOf(",") + 2); // обрезать строку до запятой
                        int indexOfChar = str_lb2.IndexOf(" ");
                        int index_end = str_lb2.Length;

                        string str_lb2_result = str_lb2.Remove(str_lb2.IndexOf(','), str_lb2.Length - str_lb2.IndexOf(',')); //обрезать конец строки
                        str_lb2_result = str_lb2_result.TrimEnd(' ');//удалить пробелы
                                                        
                                DateTime data_lb2 = Convert.ToDateTime(str_lb2_result);// преобразовать полученную строку в дату
                            
                            
                            
                            

                        DateTime datestart = new DateTime(2020, 03, 23); //начало периода
                        DateTime dateend = new DateTime(2020, 03, 31); //конец

                        //dubele_date = data_lb2;
                        //if (i==0)
                        ////  {
                        ///   dubele_date = new DateTime(2000, 01, 01);
                        // }
                        if ((data_lb2 >= datestart) & (data_lb2 <= dateend) & (data_lb2 != dubele_date))
                        {
                            //textBox18.Text = data_lb2.ToString();
                            count_vhod++;
                            file.WriteLine(str_lb2_result);
                            dubele_date = data_lb2;
                        }
                    }
                    file.WriteLine("Итого дней за период:" + count_vhod.ToString());
                    file.WriteLine(" ");
                    file.WriteLine("------------------------------------------");
                    //------------------------ВТОРОЙ ПЕРИОД------------------------------------------------------------------------------
                    file.WriteLine("воход за период с 01.04.2020 по 30.04.2020");

                    listBox2.SelectedIndex = 0;
                    dubele_date = new DateTime(2000, 01, 01);
                    count_vhod = 0; //колличество полезных строк. входов(логинов)
                    for (int j = 0; j < lb2_count; j++)
                    {

                        string str_pr2_lb2 = listBox2.Items[j].ToString(); //строка в переменную
                        str_pr2_lb2 = str_pr2_lb2.Substring(str_pr2_lb2.IndexOf(",") + 2); // обрезать строку до запятой
                        int indexOfChar2 = str_pr2_lb2.IndexOf(" ");
                        int index_end2 = str_pr2_lb2.Length;

                        string str_pr2_lb2_result = str_pr2_lb2.Remove(str_pr2_lb2.IndexOf(','), str_pr2_lb2.Length - str_pr2_lb2.IndexOf(',')); //обрезать конец строки
                        str_pr2_lb2_result = str_pr2_lb2_result.TrimEnd(' ');//удалить пробелы
                        DateTime data_pr2_lb2 = Convert.ToDateTime(str_pr2_lb2_result);// преобразовать полученную строку в дату
                        DateTime datestartpr2_ = new DateTime(2020, 04, 01); //начало периода
                        DateTime dateendpr2_ = new DateTime(2020, 04, 30); //конец

                        //dubele_date = data_lb2;
                        //if (i==0)
                        ////  {
                        ///   dubele_date = new DateTime(2000, 01, 01);
                        // }
                        if ((data_pr2_lb2 >= datestartpr2_) & (data_pr2_lb2 <= dateendpr2_) & (data_pr2_lb2 != dubele_date))
                        {
                            //textBox18.Text = data_lb2.ToString();
                            count_vhod++;
                            file.WriteLine(str_pr2_lb2_result);
                            dubele_date = data_pr2_lb2;
                        }


                    }
                    file.WriteLine("Итого дней за период:" + count_vhod.ToString());
                    file.WriteLine(" ");
                    file.WriteLine("------------------------------------------");

                    //-------------------------------Оригинал
                    file.WriteLine("оригинал входы ");

                    listBox2.SelectedIndex = 0;
                    dubele_date = new DateTime(2000, 01, 01);
                    count_vhod = 0; //колличество полезных строк. входов(логинов)
                    for (int x = 0; x < lb2_count; x++)
                    {
                        string str_org_lb2_org = listBox2.Items[x].ToString(); //строка в переменную \ вторая переменная для полной строки
                        string str_org_lb2 = listBox2.Items[x].ToString(); //строка в переменную
                        str_org_lb2 = str_org_lb2.Substring(str_org_lb2.IndexOf(",") + 2); // обрезать строку до запятой
                        int indexOfChar3 = str_org_lb2.IndexOf(" ");
                        int index_end3 = str_org_lb2.Length;

                        string str_org_lb2_result = str_org_lb2.Remove(str_org_lb2.IndexOf(','), str_org_lb2.Length - str_org_lb2.IndexOf(',')); //обрезать конец строки
                        str_org_lb2_result = str_org_lb2_result.TrimEnd(' ');//удалить пробелы
                        DateTime data_pr2_lb2 = Convert.ToDateTime(str_org_lb2_result);// преобразовать полученную строку в дату
                        DateTime datestartpr2_ = new DateTime(2020, 03, 23); //начало периода
                        DateTime dateendpr2_ = new DateTime(2020, 04, 30); //конец

                        //dubele_date = data_lb2;
                        //if (i==0)
                        ////  {
                        ///   dubele_date = new DateTime(2000, 01, 01);
                        // }
                        if ((data_pr2_lb2 >= datestartpr2_) & (data_pr2_lb2 <= dateendpr2_) & (data_pr2_lb2 != dubele_date))
                        {
                            //textBox18.Text = data_lb2.ToString();
                            count_vhod++;
                            file.WriteLine(str_org_lb2_org);
                            dubele_date = data_pr2_lb2;
                        }


                    }

                    file.WriteLine(" ");
                    file.WriteLine("-----------------конец-------------------------");


                    file.Close();
                    //--------------------------------------------------------------------
                }///конец цикла отчет по пользователю далее переход к следующему
            }//конец цикла перебора пользователей
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
                textBox18.Text = listBox1.SelectedIndex.ToString();
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                string v;
                int iv;
                v = listBox1.SelectedIndices.Count.ToString();
                label1.Text = v;
                iv = listBox1.SelectedIndices.Count;

                if (v == "1")
                {

                    // textBox2.Text = "";
                    //listBox2.ClearSelected();
                    listBox2.Items.Clear();
                    //   label2.Text = "";
                    download_log();
                    get_user_info();
                }

                else
                {
                    listBox1.SelectedItems.Clear();
                    listBox2.Items.Clear();
                    label1.Text = "";

                }
            }
        }

        private void metroTile119_Click(object sender, EventArgs e)
        {
            textBox19.Text = listBox1.Items.Count.ToString();
        }

        private void metroTile120_Click(object sender, EventArgs e)
        {
            string s;
            try
            {
                s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);

                Process.Start("compmgmt.msc", " /computer=" + s);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "compmgmt.msc", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile117_Click(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe ", "10.10.0.8");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail start chrome", "chrome pach err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void metroTile122_Click(object sender, EventArgs e)
        {
            string pas;
            pas ="d1nfVkghnvouw435rtpfpsd5fiesrjugpi2oerajgwpw";

            //HtmlElementCollection inputs;
            // HtmlElement body = webBrowser1.Document.Body;
            //  inputs = body.GetElementsByTagName("userName");
            //   inputs[0].SetAttribute("userName", "okhten");
            webBrowser1.Document.GetElementById("userName").InnerText = "okhten";
            webBrowser1.Document.GetElementById("Password").InnerText = "Valinor512512";
            webBrowser1.Document.GetElementById("Domain").InnerText = "sib";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void table1BindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void metroTile123_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox1.Text);
            MessageBox.Show("copy ok " ,  "" , MessageBoxButtons.OK, MessageBoxIcon.Information);
            
        }

        private void metroTile124_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox7.Text);
            MessageBox.Show("copy ok ", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile125_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("31415926535");
            MessageBox.Show("copy ok ", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void metroTile126_Click(object sender, EventArgs e)
        {
            string s;
            try
            {
                s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);

                
                Process.Start(msinfo_xml, @" /computer " + s);//
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "compmgmt.msc", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
    

