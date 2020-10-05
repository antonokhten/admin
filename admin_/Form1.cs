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
using Ionic.Zip;


namespace admin
{
    public partial class Form1 : Form
    {
        //  string dmwrPath = @"C:/Program Files (x86)/SolarWinds/DameWare Remote Support";
        string dmwrPath = Properties.Settings.Default.dmwrPath;
        string file_msc = Properties.Settings.Default.mmc;
        string msinfo_xml = Properties.Settings.Default.msinfo_xml;
        string acronis_xml = Properties.Settings.Default.acronis_xml;
        // system32\msinfo32.exe
        public void ps_user_info()
        {
            string s;


           
            //Clipboard.SetText();
            try
            {
                s = listBox1.SelectedItem.ToString();
                Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe ", @"-noexit -command import-module ActiveDirectory; Get-ADUser -identity " + s + " -properties Company,CanonicalName,telephoneNumber,whenCreated,LastLogonDate,HomeDirectory; Get-ADUser "+s+" -Properties Memberof | Select -ExpandProperty memberOf;" + @" Read-Host -Prompt" + @"""Press Enter to exit""");
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
                    Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe ", @" -command import-module ActiveDirectory; Add-ADGroupMember SCAN " + s );
                }

             }

            catch (Exception)
            {
                MessageBox.Show(@"error", "ps err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public Form1()
        {
          InitializeComponent();
           
        }
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
                    string path = @"\\sibsup\_LoginLog\userlog\" + fio + ".txt";

                    string[] arStr = File.ReadAllLines(path, Encoding.GetEncoding(1251));
                    listBox2.Items.AddRange(arStr);
                    listBox2.SelectedIndex = listBox2.Items.Count - 1;


                }
                catch (Exception)
                {
                    MessageBox.Show(@"error, wrong path \\sibsup\_LoginLog\userlog", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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
                        Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:" + ss + " -a:1 -x:");
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
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-dzr-j83k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j76k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-dzr-j31k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-dzr-j82k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-dzr-j22k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-dzr-j111k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-dzr-j2k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j6k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j1k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j1k3 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j10vk1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-j10k-nar -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j4k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-dzr-j44k2 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j7k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j8k -a:1 -x:");
        }

        private void button25_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-reu2-k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j9k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j35k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j16k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j3k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-do-kas -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button27_Click(object sender, EventArgs e)
        {
            string s;
            // s = label9.Text;
            s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);
            if (s != "") 
                    {
                Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:" + s + " -a:1 -x:");
                //Form1.ActiveForm.WindowState = FormWindowState.Minimized;
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-pers-j25k -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-perj-j73k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j6k2 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j7k4 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button29_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:sib-mks-j35k1 -a:1 -x:");
            Form1.ActiveForm.WindowState = FormWindowState.Minimized;
        }

        private void button31_Click(object sender, EventArgs e)
        {

            // string compPatch = @"\\sib-dzr-j111k1\c$";
            Process.Start("explorer.exe ", @" \\sib-dzr-j111k1\c$");



        }

        private void button38_Click(object sender, EventArgs e)
        {
            Process.Start("\\\\sibsup\\DISTR\\AA_v3.exe" );
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            
        }

        private void button45_Click(object sender, EventArgs e)
        {
            Process.Start(dmwrPath + "/DWRCC.exe", "-h -m:10.10.0. -a:1 -x:");
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
            Process.Start("explorer.exe ", @" \\sibsup\");
        }

        private void button48_Click_1(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @" \\sibstore\");
        }

        private void button46_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @" \\sibsup\distr");
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void button54_Click(object sender, EventArgs e)
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
            Process.Start("mstsc.exe"," /v:sibsup");
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:sibstore");
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:sibdc");
        }

        private void button31_Click_1(object sender, EventArgs e)
        {
            Process.Start("mstsc.exe", " /v:10.10.0.164");
        }

        private void button32_Click(object sender, EventArgs e)
        {
            Process.Start("cmd.exe");
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
                label1.Text ="";
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
                    string path = @"\\sibsup\_LoginLog\userlog\" + fio + ".txt";
                    
                    string[] arStr = File.ReadAllLines(path, Encoding.GetEncoding(1251));
                    listBox2.Items.AddRange(arStr);
                    listBox2.SelectedIndex = listBox2.Items.Count - 1;
                    
                    
                }
                catch (Exception)
                {
                    MessageBox.Show(@"error, wrong path \\sibsup\_LoginLog\userlog", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            else
            {
                MessageBox.Show(@"выбрано больше одного", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

        }

        private void button33_Click(object sender, EventArgs e)
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

        private void button2_Click_2(object sender, EventArgs e)
        {

            string  ncomp, s,ss;
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
                        Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:" + ss + " -a:1 -x:");
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

        private void button45_Click_1(object sender, EventArgs e)
        {
            string s;
            s = textBox1.Text;
            // Process.Start(dmwrPath + "/DWRCC.exe", "-c: -m:"+s+"  -a:1 -x:");
            Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:" + s + "  -a:1 -x:");
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
                        Process.Start(dmwrPath + "/DWRCC.exe", "-c: -h -m:" + ss + " -a:1 -x:");
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
            try
            {
                Process.Start("explorer.exe ", @" \\"+label2.Text+@"\c$");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error pach ", "select", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            int i = 1;
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
                    i = i++;
                    FileInfo f = new FileInfo(@"\\sibsup\UK\pu.exe");
                    f.CopyTo(dir + @"\pu.exe");
                    i = i++;
                    FileInfo d = new FileInfo(@"\\sibsup\UK\del.exe");
                    d.CopyTo(dir + @"\del.exe");
                    i = i++;
                    FileInfo p = new FileInfo(@"\\sibsup\UK\path.txt");
                    p.CopyTo(dir + @"\path.txt");
                    i = i++;
                    FileInfo l = new FileInfo(@"\\sibsup\UK\log.txt");
                    l.CopyTo(dir + @"\log.txt");
                    i = i++;
                    FileInfo n = new FileInfo(@"\\sibsup\UK\nircmd.exe");
                    n.CopyTo(dir + @"\nircmd.exe");
                    i = i++;
                    //FileInfo lnk = new FileInfo(@"\\sibsup\UK\pu.lnk");
                    //lnk.CopyTo(@"\\sibsup\ + @"\pu.lnk");

                    DirectoryInfo directory = new DirectoryInfo(dir);
                    i = i++;
                    FileSystemAccessRule rule = new FileSystemAccessRule(@"sib\PU", FileSystemRights.FullControl, AccessControlType.Allow);
                    i = i++;
                    DirectorySecurity security = new DirectorySecurity();
                    i = i++;
                    security.AddAccessRule(rule);
                    i = i++;
                    directory.SetAccessControl(security);
                    i = i++;
                    MessageBox.Show(@" copy - ok ", "copy", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
            }
            catch (Exception)
            {
               
                    MessageBox.Show(@"error copy "+i.ToString(), "copy", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }
        }

        private void button10_Click_1(object sender, EventArgs e)
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

        private void button11_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(label2.Text);

        }

        private void button12_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(listBox1.SelectedItem.ToString());
        }

        private void button5_Click_2(object sender, EventArgs e)
        {
            //C: \Users\okhten\Desktop\putty.exe
            Process.Start("C:/Users/okhten/Desktop/putty.exe");
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            Form1.ActiveForm.Close();
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            
            Process.Start("C:/totalcmd/TOTALCMD64.EXE");
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @" C:\");
        }

        private void button16_Click_1(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @" H:\");
        }

        private void button17_Click_1(object sender, EventArgs e)
        {
            //%windir%\system32\dhcpmgmt.msc
            //Process.Start(@"");
            Process.Start(file_msc);
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

            Process.Start(@"C:\Program Files (x86)\VMware\Infrastructure\Virtual Infrastructure Client\Launcher\VpxClient.exe");
        }

        private void button21_Click_1(object sender, EventArgs e)
        {

            Process.Start(msinfo_xml, @" /computer "+label2.Text);


        }

        private void button22_Click_1(object sender, EventArgs e)
        {
            Process.Start(acronis_xml);
        }

        private void button23_Click_1(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @"\\sibsup\OBMEN_SIB");
        }

        private void button24_Click_1(object sender, EventArgs e)
        {
            
                 Process.Start("explorer.exe ", @"\\dcfilesrv\IT\10.0_SIB");
        }

        private void button25_Click_1(object sender, EventArgs e)
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

        private void button28_Click_1(object sender, EventArgs e)
        {

            string s;

            //  s = label9.Text;
            s = Convert.ToString(table1DataGridView[5, table1DataGridView.CurrentRow.Index].Value);
            try
            {
                Process.Start("cmd.exe", @"/K psexec \\" + s+@" cmd.exe");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button26_Click_1(object sender, EventArgs e)
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

        private void button29_Click_1(object sender, EventArgs e)
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

        private void button34_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @"\\sibsup\_LoginLog");

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
                label1.Text ="";
                
            }
}

        private void button40_Click(object sender, EventArgs e)
        {
            ps_user_info();
        }

        private void button30_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe ", @"Get-ADUser  -filter {Enabled -eq $True} -properties Description,LastLogonDate | where {$_.lastlogondate -lt (get-date).addmonths(-2)} | sort LastLogonDate | FT Name, LastLogonDate,Description;" + @" Read-Host -Prompt"+@"""Press Enter to exit""");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button42_Click(object sender, EventArgs e)
        {
            //C:\Program Files\Internet Explorer\iexplore.exe
            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button41_Click(object sender, EventArgs e)
        {

        }

        private void button41_Click_1(object sender, EventArgs e)
        {
            //% SystemRoot %\system32\WindowsPowerShell\v1.0\powershell.exe
            try
            {
                Process.Start(@"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "powershell", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load_1(object sender, EventArgs e)
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
        }

        private void button43_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("cmd.exe", @" psexec \\" + label2.Text+@" gpupdate /force");
            }
            catch (Exception)
            {
                MessageBox.Show(@"error", "ping err", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            try
            {
                    string in_p,out_p, user_n,cmd_com;
                user_n = listBox1.SelectedItem.ToString() + ".zip";
                in_p = textBox6.Text;
                   out_p = @"\\10.10.0.27\HOME_disable_user_SIBSUP\" + listBox1.SelectedItem.ToString() + ".zip";

                // out_p = @"H:\" + listBox1.SelectedItem.ToString() + ".zip";
                //     ZipFile zf = new ZipFile(out_p);
                //      zf.AddDirectory(in_p);
                //        zf.Save();
                //      MessageBox.Show("Архивация прошла успешно.", "Выполнено");
                //  Process.Start(@"C:\Program Files\7-Zip\7z.exe" , " a -tzip -ssw -mx9 -r0 " + out_p+"   "+ in_p);
                cmd_com = (@"""C:\Program Files\7-Zip\7z.exe""  a -tzip -ssw -mx9 -r0 " + out_p + "   " + in_p+" && pause");
                Process.Start("cmd.exe" ,"/K "+cmd_com);
              //  MessageBox.Show(cmd_com, "profile save zip", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception)
            {
                MessageBox.Show(@"error ", "profile save zip", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button50_Click(object sender, EventArgs e)
        {

        }

        private void button51_Click(object sender, EventArgs e)
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

        private void button52_Click(object sender, EventArgs e)
        {
            if (textBox6.Text != "")
            {
                Process.Start("explorer.exe ", textBox6.Text);
            }
        }

        private void button53_Click(object sender, EventArgs e)
        {
            Process.Start("excel.exe", @"""\\dcfilesrv\DOC\телефонная книга.xlsx""");
        }

        private void button55_Click(object sender, EventArgs e)
        {
            textBox1.Text = "10.10.";
        }

        private void button35_Click_1(object sender, EventArgs e)
        {
        
            Clipboard.SetText(@"sib\"+listBox1.SelectedItem.ToString());
        }

        private void button36_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(@"dcxenapp.guo.local" );
        }

        private void button37_Click_1(object sender, EventArgs e)
        {
          //  Process.Start("printmanagement.msc");
            // printmanagement.msc
        }

        private void button37_Click_2(object sender, EventArgs e)
        {
            Process.Start("printmanagement.msc");
        }

        private void button39_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(@"printmanagement.msc");
        }

        private void button56_Click(object sender, EventArgs e)
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

        private void button57_Click(object sender, EventArgs e)
        {
            
            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe",  "http://files.regenergy.ru/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button58_Click(object sender, EventArgs e)
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

        private void button49_Click_1(object sender, EventArgs e)
        {

        }

        private void button59_Click(object sender, EventArgs e)
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

        private void button60_Click(object sender, EventArgs e)
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

        private void button49_Click_2(object sender, EventArgs e)
        {
                                
            try
            {
                //textBox6.Text;

                String dir_work =textBox6.Text;
                String dir_arch = @"\\10.10.0.27\HOME_disable_user_SIBSUP\" + listBox1.SelectedItem.ToString() + ".zip";
                if (!System.IO.File.Exists(dir_arch))
                {
                    MessageBox.Show(@"archive Exist " +dir_arch, "del profile", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            try
            {

           String user = listBox1.SelectedItem.ToString();
            String dir = @"\\sibstore\e$\scan\" + user;
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                    DirectoryInfo directory = new DirectoryInfo(dir);
                    FileSystemAccessRule rule_Allow = new FileSystemAccessRule(@"sib\" + user, FileSystemRights.Read | FileSystemRights.ListDirectory | FileSystemRights.WriteData | FileSystemRights.ExecuteFile | FileSystemRights.Delete, InheritanceFlags.ObjectInherit,
            PropagationFlags.None, AccessControlType.Allow);
                    FileSystemAccessRule rule_Deny = new FileSystemAccessRule(@"sib\" + user,  FileSystemRights.Delete, AccessControlType.Deny);

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

        private void button62_Click(object sender, EventArgs e)
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
                    fn.CopyTo(dir+ @"\HCS_Client.exe");
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

        private void button63_Click(object sender, EventArgs e)
        {
            String user = listBox1.SelectedItem.ToString();
            String dir = @"\\sibstore\SCAN\" + user;

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

        private void button64_Click(object sender, EventArgs e)
        {
            String user = listBox1.SelectedItem.ToString();
            String dir = @"\\sibstore\SCAN\" + user;

            Process.Start("explorer.exe ", dir);
          
        }

        private void button66_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(@"sib\scansib");
        }

        private void button65_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("E;fcysqgfhjkm");
        }

        private void button67_Click(object sender, EventArgs e)
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

        private void button68_Click(object sender, EventArgs e)
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

        private void button57_Click_1(object sender, EventArgs e)
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

        private void button69_Click(object sender, EventArgs e)
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

        private void button70_Click(object sender, EventArgs e)
        {
            String dir = @"https://ess-p.ru";
            Clipboard.SetText(dir);
        }

        private void button71_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe ", @"\\dcfilesrv");
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
            try
            {
                Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe", "https://10.10.0.7/");
            }
            catch (Exception)
            {
                MessageBox.Show(@"fail", "iexplorer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button74_Click_1(object sender, EventArgs e)
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

        private void button75_Click_1(object sender, EventArgs e)
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

        private void button76_Click(object sender, EventArgs e)
        {
            Process.Start(@"\\sibsup\DISTR\okhten\Zuma Deluxe\Zuma.exe");
        }

        private void button77_Click(object sender, EventArgs e)
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

        

       
    }
}
    

