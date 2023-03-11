using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.DirectoryServices;

namespace lab2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://DC=VARCHENKO,DC=UA"))
            {
                using (DirectoryEntry u = AD.Children.Add("OU=SecurityTeam", "organizationalUnit"))
                {
                    u.CommitChanges();
                }
            }
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
            {
                using (DirectoryEntry u = AD.Children.Add("OU=ImportantUsers", "organizationalUnit"))
                {
                    u.CommitChanges();
                }
            }
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
            {
                using (DirectoryEntry u = AD.Children.Add("OU=Partners", "organizationalUnit"))
                {
                    u.CommitChanges();
                }
            }
            MessageBox.Show("OUs created");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=Partners,OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
            {
                string Name = "Oleksandr";
                for (int i = 1; i < 31; i++)
                {
                    string loginName = Name + i.ToString("00");
                    using (DirectoryEntry u = AD.Children.Add($"CN={loginName}", "user"))
                    {
                        u.Properties["displayName"].Add($"{loginName}");
                        u.Properties["userPrincipalName"].Add($"{loginName}" + "@VARCHENKO.UA");
                        u.Properties["sAMAccountName"].Add($"{loginName}");
                        u.Properties["department"].Add("Branch Office");
                        u.Properties["description"].Add("Very Goog User");
                        u.Properties["businessCategory"].Add("Manager");
                        u.Properties["businessCategory"].Add("HelpDesk");
                        u.Properties["businessCategory"].Add("Secretary");


                        u.CommitChanges();

                        SetPassword(u, "P@ssw0rd");
                        u.Properties["userAccountControl"].Value = "544";
                        u.CommitChanges();
                    }
                }

            }
            MessageBox.Show("Users successfully created");

        }
        private static void SetPassword(DirectoryEntry UE, string password)
        {
            object[] oPassword = new object[] { password };
            UE.Invoke("SetPassword", oPassword);
            UE.CommitChanges();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry OU = new DirectoryEntry("LDAP://OU=Partners,OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
            {
                foreach (DirectoryEntry u in OU.Children)
                {
                    if (u.Properties["objectClass"].Contains("user"))
                    {
                        u.Properties["Company"].Add("Kiev National University");
                        u.Properties["telephoneNumber"].Add("044-5260522");
                        u.Properties["VarchenkoDescription"].Add("Прогресивний працівник");
                        u.Properties["VarchenkoID"].Add("23456");
                        u.Properties["VarchenkoTaxID"].Add("8877665544");
                        u.Properties["VarchenkoCovidVaccinated"].Add(true);
                        u.Properties["otherMobile"].Add("067-2334383");
                        u.Properties["otherMobile"].Add("050-2206789");
                        u.Properties["otherMobile"].Add("063-2184545");
                        u.CommitChanges();
                    }
                }
            }
            MessageBox.Show("Attributes successfully updated");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
            {
                using (DirectoryEntry g = AD.Children.Add("CN=AdvancedSpecialists", "group"))
                {
                    g.Properties["sAMAccountName"].Add("AdvancedSpecialists");
                    g.CommitChanges();
                }
            }
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
            {
                using (DirectoryEntry g = AD.Children.Add("CN=AdvancedWorkers", "group"))
                {
                    g.Properties["sAMAccountName"].Add("AdvancedWorkers");
                    g.CommitChanges();
                }
            }
            using (DirectoryEntry g = new DirectoryEntry("LDAP://CN=AdvancedWorkers,OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
            {
                g.Properties["member"].Add("CN=AdvancedSpecialists,OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA");
                g.CommitChanges();
            }
            MessageBox.Show("Groups created successfully. Group AdvancedSpecialists added to members of group AdvancedWorkers.");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry g = new DirectoryEntry("LDAP://CN=AdvancedSpecialists,OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
            {
                using (DirectoryEntry OU = new DirectoryEntry("LDAP://OU=Partners,OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
                {
                    foreach (DirectoryEntry u in OU.Children)
                    {
                        if (u.Properties["objectClass"].Contains("user"))
                        {
                            g.Properties["member"].Add(u.Properties["distinguishedName"].Value.ToString());
                            g.CommitChanges();
                        }
                    }
                }
            }
            MessageBox.Show("All user of OU 'Partners' added to group AdvancedSpecialists");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            StringBuilder groupUsers = new StringBuilder();
            using (DirectoryEntry G = new DirectoryEntry("LDAP://CN=AdvancedSpecialists,OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
            {
                using (DirectoryEntry OU = new DirectoryEntry("LDAP://OU=Partners,OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA"))
                {
                    foreach (DirectoryEntry u in OU.Children)
                    {
                        if (G.Properties["member"].Contains(u.Properties["distinguishedName"].Value.ToString()))
                        {
                            groupUsers.Append(u.Properties["displayName"].Value.ToString()).Append("\n");
                        }
                    }
                }
            }
            MessageBox.Show(groupUsers.ToString());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (IsADAuthenticated(textBox1.Text, textBox2.Text)){
                MessageBox.Show("Authentication success");
            }
            else
            {
                MessageBox.Show("Authentication failed");
            }
        }
        public bool IsADAuthenticated(string L, string P)
        {
            try
            {
                using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=Partners,OU=ImportantUsers,OU=SecurityTeam,DC=VARCHENKO,DC=UA", L, P))
                {
                    DirectorySearcher S = new DirectorySearcher(AD);
                    S.Filter = "(SAMAccountName=" + L + ")"; S.PropertiesToLoad.Add("cn");
                    SearchResult R = S.FindOne();
                    if (R == null)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
        }
    }
      
}
