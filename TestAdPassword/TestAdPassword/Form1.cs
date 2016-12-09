using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestAdPassword
{
    public class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            this.passwordChgTextBox.Enabled = false;
            this.passwordChgReTextBox.Enabled = false;
            this.changeButton.Enabled = false;

            this.Load += Form1_Load;
            this.connectButton.Click += btnTest_Click;
            this.changeButton.Click += btnChg_Click;
        }

        void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                this.domainTextBox.Focus();
            }
            catch (Exception ex)
            {
                WriteToTextBox(ex);
            }
        }

        void btnTest_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                if (button.Text.Equals("Connect"))
                {
                    try
                    {

                        if (string.IsNullOrEmpty(this.domainTextBox.Text.Trim()))
                        {
                            MessageBox.Show("도메인명을 입력하세요", "알림");
                            this.domainTextBox.Focus();
                        }
                        else if (string.IsNullOrEmpty(this.adminUsernameTextBox.Text.Trim()))
                        {
                            MessageBox.Show("관리자 계정이름을 입력하세요.", "알림");
                            this.adminUsernameTextBox.Focus();
                        }
                        else if (string.IsNullOrEmpty(this.adminPasswordTextBox.Text.Trim()))
                        {
                            MessageBox.Show("관리자 비밀번호를 입력하세요.", "알림");
                            this.adminPasswordTextBox.Focus();
                        }
                        else if (string.IsNullOrEmpty(this.lookupUsernameTextBox.Text.Trim()))
                        {
                            MessageBox.Show("비밀번호를 변경할 사용자 계정이름을 입력하세요.", "알림");
                            this.lookupUsernameTextBox.Focus();
                        }
                        else
                        {
                            string adDomain = this.domainTextBox.Text.Trim();
                            string userName = this.adminUsernameTextBox.Text.Trim();
                            string userPassword = this.adminPasswordTextBox.Text.Trim();
                            string lookupUsername = this.lookupUsernameTextBox.Text.Trim();

                            Connect(adDomain, userName, userPassword, lookupUsername);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.passwordChgTextBox.Enabled = false;
                        this.passwordChgReTextBox.Enabled = false;
                        this.changeButton.Enabled = false;

                        WriteToTextBox(ex);
                    }
                }
                else
                {
                    Disconnect();
                }

            }
        }

        void btnChg_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(this.passwordChgTextBox.Text.Trim()))
                {
                    MessageBox.Show("변경할 비밀번호를 입력하세요", "알림");
                    this.passwordChgTextBox.Focus();
                }
                else if (string.IsNullOrEmpty(this.passwordChgTextBox.Text.Trim()))
                {
                    MessageBox.Show("변경할 비밀번호를 입력하세요.", "알림");
                    this.passwordChgReTextBox.Focus();
                }
                else if (!this.passwordChgTextBox.Text.Trim().Equals(this.passwordChgReTextBox.Text.Trim()))
                {
                    MessageBox.Show("변경할 비밀번호를 확인하세요.\r\n변경할 비밀번호와 다시 입력한 변경할 비밀번호는 동일해야 합니다.", "알림");
                    this.passwordChgReTextBox.Focus();
                }
                else
                {

                    string adDomain = this.domainTextBox.Text.Trim();
                    string userName = this.adminUsernameTextBox.Text.Trim();
                    string userPassword = this.adminPasswordTextBox.Text.Trim();
                    string lookupUsername = this.lookupUsernameTextBox.Text.Trim();

                    string sMessage = string.Empty;

                    DirectoryEntry adUser = this.GetUserInfo(adDomain, userName, userPassword, lookupUsername, out sMessage);

                    if (adUser == null || !string.IsNullOrEmpty(sMessage))
                    {
                        throw new Exception(sMessage);
                    }

                    foreach (string propertyName in adUser.Properties.PropertyNames)
                    {
                        string oneNode = String.Format("{0}: {1}", propertyName, adUser.Properties[propertyName][0]);
                        Console.WriteLine(oneNode);
                        WriteToTextBox("Found User");
                        WriteToTextBox(oneNode);
                    }

                    string chgPassword = this.passwordChgTextBox.Text.Trim();

                    int val = (int)adUser.Properties["userAccountControl"].Value;
                    adUser.Properties["userAccountControl"].Value = val & ~0x2;
                    adUser.Properties["pwdLastSet"].Value = 0;
                    adUser.CommitChanges();

                    adUser.Invoke("SetPassword", chgPassword);
                    adUser.CommitChanges();

                    WriteToTextBox("Success");
                    WriteToTextBox("비밀번호가 변경되었습니다:");

                }
            }
            catch (Exception ex)
            {
                WriteToTextBox(ex);
            }
        }

        void Connect(string domain, string adminUsername, string adminPassword, string lookupUsername)
        {
            try
            {

                string sMessage = string.Empty;

                DirectoryEntry adUser = this.GetUserInfo(domain, adminUsername, adminPassword, lookupUsername, out sMessage);
                if (adUser == null || !string.IsNullOrEmpty(sMessage))
                {
                    throw new ApplicationException(sMessage);
                }

                foreach (string propertyName in adUser.Properties.PropertyNames)
                {
                    string oneNode = String.Format("{0} : {1}", propertyName, adUser.Properties[propertyName][0]);

                    Console.WriteLine(oneNode);

                    WriteToTextBox("Found User");
                    WriteToTextBox(oneNode);
                    this.consoleTextBox.ScrollToCaret();
                }

                WriteToTextBox("[Test Success]");

                this.domainTextBox.Enabled = false;
                this.adminUsernameTextBox.Enabled = false;
                this.adminPasswordTextBox.Enabled = false;
                this.lookupUsernameTextBox.Enabled = false;
                this.connectButton.Text = "Disconnect";

                this.passwordChgTextBox.Enabled = true;
                this.passwordChgReTextBox.Enabled = true;
                this.changeButton.Enabled = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        void Disconnect()
        {
            this.domainTextBox.Enabled = true;
            this.adminUsernameTextBox.Enabled = true;
            this.adminPasswordTextBox.Enabled = true;
            this.lookupUsernameTextBox.Enabled = true;
            this.connectButton.Text = "Connect";

            this.passwordChgTextBox.ResetText();
            this.passwordChgTextBox.Enabled = false;

            this.passwordChgReTextBox.ResetText();
            this.passwordChgReTextBox.Enabled = false;
            this.changeButton.Enabled = false;
        }

        DirectoryEntry GetUserInfo(string domain, string adminUsername, string adminPassword, string lookUpUsername, out string message)
        {
            message = string.Empty;
            DirectoryEntry adUser = null;

            try
            {
                string[] atmpDomain = domain.Split('.');
                string userNameDomain = string.Empty;

                if (domain.Contains("@"))
                {
                    userNameDomain = adminUsername + "@" + atmpDomain;
                }
                else
                {
                    if (atmpDomain.Length > 0)
                    {
                        userNameDomain = atmpDomain[0] + @"\" + adminUsername;
                    }
                    else
                    {
                        throw new ApplicationException("도메인명 처리중 오류가 발생했습니다");
                    }
                }

                DirectoryEntry entryPC = new DirectoryEntry("LDAP://" + domain, userNameDomain, adminPassword);
                //entryPC.AuthenticationType = AuthenticationTypes.Secure;
                if (entryPC == null)
                {
                    throw new ApplicationException("사용자를 찾을 수 없습니다.");
                }

                DirectorySearcher search = new DirectorySearcher(entryPC);
                search.Filter = "(SAMAccountName=" + lookUpUsername + ")";
                search.PropertiesToLoad.Add("cn");
                SearchResult result = search.FindOne();
                adUser = result.GetDirectoryEntry();

                return adUser;
            }
            catch (Exception ex)
            {
                WriteToTextBox(ex);
                message = ex.Message;
                return adUser;
            }
        }

        void WriteToTextBox(string format, params object[] args)
        {
            string message = String.Format(format, args);
            consoleTextBox.AppendText(message);
            consoleTextBox.AppendText(Environment.NewLine);
            consoleTextBox.ScrollToCaret();
        }

        void WriteToTextBox(string message)
        {
            WriteToTextBox("{0}", message);
        }

        void WriteToTextBox(Exception ex)
        {
            WriteToTextBox("--[Exception: {0, 30}]-------------------------------", ex.GetType().FullName);

            WriteToTextBox("Message: {0}", ex.Message);
            WriteToTextBox("Stack trace: {0}{1}", Environment.NewLine, ex.StackTrace);

            if(ex.InnerException != null)
            {
                WriteToTextBox("--[Inner Exception: {0, 30}]-------------------------", ex.InnerException.GetType().FullName);
                WriteToTextBox("Message: {0}", ex.InnerException.Message);
                WriteToTextBox("Stack trace: {0}{1}", Environment.NewLine, ex.InnerException.StackTrace);
            }
            WriteToTextBox("--[End Exception]------------------------------------");
        }

        #region Designer.cs
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.domainTextBox = new System.Windows.Forms.TextBox();
            this.adminUsernameTextBox = new System.Windows.Forms.TextBox();
            this.adminPasswordTextBox = new System.Windows.Forms.TextBox();
            this.passwordChgReTextBox = new System.Windows.Forms.TextBox();
            this.passwordChgTextBox = new System.Windows.Forms.TextBox();
            this.consoleTextBox = new System.Windows.Forms.RichTextBox();
            this.connectButton = new System.Windows.Forms.Button();
            this.changeButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lookupUsernameTextBox = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // domainTextBox
            // 
            this.domainTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.domainTextBox.Location = new System.Drawing.Point(155, 20);
            this.domainTextBox.Name = "domainTextBox";
            this.domainTextBox.Size = new System.Drawing.Size(277, 20);
            this.domainTextBox.TabIndex = 0;
            // 
            // adminUsernameTextBox
            // 
            this.adminUsernameTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.adminUsernameTextBox.Location = new System.Drawing.Point(155, 49);
            this.adminUsernameTextBox.Name = "adminUsernameTextBox";
            this.adminUsernameTextBox.Size = new System.Drawing.Size(277, 20);
            this.adminUsernameTextBox.TabIndex = 1;
            // 
            // adminPasswordTextBox
            // 
            this.adminPasswordTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.adminPasswordTextBox.Location = new System.Drawing.Point(155, 87);
            this.adminPasswordTextBox.Name = "adminPasswordTextBox";
            this.adminPasswordTextBox.PasswordChar = '*';
            this.adminPasswordTextBox.Size = new System.Drawing.Size(277, 20);
            this.adminPasswordTextBox.TabIndex = 2;
            // 
            // passwordChgReTextBox
            // 
            this.passwordChgReTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.passwordChgReTextBox.Location = new System.Drawing.Point(134, 45);
            this.passwordChgReTextBox.Name = "passwordChgReTextBox";
            this.passwordChgReTextBox.PasswordChar = '*';
            this.passwordChgReTextBox.Size = new System.Drawing.Size(277, 20);
            this.passwordChgReTextBox.TabIndex = 12;
            // 
            // passwordChgTextBox
            // 
            this.passwordChgTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.passwordChgTextBox.Location = new System.Drawing.Point(134, 19);
            this.passwordChgTextBox.Name = "passwordChgTextBox";
            this.passwordChgTextBox.PasswordChar = '*';
            this.passwordChgTextBox.Size = new System.Drawing.Size(277, 20);
            this.passwordChgTextBox.TabIndex = 11;
            // 
            // consoleTextBox
            // 
            this.consoleTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.consoleTextBox.Location = new System.Drawing.Point(12, 290);
            this.consoleTextBox.Name = "consoleTextBox";
            this.consoleTextBox.ReadOnly = true;
            this.consoleTextBox.Size = new System.Drawing.Size(420, 99);
            this.consoleTextBox.TabIndex = 5;
            this.consoleTextBox.TabStop = false;
            this.consoleTextBox.Text = "";
            // 
            // connectButton
            // 
            this.connectButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.connectButton.Location = new System.Drawing.Point(357, 151);
            this.connectButton.Name = "connectButton";
            this.connectButton.Size = new System.Drawing.Size(75, 23);
            this.connectButton.TabIndex = 6;
            this.connectButton.Text = "Connect";
            this.connectButton.UseVisualStyleBackColor = true;
            // 
            // changeButton
            // 
            this.changeButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.changeButton.Location = new System.Drawing.Point(339, 71);
            this.changeButton.Name = "changeButton";
            this.changeButton.Size = new System.Drawing.Size(75, 23);
            this.changeButton.TabIndex = 21;
            this.changeButton.Text = "Change";
            this.changeButton.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(12, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 20);
            this.label1.TabIndex = 8;
            this.label1.Text = "Domain";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(12, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 38);
            this.label2.TabIndex = 9;
            this.label2.Text = "Username to log on as administrator";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(12, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 38);
            this.label3.TabIndex = 10;
            this.label3.Text = "Password to log on as administrator";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(6, 19);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 20);
            this.label4.TabIndex = 11;
            this.label4.Text = "Password";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(6, 46);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 20);
            this.label5.TabIndex = 12;
            this.label5.Text = "Password confirm";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lookupUsernameTextBox
            // 
            this.lookupUsernameTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lookupUsernameTextBox.Location = new System.Drawing.Point(155, 125);
            this.lookupUsernameTextBox.Name = "lookupUsernameTextBox";
            this.lookupUsernameTextBox.Size = new System.Drawing.Size(277, 20);
            this.lookupUsernameTextBox.TabIndex = 3;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(12, 116);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(100, 38);
            this.label6.TabIndex = 14;
            this.label6.Text = "Username to change a password";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.passwordChgTextBox);
            this.groupBox1.Controls.Add(this.passwordChgReTextBox);
            this.groupBox1.Controls.Add(this.changeButton);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Location = new System.Drawing.Point(12, 180);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(420, 104);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Chage user\'s password";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 401);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lookupUsernameTextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.connectButton);
            this.Controls.Add(this.consoleTextBox);
            this.Controls.Add(this.adminPasswordTextBox);
            this.Controls.Add(this.adminUsernameTextBox);
            this.Controls.Add(this.domainTextBox);
            this.MinimumSize = new System.Drawing.Size(460, 440);
            this.Name = "Form1";
            this.Text = "Change Active Directory User Password";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox domainTextBox;
        private System.Windows.Forms.TextBox adminUsernameTextBox;
        private System.Windows.Forms.TextBox adminPasswordTextBox;
        private System.Windows.Forms.TextBox passwordChgReTextBox;
        private System.Windows.Forms.TextBox passwordChgTextBox;
        private System.Windows.Forms.RichTextBox consoleTextBox;
        private System.Windows.Forms.Button connectButton;
        private System.Windows.Forms.Button changeButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox lookupUsernameTextBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox1;
        #endregion
    }
}
