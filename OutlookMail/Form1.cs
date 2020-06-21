using System;
using System.Data;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMail
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(() =>
            {
                try
                {


                    Outlook._Application _app = new Outlook.Application();


                    //attach = new Attach.Attachment(txbFile.Text);
                    //if (attach != null)
                    //{
                    //mail.Attachments.Add(attach, Outlook.OlAttachmentType.olByValue);

                    //}

                    string[] emails = txtTo.Text.Split(' ');
                    foreach (var item in emails)
                    {
                        var t = item.Trim();
                        Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                        mail.Subject = txtSubject.Text;
                        mail.Body = txtMessage.Text;
                        mail.Importance = Outlook.OlImportance.olImportanceNormal;
                        if (!string.IsNullOrWhiteSpace(txbFile.Text))
                        {
                            FileInfo file = new FileInfo(txbFile.Text);
                            mail.Attachments.Add(txbFile.Text);
                        }
                        if (t.LastIndexOf(',') == t.Length - 1)
                        {
                            var temp = item.Remove(item.Length - 1);
                            mail.To = temp;
                            ((Outlook._MailItem)mail).Send();
                            continue;
                        }
                        mail.To = t;
                        ((Outlook._MailItem)mail).Send();
                    }
                    MessageBox.Show("Your message has been successfully sent.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });
            thread.Start();
        }

        DataTable dt;
        private void btnReceive_Click(object sender, EventArgs e)
        {
            try
            {

                Outlook._Application _app = new Outlook.Application();
                Outlook._NameSpace _ns = _app.GetNamespace("MAPI");
                Outlook.MAPIFolder inbox = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                _ns.SendAndReceive(true);
                dt = new DataTable("Inbox");
                dt.Columns.Add("Subject", typeof(string));
                dt.Columns.Add("Sender", typeof(string));
                dt.Columns.Add("Body", typeof(string));
                dt.Columns.Add("Date", typeof(string));
                dataGrid.DataSource = dt;
                foreach (Outlook.MailItem item in inbox.Items)
                    dt.Rows.Add(new object[] { item.Subject, item.SenderName, item.HTMLBody, item.SentOn.ToLongDateString() + " " + item.SentOn.ToLongTimeString() });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < dt.Rows.Count && e.RowIndex >= 0)
                webBrowser.DocumentText = dt.Rows[e.RowIndex]["Body"].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StreamReader sr = new StreamReader(dialog.FileName);
                string emails = "";
                string email;
                //var emails = new List<string>(1000);
                while ((email = sr.ReadLine()) != null)
                {
                    emails += email + ", ";
                }
                //var str = "16520530@gm.uit.edu.vn, huypvn1998@gmail.com,";
                txtTo.Text = emails.Remove(emails.Length - 2);
                sr.Close();
            }
        }

        private void btnAttachFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txbFile.Text = dialog.FileName;
            }
        }
    }
    ////gchghvh
}
