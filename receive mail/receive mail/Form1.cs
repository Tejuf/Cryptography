﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace receive_mail
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataTable dt;
        private void Receive_Click(object sender, EventArgs e)
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
                dataGridView1.DataSource = dt;
                foreach (Outlook.MailItem item in inbox.Items)
                    dt.Rows.Add(new object[] { item.Subject, item.SenderName, item.HTMLBody, item.SentOn.ToLongDateString() + " " + item.SentOn.ToLongDateString() });
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex<dt.Rows.Count &&e.RowIndex>=0)
            {
                webBrowser1.DocumentText = dt.Rows[e.RowIndex]["Body"].ToString();
            }
        }
    }
}
