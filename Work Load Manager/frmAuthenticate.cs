using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WorkLoad
{
    public partial class frmAuthenticate : Form
    {
        private string sessionID;

        public string SessionID
        {
            get { return sessionID; }
            //set { sessionID = value; }
        }

        public frmAuthenticate()
        {
            InitializeComponent();
        }

        private void frmAuthenticate_Load(object sender, EventArgs e)
        {
            webBrowser1.Navigate(RallyConstants.Rally_SSO_URL);
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            HtmlDocument doc = webBrowser1.Document;

            if (doc.Body.InnerText.IndexOf(RallyConstants.Rally_SessionID_Key) >-1)
            {
                string body = doc.Body.InnerText;
                sessionID = body.Substring(body.IndexOf("cookie value:")+ ("cookie value:").Length).Trim ();
                sessionID=sessionID.Replace("\n", "").Trim();
                sessionID = sessionID.Replace("\r", "").Trim();
                if (sessionID.Length > 0)
                {
                    this.Dispose();
                }
            }
        }
    }
}
