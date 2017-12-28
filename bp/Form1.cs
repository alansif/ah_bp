using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Diagnostics;

namespace bp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            if (!System.IO.Directory.Exists("log"))
            {
                System.IO.Directory.CreateDirectory("log");
            }
            Trace.Listeners.Add(new System.Diagnostics.TextWriterTraceListener("log\\" + DateTime.Now.ToString("yyyyMMdd") + ".log"));
            Trace.Listeners.Add(new System.Diagnostics.TextWriterTraceListener(Console.Out));
        }

        private void OpenComm()
        {
            string PortName = System.Configuration.ConfigurationManager.AppSettings["PortName"];
            serialPort.PortName = PortName;
            serialPort.BaudRate = 2400;
            serialPort.DataBits = 7;
            serialPort.StopBits = System.IO.Ports.StopBits.One;
            serialPort.Parity = System.IO.Ports.Parity.Even;
            serialPort.Open();
        }

        public delegate void MyDelegate(Form1 f, string s);

        private void serialPort_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            try
            {
                string s = serialPort.ReadTo("\x03");
                MyDelegate md = new MyDelegate(showResult);
                this.BeginInvoke(md, this, s);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void showResult(Form1 f, string a)
        {
            Trace.WriteLine(a);
            Trace.Flush();
            f.textBox4.Text = a;
//            string a = @"SID99999999B17/11/22/14:50 125 076 067 E";
            int bph = Convert.ToInt32(a.Substring(27, 3));
            int bpl = Convert.ToInt32(a.Substring(31, 3));
            f.textBox1.Text = bph.ToString();
            f.textBox2.Text = bpl.ToString();
            SendToCHIS(bph.ToString(), bpl.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string Url = System.Configuration.ConfigurationManager.AppSettings["Url"];
            SHDocVw.InternetExplorer ie = new SHDocVw.InternetExplorer();
            ie.Navigate(Url);
            ie.Visible = true;
            OpenComm();
        }

        private static void SendToCHIS(string bp1, string bp2)
        {
            string Url = System.Configuration.ConfigurationManager.AppSettings["Url"];
            SHDocVw.InternetExplorer ie = getInternetExploer(Url);
            if (ie != null)
            {
                try
                {
                    MSHTML.HTMLDocument doc = ie.Document;
                    object oFrameIndex = 0;
                    var f = (doc.frames.item(ref oFrameIndex)).document;
                    f.getElementById("DefaultContent_Console_2001001_1").innerText = bp1;
                    f.getElementById("DefaultContent_Console_2001001_2").innerText = bp2;
                    //                traverseNodes(f.documentElement, ref cont);
                }
                catch (Exception)
                {
                    MessageBox.Show("请先将CHIS转到内科页面再测量血压");
                }
            }
        }

        private void traverseNodes(MSHTML.IHTMLElement parentNode, ref string cont)
        {
            if (parentNode.innerText != null)
                cont += parentNode.innerText;
            MSHTML.IHTMLElementCollection nodes = (MSHTML.IHTMLElementCollection)parentNode.children;
            IEnumerator ienum = nodes.GetEnumerator();
            while (ienum.MoveNext())
            {
                MSHTML.IHTMLElement node = (MSHTML.IHTMLElement)ienum.Current;
                traverseNodes(node, ref cont);
            }
        }

        public static SHDocVw.InternetExplorer getInternetExploer(string url)
        {
            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindows();
            foreach (object window in shellWindows)
            {
                SHDocVw.InternetExplorer ie = window as SHDocVw.InternetExplorer;
                if (ie != null &&
                        string.Equals(System.IO.Path.GetFileName(ie.FullName),
                        "iexplore.exe", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (ie.LocationURL.IndexOf(url) != -1)
                    {
                        return ie;
                    }
                }
            }
            return null;
        }
    }
}
