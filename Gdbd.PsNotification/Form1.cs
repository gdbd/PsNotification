namespace Gdbd.PsNotification
{
    using System;
    using System.Linq;
    using System.Management.Automation;
    using System.Management.Automation.Runspaces;
    using System.Windows.Forms;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.notifyIcon1.Text = "click to run script";
            this.notifyIcon1.MouseClick += (s, mc) => { if (mc.Button == MouseButtons.Left) { this.ShowInfo(); } };
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ShowInfo()
        {
            try
            {
                var runspaceConfiguration = RunspaceConfiguration.Create();

                using (Runspace runspace = RunspaceFactory.CreateRunspace(runspaceConfiguration))
                {
                    runspace.Open();

                    using (var scriptInvoker = new RunspaceInvoke(runspace))
                    {
                        scriptInvoker.Invoke("Set-ExecutionPolicy Unrestricted");
                        using (var pipeline = runspace.CreatePipeline())
                        {

                            Command myCommand = new Command(@".\Command.ps1");
                            pipeline.Commands.Add(myCommand);

                            var res = pipeline.Invoke();
                            this.notifyIcon1.BalloonTipText = res.Aggregate("", (c, n) => this.FormatOutput(c, n.BaseObject.ToString()));
                            if (string.IsNullOrEmpty(this.notifyIcon1.BalloonTipText))
                            {
                                this.notifyIcon1.BalloonTipText = "script returns no results";
                            }
                        }
                    }
                }             
            }
            catch(Exception ex)
            {
                this.notifyIcon1.BalloonTipText = ex.ToString();
            }

            this.notifyIcon1.ShowBalloonTip(500);
        }

        private string FormatOutput(string c, string n)
        {
            var current = (c != string.Empty ? c + Environment.NewLine : string.Empty);
            var next = (n.ToString().Length > 50 ? n.ToString().Substring(0, 50) + ".." : n.ToString());
            return current + next; ;
        }
    }
}
