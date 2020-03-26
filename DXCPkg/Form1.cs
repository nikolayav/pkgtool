using Microsoft.Win32.TaskScheduler;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DXCPkg {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        StringBuilder sb = new StringBuilder();

        private void button1_Click(object sender, EventArgs e) {
            ofd.Filter = "CMD|*.cmd|BAT|*.bat";

            if (ofd.ShowDialog() == DialogResult.OK) {
                fileLocationTextBox.Text = ofd.FileName;
                if (!getLogFileLocation().Equals(string.Empty)) {
                    button3.Enabled = true;
                } else {
                    button3.Enabled = false;
                }
            }
        }

        private void Create_Click(object sender, EventArgs e) {
            createSchTask("DXC Install As SYSTEM", "/i");
        }

        private void button2_Click(object sender, EventArgs e) {
            createSchTask("DXC Uninstall As SYSTEM", "/u");
        }

        private void button3_Click(object sender, EventArgs e) {
            openLogFolder(getLogFileLocation());
        }


        private void createSchTask(string name, string arg) {
            if (string.IsNullOrWhiteSpace(fileLocationTextBox.Text)) {
                MessageBox.Show("Please, select a file.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else {
                Task tsk = null;
                string fileLocation = fileLocationTextBox.Text;

                using (TaskService ts = new TaskService()) {
                    TaskDefinition td = ts.NewTask();
                    td.Principal.RunLevel = TaskRunLevel.Highest;
                    td.RegistrationInfo.Description = "Runs a command-line file with " + arg + " argument.";
                    td.Settings.DisallowStartIfOnBatteries = false;
                    ExecAction ex = td.Actions.Add(new ExecAction(fileLocation, arg, null));

                    tsk = ts.RootFolder.RegisterTaskDefinition(name, td, TaskCreation.CreateOrUpdate, "SYSTEM", null, TaskLogonType.ServiceAccount);
                    tsk.Run();
                }

                taskRunActions(tsk);
            }
        }

        private void lockControls(bool isEnabled) {
            foreach (Control c in this.Controls) {
                if (c != progressBar1 && c != label2 && c != fileLocationTextBox) {
                    c.Enabled = isEnabled;
                }
            }
        }
        private void taskRunActions(Task tsk) {
            lockControls(false);
            if (tsk.Name.Equals("DXC Install As SYSTEM")) {
                sb.Append("Running Install task...");
                richTextBox1.Text = sb.ToString();
            } else if (tsk.Name.Equals("DXC Uninstall As SYSTEM")) {
                sb.Append("\nRunning Uninstall task...");
                richTextBox1.Text = sb.ToString();
            }

            label2.Text = "Running";
            progressBar1.MarqueeAnimationSpeed = 30;
            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.Value = 0;
            backgroundWorker1.RunWorkerAsync(tsk);
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e) {
            Task tsk = e.Argument as Task;

            while (tsk.State != TaskState.Ready) {
                Thread.Sleep(500);
            }

            e.Result = tsk;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {

            progressBar1.Value = 100;
            progressBar1.Style = ProgressBarStyle.Blocks;

            Task tsk = e.Result as Task;

            if (tsk.Name.Equals("DXC Install As SYSTEM")) {
                sb.Append("\nInstall task has finished.");
                richTextBox1.Text = sb.ToString();
            } else if (tsk.Name.Equals("DXC Uninstall As SYSTEM")) {
                sb.Append("\nUninstall task has finished.");
                richTextBox1.Text = sb.ToString();
            }
            if (e.Error != null) {
                MessageBox.Show(e.Error.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            label2.Text = "Ready";
            lockControls(true);

        }

        private string getLogFileLocation() {
            string logFilePath = string.Empty;
            string cmdFile = fileLocationTextBox.Text;
            
            if (File.Exists(cmdFile)) {
                StreamReader SR = new StreamReader(cmdFile);
                string line;
                while ((line = SR.ReadLine()) != null) {
                    if (line.Contains("APP_LogPath=")) {
                        logFilePath = line.Substring((line.IndexOf("=") + 1), line.Length - (line.IndexOf("=")) - 1);
                        break;
                    }
                }

                SR.Close();
                SR.Dispose();

            }
            return logFilePath;
        }

        private void openLogFolder(string logFilePath) {
            string absolutePath = Environment.ExpandEnvironmentVariables(logFilePath);
            if (Directory.Exists(absolutePath)) {
                ProcessStartInfo startInfo = new ProcessStartInfo {
                    Arguments = absolutePath,
                    FileName = "explorer.exe"
                };
            Process.Start(startInfo);
            } else {
                MessageBox.Show(string.Format("{0} Directory does not exist!", absolutePath));
            }
        }

 
    }
}
