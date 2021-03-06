﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RDotNet;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Threading;

namespace PrimerDimerTool
{
    public partial class Form1 : Form
    {
        protected REngine engine = null;
        private string inputFileName;
        private string outputFileName;
        string installPath = null;
        private XmlDocument configDoc = new XmlDocument();  
        DataTable dt;
        public Form1()
        {
            InitializeComponent();
            string strFileName = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
            saveFileDialog1.InitialDirectory = "c:\\";
            saveFileDialog1.FileName = "dimer_result.xlsx";
            saveFileDialog1.Filter = "Excel文件(*.xlsx)|*.xlsx|所有文件(*.*)|*.*";
            configDoc.Load(strFileName);
            progressBar1.Visible = false;

            string installHome = Environment.GetEnvironmentVariable("PRIMERDIMERTOOL_HOME");
            if (installHome != null)
            {
                installPath = installHome;
            }
            else
            {
                installPath = getConfigSetting("installpath");
            }


            Environment.SetEnvironmentVariable("PATH", installPath + "/Primer3/");
            Environment.SetEnvironmentVariable("JAVA_HOME", installPath + "/Java");
            REngine.SetEnvironmentVariables(installPath + "/R/bin/x64", installPath + "/R");
            engine=REngine.GetInstance();
        }
        [DllImport(@"C:\Windows\System32\User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        private void button1_Click(object sender, EventArgs e)
        {

            if (openFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            inputFileName = openFileDialog1.FileName;
            dt = read_primer_sequence(inputFileName);
            textBox1.Text = inputFileName;
            textBox1.Update();
        }
        private DataTable read_primer_sequence(string filename)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("seqID");
            dt.Columns.Add("f_sequence");
            dt.Columns.Add("r_sequence");

            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Workbooks.Open(filename);
            //string tabName = excelApp.Workbooks[1].Worksheets[1].Name;
            excelApp.Workbooks[1].Worksheets[1].Activate();

            int i = 2;
            while (true)
            {
                DataRow dr = dt.NewRow();

                string id = Convert.ToString(excelApp.Cells[i, 1].Value);
                string f_sequence = excelApp.Cells[i, 2].Value;
                string r_sequence = excelApp.Cells[i, 3].Value;
                if (id == null || id.Length < 1 || id.Length < 1) break;
                dr[0] = id;
                dr[1] = f_sequence;
                dr[2] = r_sequence;
                dt.Rows.Add(dr);
                i++;
            }

            release_excel_app(ref excelApp);
            return dt;
        }

        private void release_excel_app(ref MSExcel.Application excelApp)
        {
            excelApp.Workbooks[1].Close();
            excelApp.Quit();
            if (excelApp != null)
            {
                int lpdwProcessID;
                GetWindowThreadProcessId(new IntPtr(excelApp.Hwnd), out lpdwProcessID);
                System.Diagnostics.Process.GetProcessById(lpdwProcessID).Kill();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            outputFileName = saveFileDialog1.FileName;
            textBox2.Text = outputFileName;
            textBox2.Update();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            progressBar1.Visible=true;
            string cwd = System.Environment.CurrentDirectory;
            string tmp_path = System.IO.Path.GetRandomFileName();

            tmp_path = cwd +"\\"+ tmp_path;
            if (Directory.Exists(tmp_path))
            {
                DirectoryInfo di = new DirectoryInfo(tmp_path);
                di.Delete(true);
            }else if(File.Exists(tmp_path)){
                return;
            }
            if (textBox2.Text == "")
            {
                return;
            }
            Directory.CreateDirectory(tmp_path);
            string primer3path = installPath + "/Primer3";
            string isDeleteTempDir = getConfigSetting("deleteTempDir");
            string nProcess = getConfigSetting("processNum");

            string[,] primerMat = new string[dt.Rows.Count,3];
            label3.Text = "preparing ...";
            label3.Update();
            engine.Evaluate("library(Biostrings)");
            progressBar1.Value = 5;
            engine.Evaluate("library(xlsx)");
            progressBar1.Value = 8;
            engine.Evaluate("source(\"primer_dimer_check.R\")");
            progressBar1.Value = 10;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string seqID = dt.Rows[i]["seqID"].ToString();
                string f_seq = dt.Rows[i]["f_sequence"].ToString();
                string r_seq = dt.Rows[i]["r_sequence"].ToString();
                primerMat[i, 0] = seqID;
                primerMat[i, 1] = f_seq;
                primerMat[i, 2] = r_seq;
            }
            CharacterMatrix primer = engine.CreateCharacterMatrix(primerMat);
            engine.SetSymbol("tmp_dir", engine.CreateCharacter(tmp_path));
            engine.SetSymbol("primer",primer);
            engine.SetSymbol("primer3dir", engine.CreateCharacter(primer3path));
            if (nProcess != null) {
                engine.SetSymbol("nprocess", engine.CreateInteger(Convert.ToInt32(nProcess)));
            }else{
                engine.SetSymbol("nprocess", engine.CreateInteger(4));
            }
            engine.SetSymbol("outputfile", engine.CreateCharacter(textBox2.Text));
            string[] bat_cmds = engine.Evaluate("prepare_bat(tmp_dir,primer,primer3dir,nprocess)").AsCharacter().ToArray();
            label3.Text = "dimer calculating ...";
            label3.Update();
            progressBar1.Value = 20;
            AutoResetEvent[] resets = new AutoResetEvent[bat_cmds.Length];

            for (int i = 0; i < bat_cmds.Length; i++)
            {
                resets[i] = new AutoResetEvent(false);
                ThreadTransfer transfer = new ThreadTransfer(bat_cmds[i],resets[i]);
                Thread thread = new Thread(new ParameterizedThreadStart(run_cmd));
                thread.Start(transfer);
            }
            foreach (var v in resets)
            {
                v.WaitOne();
                progressBar1.Value += 60 / resets.Length;
            }
            label3.Text = "result generating ...";
            label3.Update();
            progressBar1.Value = 80;
            engine.Evaluate("output_result(tmp_dir,primer,outputfile)");
            if ("true" == isDeleteTempDir)
            {
                DirectoryInfo di = new DirectoryInfo(tmp_path);
                di.Delete(true);
            }
            progressBar1.Value = 100;
            label3.Text = "";
            progressBar1.Visible = false;
            MessageBox.Show(this,"complete!");
        }
        public static string Execute(string dosCommand)
        {
            return Execute(dosCommand, 10);
        }
        public static string Execute(string command, int seconds)
        {
            string output = ""; 
            if (command != null && !command.Equals(""))
            {
                Process process = new Process();
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = "cmd.exe";
                startInfo.Arguments = "/C " + command;
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardInput = false;
                startInfo.RedirectStandardOutput = true; 
                startInfo.CreateNoWindow = true;
                process.StartInfo = startInfo;
                try
                {
                    if (process.Start())
                    {
                        if (seconds == 0)
                        {
                            process.WaitForExit();
                        }
                        else
                        {
                            process.WaitForExit(seconds); 
                        }
                        output = process.StandardOutput.ReadToEnd();
                    }
                }
                catch
                {
                }
                finally
                {
                    if (process != null)
                        process.Close();
                }
            }
            return output;
        }
        public string getConfigSetting(string strKey) {
            XmlNodeList nodes = configDoc.GetElementsByTagName("add");
            string value=null;
            for(int i = 0 ; i < nodes.Count;i++){
                XmlAttribute att = nodes[i].Attributes["key"];
                if (att.Value == strKey)
                {
                    att = nodes[i].Attributes["value"];
                    value = att.Value;
                    break;
                }
            }
            return value;
        }
        static void run_cmd(object obj)
        {
            ThreadTransfer transfer = (ThreadTransfer)obj;
            Execute(transfer.cmd);
            transfer.evt.Set();
        }

        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {
            button1_Click(sender, e);
        }
    }
    public class ThreadTransfer
    {
        public string cmd;
        public AutoResetEvent evt;
        public ThreadTransfer(string cmd , AutoResetEvent evt)
        {
            this.cmd = cmd;
            this.evt = evt;
        }
    }
}
