using System;
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

namespace PrimerDimerTool
{
    public partial class Form1 : Form
    {
        protected REngine engine = REngine.GetInstance();
        private string inputFileName;
        private string outputFileName;
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

            //TODO 做成R的primer格式
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
            string cwd = System.Environment.CurrentDirectory;
            string tmp_path = System.IO.Path.GetRandomFileName();

            tmp_path = cwd +"\\"+ tmp_path;
            if (Directory.Exists(tmp_path) || File.Exists(tmp_path))
            {
                return;
            }
            Directory.CreateDirectory(tmp_path);
            string primer3path = getConfigSetting("primer3dir");
            string isDeleteTempDir = getConfigSetting("deleteTempDir");

            string[,] primerMat = new string[dt.Rows.Count,3];
            engine.Evaluate("library(Biostrings)");
            engine.Evaluate("library(xlsx)");
            engine.Evaluate("source(\"primer_dimer_check.R\")");
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
            engine.SetSymbol("outputfile", engine.CreateCharacter(textBox2.Text));
            //engine.Evaluate("check_dimer(primer,outputfile)");
            engine.Evaluate("prepare_bat(tmp_dir,primer,primer3dir)");
            Execute(tmp_path+ "/batch_run.bat");
            engine.Evaluate("output_result(tmp_dir,primer,outputfile)");
            if (isDeleteTempDir == "true")
            {
                DirectoryInfo di = new DirectoryInfo(tmp_path);
                di.Delete(true);
            }
            
            //TODO engine资源回收
        }
        public string Execute(string dosCommand)
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
    }
}
