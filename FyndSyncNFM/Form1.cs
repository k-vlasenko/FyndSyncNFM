using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml.Linq;
using System.Diagnostics;
using System.Reflection;
using System.Net;
using System.Text.RegularExpressions;
using System.Globalization;
//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;

namespace FyndSyncNFM
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();       

            dataGridViewLog.Visible = false;
            dataGridViewLog[0, 0].Value = "   ";
            dataGridViewLog.Rows.Add(12);
            this.Text = "FindSync_v0.4.1.8";
            Height = 120;
            messageSync.Text = "Введите название и ноду, MARS-Done/ImpEx,FTP-UP/DOWN";
            messageRepLog.Text = "Введите название и ноду";
            //messageFTP.Text = "Введите название, ноду и up/down";
            errorMessageSync.Text = "";
            errorMessageRepLog.Text = "";
            //errorMessageFTP.Text = "";
            string input = "";
        }

        private int row = 0;
        private string input;

        private void ButtonSync_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory() + @"\out.txt";

            // Delete the file if it exists.
            if (System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
            }

            this.TextMessageSyncChange("Введите название и ноду, MARS-Done/ImpEx,FTP-UP/DOWN");

            string inp; //= Console.ReadLine();
            inp = textBoxSync.Text;
            char[] separators = { ' ', '_', '[', ']', '-', '\\', '/', '|' };
            string[] inpCheck = inp.Split(separators);

            //временный(надеюсь) костыль для аст
            if (inp.ToLower() == "ast 1")
            {
                OpenFolder(@"\\st-files\SYNCFOLDER\AST\CO");
                return;
            }
            if (inp.ToLower() == "ast 2")
            {
                OpenFolder(@"\\st-files\SYNCFOLDER\AST\Distributor1");
                return;
            }
            //и для черкизово сап
            if (inp.ToLower() == "cherkizovo 77")
            {
                OpenFolder(@"\\Cherkizovo-nsf\077_petelino_moscow_sap");
                return;
            }

            //костыль для марса
            if (inpCheck[0].ToUpper() == "MARS")
            {
                if (inpCheck.Length != 3)
                {
                    errorMessageSync.Text = "Неправильный формат ввода, выберите Done или ImpEx";
                    return;
                }
                else if (inpCheck[2].ToUpper() != "DONE" && inpCheck[2].ToUpper() != "IMPEX")
                {
                    errorMessageSync.Text = "Неправильный формат ввода, выберите Done или ImpEx";
                    return;
                }
            }
            else
            {
                if (inpCheck.Length != 2 && inpCheck.Length != 3)
                {
                    errorMessageSync.Text = "Неправильный формат ввода";
                    return;
                }
            }

            if (!Int32.TryParse(inpCheck[1], out int a))
            {
                errorMessageSync.Text = "Неправильный формат ввода";
                return;
            }

            //FTP folder
            //textBoxLog.AppendText(inpCheck.Length.ToString());
            if (inpCheck.Length == 3)
            {
                // textBoxLog.AppendText("kekega");
                if (inpCheck[2].ToUpper() == "UP" || inpCheck[2].ToUpper() == "DOWN" || inpCheck[2].ToUpper() == "FTP")
                {
                    //  textBoxLog.AppendText("kekega1");
                    if (FindFTPFolder(inp))
                    {
                        errorMessageSync.Text = "Путь найден";
                    }
                    else
                    {
                        errorMessageSync.Text = "Путь не найден";
                    }
                    return;
                }
            }

            string nodePath = FindFolderByName(inp);

            if (nodePath.Length > 0)
            {
                ErrorSyncChange(nodePath);
            }

            else
            {
                ErrorSyncChange("Ошибка! Путь не найден.");
            }

            TextMessageSyncChange("Введите название и ноду.");

            if (System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
            }

            return;
        }

        private static void OpenFolder(string path)
        {
            try
            {
                Process.Start(path);
            }
            catch (Win32Exception win32Exception)
            {
                //The system cannot find the file specified...
                //AddText(win32Exception.Message);
                Console.WriteLine(win32Exception);
            }
        }

        private static void OpenFile(string path)
        {
            try
            {
                Process.Start(path);
            }
            catch (Win32Exception win32Exception)
            {
                //The system cannot find the file specified...
                //AddText(win32Exception.Message);
                Console.WriteLine(win32Exception);
            }
        }

        private void ClearGrid()
        {
            int rows = dataGridViewLog.Rows.Count;
            int colums = dataGridViewLog.Columns.Count;
            for (int i = 0; i < colums; i++)
            {
                for (int j = 0; j < rows; j++)
                {
                    dataGridViewLog[i, j].Value = "";
                }
            }
            row = 0;
        }

        private void WriteToGrid(string s, int column)
        {
            int rows = dataGridViewLog.Rows.Count;
            if (row+1 == rows) dataGridViewLog.Rows.Add(2);

            dataGridViewLog[column, row].Value = s;
        }

        private void QuickAnalyzeSyncfolder(string path)
        {
            ClearGrid();
            dataGridViewLog.Visible = true;


            if (System.IO.File.Exists(path + @"\registers.xml"))
            {
                WriteToGrid("registers.xml", 0);
                textBoxLog.AppendText("registers ");
                AnalyzeXMLQuick(path + @"\registers.xml");
                row++;
            }
            else
            {
                if (System.IO.File.Exists(path + @"\rests.xml"))
                {
                    WriteToGrid("rests.xml", 0);
                    textBoxLog.AppendText("rests ");
                    AnalyzeXMLQuick(path + @"\rests.xml");
                    row++;
                }
                else
                {
                    WriteToGrid("rests.xml", 0);
                    WriteToGrid("doesn't exist", 1);
                    row++;
                    textBoxLog.AppendText("rests.xml doesn't exist\n");
                } 
                    

                if (System.IO.File.Exists(path + @"\debts.xml"))
                {
                    WriteToGrid("debts.xml", 0);
                    textBoxLog.AppendText("debts ");
                    AnalyzeXMLQuick(path + @"\debts.xml");
                    row++;
                }
                else
                {
                    WriteToGrid("debts.xml", 0);
                    WriteToGrid("doesn't exist", 1);
                    row++;
                    textBoxLog.AppendText("debts.xml doesn't exist\n");
                }
                    
            }
            if (System.IO.File.Exists(path + @"\references.xml"))
            {
                WriteToGrid("references.xml", 0);
                WriteToGrid(System.IO.File.GetLastWriteTime(path + @"\references.xml").ToString("dd.MM.yyyy"), 1);
                WriteToGrid(path + @"\references.xml", 3);
                row++;
                textBoxLog.AppendText("references.xml last change: " + System.IO.File.GetLastWriteTime(path + @"\references.xml").ToString("dd.MM.yyyy") + "\r\n\r\n");
            }
            else
            {
                WriteToGrid("references.xml", 0);
                WriteToGrid("doesn't exist", 1);
                row++;
                textBoxLog.AppendText("references.xml doesn't exist\r\n");
            }


            if (Directory.Exists(path + @"\client"))
            {
                var client = new DirectoryInfo(path + @"\client").GetFiles().OrderByDescending(x => x.LastWriteTime).ToList();
                int n = client.Count();
                if (n == 0)
                {
                    WriteToGrid("Client folder is empty", 0);
                    WriteToGrid("doesn't exist", 1);
                    row++;
                    textBoxLog.AppendText("Client folder is empty\r\n");
                    return;
                }



                int difDates = 0;
                var latestDate = client[0].LastWriteTime;
                for (int i = 0; i < n; i++)
                {
                    if (latestDate.Date == client[i].LastWriteTime.Date)
                    {
                        WriteToGrid(client[i].Name, 0);
                        textBoxLog.AppendText(client[i].Name + "\n");
                        AnalyzeXMLQuick(client[i].FullName);
                        row++;
                    }
                    //textBoxLog.AppendText(client[i].Name + " lastchange " + client[i].LastWriteTime.ToString("dd.MM.yyyy") + "\n");
                }

            }
            else
            {
                WriteToGrid("client dir", 0);
                WriteToGrid("doesn't exist", 1);
                row++;
            }
        }


        private string FindFolderByNameInDirectories(DirectoryInfo syncfolderDirectory, string distrib,
                                                              string node, int mode, string adParam)
        {
            char[] separators = { ' ', '_', '[', ']', '-', '\\', '/', '|' };

            foreach (DirectoryInfo dir in syncfolderDirectory.GetDirectories())
            {
                if (dir.Name.ToUpper() == distrib.ToUpper())
                {
                    DirectoryInfo nodeDir = dir;
                    if (mode == 1) nodeDir = new DirectoryInfo(dir.FullName + @"\r4000-archive");
                    if (distrib == "MARS")
                    {
                        if (adParam == "DONE") nodeDir = new DirectoryInfo(@"\\backup2.mtproject.ru\TENANT_TEAMS\MARS\HoustonArchive\Production\cicerone-archive");
                        else nodeDir = new DirectoryInfo(@"\\ST-FiLES.mtproject.ru\PLATFORM\MARS\SyncFolder");
                    }
                    foreach (DirectoryInfo d in nodeDir.GetDirectories())
                    {
                        string[] subs = d.Name.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        //AddText(subs[0]);
                        bool nodeCheck = false;
                        int val=0;
                        nodeCheck = Int32.TryParse(subs[0], out val);
                        if (!nodeCheck && subs.Length > 1) nodeCheck = Int32.TryParse(subs[1], out val);
                        
                        if (nodeCheck)
                        {
                            if (val == Convert.ToInt32(node))
                            {
                                OpenFolder(d.FullName);

                                if (QuickAnalyzeBox.Checked)
                                {
                                    QuickAnalyzeBox.Checked = false;
                                    QuickAnalyzeSyncfolder(d.FullName);
                                    TabContr.SelectTab(Log);
                                }

                                return d.FullName;
                            }
                        }

                    }
                    break;
                }
            }
            return "";
        }

        private string FindFolderByName(string name)
        {
            char[] separators = { ' ', '_', '[', ']', '-' };
            string[] splitName = name.Split(separators);
            string distrib = splitName[0];
            string node = splitName[1];
            string adParam = "";

            if (splitName.Length == 3) adParam = splitName[2].ToUpper();

            distrib = distrib.ToUpper();
            node = node.ToUpper();


            var aliases = AliasesInit();
            distrib = AliasesCheck(distrib, aliases);

            DirectoryInfo syncfolderDirectory, houstonDirectory;
            if (Directory.Exists(@"\\ST-FiLES.mtproject.ru\SYNCFOLDER"))
            {
                syncfolderDirectory = new DirectoryInfo(@"\\ST-FiLES.mtproject.ru\SYNCFOLDER");
            }
            else
            {
                return "";
            }

            if (Directory.Exists(@"\\ST-FiLES.mtproject.ru\HOUSTON"))
            {
                houstonDirectory = new DirectoryInfo(@"\\ST-FiLES.mtproject.ru\HOUSTON");
            }
            else
            {
                return "";
            }

            string ans = FindFolderByNameInDirectories(syncfolderDirectory, distrib, node, 0, adParam);
            if (ans != "") return ans;

            ans = FindFolderByNameInDirectories(houstonDirectory, distrib, node, 1, adParam);
            return ans;
        }

        private static async Task AddTextAsync(string value) //Запись в файл out.txt в диркетории программы
        {
            string path = Directory.GetCurrentDirectory() + @"\tmplog.txt";

            if (System.IO.File.Exists(path))
            {
                using (System.IO.FileStream fs = System.IO.File.Open(path, FileMode.Append))
                {
                    byte[] info = new UTF8Encoding(true).GetBytes(value + "\n");
                    await fs.WriteAsync(info, 0, info.Length);
                }
            }
            else
            {
                using (FileStream fs = System.IO.File.Create(path))
                {
                    byte[] info = new UTF8Encoding(true).GetBytes(value);
                    await fs.WriteAsync(info, 0, info.Length);
                }
            }
        }

        private void AddText(string[] value)
        {
            string path = Directory.GetCurrentDirectory() + @"\tmplog.txt";
            System.IO.File.WriteAllLines(path, value);
        }

        public void TextMessageSyncChange(string s)
        {
            this.messageSync.Text = s;
        }

        public void ErrorSyncChange(string s)
        {
            this.errorMessageSync.Text = s;
        }

        private void TextBoxSync_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r') //ENTER key
            {
                ButtonSync_Click(this, new EventArgs());
            }
            if (e.KeyChar == (char)27) //ESC key
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        private static string AliasesCheck(string name, Tuple<string, string>[] aliases)
        {
            foreach (Tuple<string, string> a in aliases)
            {
                if (a.Item2.ToUpper() == name.ToUpper())
                {
                    return a.Item1;
                }
            }
            return name;
        }

        private static Tuple<string, string>[] AliasesInit()
        {
            return new Tuple<string, string>[]
            {
                new Tuple<string, string> ("ALADUSHKIN", "PMK"),

                new Tuple<string, string> ("CCH", "COLA"),
                new Tuple<string, string> ("CCH", "CocaCola"),

                new Tuple<string, string> ("ROUST", "RA"),

                new Tuple<string, string> ("BP", "BritishPetroleum"),

                new Tuple<string, string> ("MDLZ", "Mondelez"),


            };
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Tab))
            {
                TabContr.SelectedIndex = (TabContr.SelectedIndex + 1 < TabContr.TabCount) ?
                                         TabContr.SelectedIndex + 1 : 0;
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void TextBoxSync_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) //ENTER key
            {
                ButtonSync_Click(this, new EventArgs());
            }
            if (e.KeyCode == Keys.Escape) //ESC key
            {
                System.Windows.Forms.Application.Exit();
            }

        }

        private void TextBoxRepLog_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                RepLogButton_Click(this, new EventArgs());
            }
            if (e.KeyCode == Keys.Escape)
            {
                System.Windows.Forms.Application.Exit();
            }

        }

        /*
        private void TextBoxFTP_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ButtonFTP_Click(this, new EventArgs());
            }
            if (e.KeyCode == Keys.Escape)
            {
                Application.Exit();
            }

        }

        private void TextBoxXML_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ButtonFTP_Click(this, new EventArgs());
            }
            if (e.KeyCode == Keys.Escape)
            {
                Application.Exit();
            }
        }
        */

        public string SchedulerForDistr(string path, string nodeid)
        {
            //textBoxLog.AppendText("fordistr" + path + "\r\n");
            if (!System.IO.File.Exists(path)) return String.Empty;
            string log = string.Empty;
            char[] separators = { ' ', '_', '[', ']', '-', '\\', '/', '|', '=' };
            using (StreamReader fs = new StreamReader(path, Encoding.GetEncoding(1251)))
            {
                string s;
                while ((s = fs.ReadLine()) != null)
                {
                    string[] sp = s.Split(separators);
                    string prev = string.Empty;
                    foreach (string x in sp)
                    {
                        if (x == nodeid && (prev == "Дистрибьютор" || prev == "nodeId" || prev == "contextNodeID"))
                        {
                            log += s + "\r\n";

                            break;
                        }
                        prev = x;
                    }
                }
            }
            return log;
        }

        private bool AnalyzeLogDir(string path, int node, int add)
        {
            //textBoxLog.AppendText("logdir\r\n");
            DirectoryInfo logDir = new DirectoryInfo(path);
            char[] separators = { ' ', '_', '[', ']', '-' };
            string[] log = new string[add + 1];
            foreach (FileInfo file in logDir.GetFiles())
            {
                string[] s = file.Name.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                if (Int32.TryParse(s[1], out int i))
                {
                    if (i == node)
                    {
                        if (checkBoxFolder.Checked) OpenFolder(path);
                        if (checkBoxLog.Checked) OpenFile(file.FullName);

                        if (CheckBoxScheduler.Checked)
                        {
                            CheckBoxScheduler.Checked = false;
                            if (add > 0)
                            {
                                for (int j = add; j >= 1; j--)
                                {
                                    //textBoxLog.AppendText("she" + j);
                                    log[add - j] = SchedulerForDistr(path + @"\She_rpl_Core.utf8.dll.log." + j.ToString(), node.ToString());

                                }
                            }
                            //textBoxLog.AppendText(add.ToString() + "\r\n");
                            log[add] = SchedulerForDistr(path + @"\She_rpl_Core.utf8.dll.log", node.ToString());

                            AddText(log);
                            OpenFile(Directory.GetCurrentDirectory() + @"\tmplog.txt");
                        }
                        return true;
                    }
                }
            }
            return false;
        }

        private bool FindReplicationLogFolder(string inp)
        {
            char[] separators = { ' ', '_', '[', ']', '-' };
            string[] splitName = inp.Split(separators);
            string distrib = splitName[0];
            string node = splitName[1];
            string add = "0";
            if (splitName.Length == 3) add = splitName[2];
            distrib = distrib.ToUpper();
            node = node.ToUpper();

            var aliases = AliasesInit();
            distrib = AliasesCheck(distrib, aliases);

            DirectoryInfo replicationDir;
            if (Directory.Exists(@"\\ST-FiLES.mtproject.ru\REPLICATION"))
            {
                replicationDir = new DirectoryInfo(@"\\ST-FiLES.mtproject.ru\REPLICATION");
            }
            else
            {
                errorMessageRepLog.Text = "Ошибка! Путь не найден.";
                return false;
            }

            bool f = false;

            foreach (DirectoryInfo dir in replicationDir.GetDirectories())
            {
                if (dir.Name.ToUpper() == distrib.ToUpper())
                {

                    if (Directory.Exists(dir.FullName + @"\Replication.FTP"))
                    {
                        if (Directory.Exists(dir.FullName + @"\Replication.FTP\Logs\Production"))
                        {
                            f = f || AnalyzeLogDir(dir.FullName + @"\Replication.FTP\Logs\Production", Int32.Parse(node), Int32.Parse(add));
                        }
                        else
                        {
                            f = f || AnalyzeLogDir(dir.FullName + @"\Replication.FTP\Logs", Int32.Parse(node), Int32.Parse(add));
                        }
                    }

                    if (Directory.Exists(dir.FullName + @"\Replication.Shuttle"))
                    {
                        if (Directory.Exists(dir.FullName + @"\Replication.Shuttle\Logs\Production"))
                        {
                            f = f || AnalyzeLogDir(dir.FullName + @"\Replication.Shuttle\Logs\Production", Int32.Parse(node), Int32.Parse(add));
                            if (!f)
                            {
                                f = f || AnalyzeLogDir(dir.FullName + @"\Replication.Shuttle\Logs", Int32.Parse(node), Int32.Parse(add));
                            }
                        }
                        else
                        {
                            f = f || AnalyzeLogDir(dir.FullName + @"\Replication.Shuttle\Logs", Int32.Parse(node), Int32.Parse(add));
                        }
                    }
                    if (Directory.Exists(dir.FullName + @"\Replication.Export"))
                    {
                        f = f || AnalyzeLogDir(dir.FullName + @"\Replication.Export\Logs", Int32.Parse(node), Int32.Parse(add));
                    }
                }
            }

            return f;
        }

        private DirectoryInfo FindDirectory(DirectoryInfo current, string name)
        {
            DirectoryInfo ans = null;
            //textBoxLog.AppendText(current.FullName + "\r\n");
            if (current.Name.ToUpper() == name.ToUpper())
            {
                return current;
            }
            else
            {
                //DirectoryInfo[] child = current.GetDirectories()
                foreach (DirectoryInfo d in current.GetDirectories())
                {
                    current = FindDirectory(d, name);

                    ans = (current != null ? current : ans);
                }
            }
            return ans;
        }


        private bool FindFTPFolder(string inp)
        {
            char[] separators = { ' ', '_', '[', ']', '-' };
            string[] splitName = inp.Split(separators);

            string distrib = splitName[0].ToUpper();
            int node = Int32.Parse(splitName[1]);
            string mod = splitName[2].ToUpper();
            //textBoxLog.AppendText(splitName[0] + " " + splitName[1] + "\n");

            var aliases = AliasesInit();
            distrib = AliasesCheck(distrib, aliases);

            DirectoryInfo FTPDir;

            if(mod.ToUpper() == "FTP")
            {
                if (Directory.Exists(@"\\ST-FiLES.mtproject.ru\FTP\FTP_HOUSTON\MT"))
                {
                    FTPDir = new DirectoryInfo(@"\\ST-FiLES.mtproject.ru\FTP\FTP_HOUSTON\MT");
                    foreach(DirectoryInfo d in FTPDir.GetDirectories())
                    {
                        string[] dirSplitName = d.Name.Split(separators);
                        
                        if(dirSplitName.Length >= 3)
                        {
                            //textBoxLog.AppendText(dirSplitName[0] + " " + dirSplitName[2] + "\n");
                            if (dirSplitName[0].ToUpper() == distrib.ToUpper() && dirSplitName[2] == splitName[1])
                            {
                                OpenFolder(d.FullName);
                                return true;
                            }
                        }
                    }

                }
                else
                {
                    errorMessageSync.Text = "Ошибка! Каталог FTP не найден.";
                    return false;
                }
            }

            if (Directory.Exists(@"\\" + distrib + @"-nsf.mtproject.ru\FTP"))
            {
                FTPDir = new DirectoryInfo(@"\\" + distrib + @"-nsf.mtproject.ru\FTP");
            }
            else
            {
                errorMessageSync.Text = "Ошибка! Каталог FTP не найден.";
                return false;
            }

            DirectoryInfo tmpDir = FindDirectory(FTPDir, mod);
            if (tmpDir == null)
            {
                tmpDir = FindDirectory(FTPDir, "y"+mod);
                if (tmpDir == null) return false;
            }
            FTPDir = tmpDir;
            //textBoxLog.AppendText(FTPDir.FullName);
            bool f = false;

            foreach (DirectoryInfo d in FTPDir.GetDirectories())
            {
                if (Int32.Parse(d.Name) == node)
                {
                    OpenFolder(d.FullName);
                    f = true;
                }
            }

            return f;
        }

        
        private void AnalyzeXLS(string path)
        {

        }

        private async Task AnalyzeXMLQuick(string path)
        {
            WriteToGrid(path, 3);
            TabContr.SelectTab(Log);
            DateTime[] dateTimes = new DateTime[10000];
            int cnt = 0;
            DateTime maxDate = new DateTime(1900, 1, 1);

            XDocument doc = new XDocument();
            try
            {

                doc = XDocument.Load(path);
            }
            catch
            {
                WriteToGrid("Could not load ", 1);
                textBoxLog.AppendText("Could not load " + path + "\r\n");
                //textBoxLog.Text += "Could not load " + x.Name + "\r\n";
                //Console.WriteLine("Could not load " + path);
                return;
            }


            foreach (XElement node in doc.Root.Descendants())
            {
                //textBoxLog.AppendText(node.FirstAttribute.ToString());
                foreach(XAttribute at in node.Attributes())
                {
                    //textBoxLog.AppendText(at.Name.ToString());
                    if (at.Name.ToString().ToLower() == "dateto" || at.Name.ToString().ToLower() == "date")
                    {
                        
                        DateTime.TryParse(at.Value, out DateTime d);
                        bool f = false;
                        foreach (DateTime n in dateTimes)
                        {
                            if (n.Date == d.Date)
                            {
                                f = true;
                                break;
                            }
                        }
                        if (!f)
                        {
                            dateTimes[cnt] = d;
                            cnt++;
                        }
                        if (DateTime.Compare(maxDate, d) <= 0)
                        {
                            maxDate = d;
                        }
                    }
                }

                //AddText(node.Name.LocalName);
                if (node.Name.LocalName.ToLower() == "date" || node.Name.LocalName.ToLower() == "period")
                {
                    DateTime.TryParse(node.Value, out DateTime d);
                    bool f = false;
                    foreach (DateTime n in dateTimes)
                    {
                        if (n.Date == d.Date)
                        {
                            f = true;
                            break;
                        }
                    }
                    if (!f)
                    {
                        dateTimes[cnt] = d;
                        cnt++;
                    }
                    //AddText(d.ToString("dd.MM.yyyy") + "\n");
                    //textBoxLog.Text += d.ToString("dd.MM.yyyy") + "\r\n";
                    //textBoxLog.L
                    if (DateTime.Compare(maxDate, d) <= 0)
                    {
                        maxDate = d;
                    }
                }
            }

            if (cnt == 0)
            {
                WriteToGrid(new FileInfo(path).LastWriteTime.Date.ToString("dd.MM.yyyy"), 1);
                WriteToGrid("Dates not found", 2);
                textBoxLog.AppendText("Dates not found\r\n");
                return;
            }

            dateTimes.OrderByDescending(x => x.Date);

            List<DateTime> dateList = new List<DateTime>();


            for (int i = 0; i < cnt; i++)
            {
                dateList.Add(dateTimes[i]);
            }

            dateList = dateList.OrderBy(x => x.Date).ToList();

            WriteToGrid(new FileInfo(path).LastWriteTime.Date.ToString("dd.MM.yyyy"), 1);
            WriteToGrid(dateList.First().ToString("dd.MM.yyyy") + " - " + dateList.Last().ToString("dd.MM.yyyy"), 2);
            textBoxLog.AppendText(dateList.First().ToString("dd.MM.yyyy") + " - " + dateList.Last().ToString("dd.MM.yyyy") + "\r\n");
            //errorMessageXML.Text = dateList.First().ToString("dd.MM.yyyy") + " - " + dateList.Last().ToString("dd.MM.yyyy");
        }

        private async Task AnalyzeXML(string path)
        {
            //dataGridViewLog.Visible = false;
            WriteToGrid(path, 3);
            TabContr.SelectTab(Log);
            DateTime[] dateTimes = new DateTime[10000];
            int cnt = 0;
            DateTime maxDate = new DateTime(1900, 1, 1);
            FileInfo[] files;
            if (System.IO.File.Exists(path))
            {
                files = new FileInfo[] { new FileInfo(path) };
            }
            else
            {
                files = new DirectoryInfo(path).GetFiles();
            }

            //textBoxLog.AppendText(DateTime.TryParse("20210505",out DateTime ttt).ToString());

            foreach (FileInfo x in files)
            {
                XDocument doc = new XDocument();
                try
                {

                    doc = XDocument.Load(x.FullName);
                }
                catch
                {
                    textBoxLog.AppendText("Could not load " + x.Name + "\r\n");
                    //textBoxLog.Text += "Could not load " + x.Name + "\r\n";
                    Console.WriteLine("Could not load " + x.Name);
                    continue;
                }


                foreach (XElement node in doc.Root.Descendants())
                {
                    //textBoxLog.AppendText(node.FirstAttribute.ToString());
                    foreach (XAttribute at in node.Attributes())
                    {
                        //textBoxLog.AppendText(at.Name.ToString());
                        if (at.Name.ToString().ToLower() == "dateto" || at.Name.ToString().ToLower() == "date" || at.Name.ToString().ToLower() == "dtlm")
                        {

                            if (DateTime.TryParse(at.Value, out DateTime d)) ;
                            else DateTime.TryParseExact(at.Value, "yyyyMMdd hh:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out d);

                                bool f = false;
                            foreach (DateTime n in dateTimes)
                            {
                                if (n.Date == d.Date)
                                {
                                    f = true;
                                    break;
                                }
                            }
                            if (!f)
                            {
                                dateTimes[cnt] = d;
                                cnt++;
                            }
                            if (DateTime.Compare(maxDate, d) <= 0)
                            {
                                maxDate = d;
                            }
                            continue;
                        }
                    }

                    //AddText(node.Name.LocalName);
                    if (node.Name.LocalName.ToLower() == "date" || node.Name.LocalName.ToLower() == "period")
                    {
                        DateTime.TryParse(node.Value, out DateTime d);
                        bool f = false;
                        foreach (DateTime n in dateTimes)
                        {
                            if (n.Date == d.Date)
                            {
                                f = true;
                                break;
                            }
                        }
                        if (!f)
                        {
                            dateTimes[cnt] = d;
                            cnt++;
                        }
                        //AddText(d.ToString("dd.MM.yyyy") + "\n");
                        //textBoxLog.Text += d.ToString("dd.MM.yyyy") + "\r\n";
                        //textBoxLog.L
                        if (DateTime.Compare(maxDate, d) <= 0)
                        {
                            maxDate = d;
                        }
                    }
                }
                WriteToGrid(x.Name, 0);
                WriteToGrid(x.LastWriteTime.Date.ToString("dd.MM.yyyy"), 1);
                textBoxLog.AppendText("Done " + x.Name + "\r\n");
            }
            if (cnt == 0)
            {
                textBoxLog.AppendText("Dates not found\r\n");
                return;
            }

            dateTimes.OrderByDescending(x => x.Date);

            List<DateTime> dateList = new List<DateTime>();


            for (int i = 0; i < cnt; i++)
            {
                dateList.Add(dateTimes[i]);
            }

            dateList = dateList.OrderBy(x => x.Date).ToList();
            foreach (DateTime d in dateList)
            {
                textBoxLog.AppendText(d.ToString("dd.MM.yyyy") + "\r\n");
                //textBoxLog.Text += d.ToString("dd.MM.yyyy") + "\r\n";
            }
            WriteToGrid(dateList.First().ToString("dd.MM.yyyy") + " - " + dateList.Last().ToString("dd.MM.yyyy"), 2);
            row++;
            textBoxLog.AppendText("Period: " + dateList.First().ToString("dd.MM.yyyy") + " - " + dateList.Last().ToString("dd.MM.yyyy") + "\r\n\r\n");
            //errorMessageXML.Text = dateList.First().ToString("dd.MM.yyyy") + " - " + dateList.Last().ToString("dd.MM.yyyy");

            textBoxLog.SelectionStart = textBoxLog.Text.Length;
            textBoxLog.ScrollToCaret();
        }


        private void RepLogButton_Click(object sender, EventArgs e)
        {
            char[] separators = { ' ', '_', '[', ']', '-', '\\', '/', '|' };
            messageRepLog.Text = "Введите название и ноду:";

            string inp = textBoxRepLog.Text;
            string[] inpCheck = inp.Split(separators);

            if (inpCheck.Length == 3)
            {
                if (!Int32.TryParse(inpCheck[2], out int b))
                {
                    errorMessageRepLog.Text = "Неправильный формат ввода";
                    return;
                }
            }
            else
            {
                if (inpCheck.Length != 2)
                {
                    errorMessageRepLog.Text = "Неправильный формат ввода";
                    return;
                }
            }

            if (!Int32.TryParse(inpCheck[1], out int a))
            {
                errorMessageRepLog.Text = "Неправильный формат ввода";
                return;
            }

            if (FindReplicationLogFolder(inp))
            {
                errorMessageRepLog.Text = "Путь найден";
            }
            else
            {
                errorMessageRepLog.Text = "Ошибка! Путь не найден.";
            }
            messageRepLog.Text = "Введите название и ноду:";
        }



        private void ReplLogTab_Enter(object sender, EventArgs e)
        {
            textBoxRepLog.Text = input;
            textBoxRepLog.Select();
        }

        private void SyncfolderTab_Enter(object sender, EventArgs e)
        {
            textBoxSync.Text = input;
            textBoxSync.Select();
        }

        /*
        private void FTP_Enter(object sender, EventArgs e)
        {
            textBoxFTP.Text = input;
            textBoxFTP.Select();
        }

        private void XmlAnalyzer_Enter(object sender, EventArgs e)
        {
            textBoxFTP.Select();
        }

        private void FTP_Leave(object sender, EventArgs e)
        {
            input = textBoxFTP.Text;
        }
        */

        private void ReplLogTab_Leave(object sender, EventArgs e)
        {
            input = textBoxRepLog.Text;
        }

        private void SyncfolderTab_Leave(object sender, EventArgs e)
        {
            input = textBoxSync.Text;
        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        /*
        private void ButtonFTP_Click(object sender, EventArgs e)
        {
            char[] separators = { ' ', '_', '[', ']', '-', '\\', '/', '|' };

            string inp = textBoxFTP.Text;
            string[] inpCheck = inp.Split(separators);


            if (inpCheck.Length != 3)
            {
                errorMessageFTP.Text = "Неправильный формат ввода";
                return;
            }
            if (!Int32.TryParse(inpCheck[1], out int a) || (inpCheck[2].ToUpper() != "UP" && inpCheck[2].ToUpper() != "DOWN"))
            {
                errorMessageFTP.Text = "Неправильный формат ввода, введите up или down";
                return;
            }

            if (FindFTPFolder(inp))
            {
                errorMessageFTP.Text = "Путь найден";
            }
            else
            {
                errorMessageFTP.Text = "Путь не найден";
            }
        }
        */




        /*private async void ButtonXML_Click(object sender, EventArgs e)
        {
            string path = textBoxXML.Text;
            if (System.IO.File.Exists(path))
            {
                await AnalyzeXML(path);
            }
            else
            {
                if (Directory.Exists(path))
                {
                    await AnalyzeXML(path);
                }
                else
                {
                    //errorMessageXML.Text = ("File/Directory does not exist");
                }
            }
        }
        */

        private void TabContr_Enter(object sender, EventArgs e)
        {

        }

        /*
        private async void ButtonOpenXmlFile_Click(object sender, EventArgs e)
        {
            string[] filePath = { };

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                if (Directory.Exists(@"\\ST-FiLES.mtproject.ru"))
                {
                    openFileDialog.InitialDirectory = @"\\ST-FiLES.mtproject.ru";
                }

                openFileDialog.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileNames;
                }
            }
            foreach (string s in filePath)
            {
                if (System.IO.File.Exists(s))
                {
                    await AnalyzeXML(s);
                }
                else
                {
                    if (Directory.Exists(s))
                    {
                        await AnalyzeXML(s);
                    }
                    else
                    {
                        errorMessageXML.Text = ("File/Directory does not exist");
                    }
                }


            }
        }
        */

        private void Log_Enter(object sender, EventArgs e)
        {
            Height = 390;
            //textBoxLog.ScrollToCaret();
        }

        private void Log_Leave(object sender, EventArgs e)
        {
            Height = 120;
        }

        private void ButtonLogClear_Click(object sender, EventArgs e)
        {
            textBoxLog.Text = string.Empty;
            ClearGrid();
            //TabContr.SelectTab(XmlAnalyzer);
            
        }


        private void AnalyzeXLSX(string path)
        {
           
        }

        private async void DragAndDrop(string[] files)
        {
            TabContr.SelectTab(Log);
            foreach (string s in files)
            {
                var regex = new Regex(@"\.[0-9]");
                string ext = Path.GetExtension(s).ToLower();
                //|| regex.IsMatch(Path.GetExtension(s).ToLower())
                textBoxLog.AppendText(Path.GetFileName(s) + "\r\n");

                /*
                if (Path.GetExtension(s).ToLower() == ".xlsx")
                {
                    AnalyzeXLSX(s);
                }
                */

                    if (Path.GetExtension(s).ToLower() == ".xml" || Directory.Exists(s))
                {
                    //textBoxXML.Text = s;
                    //ButtonXML_Click(null, null);

                    if (System.IO.File.Exists(s))
                    {
                        await AnalyzeXML(s);
                    }
                    else
                    {
                        if (Directory.Exists(s))
                        {
                            await AnalyzeXML(s);
                        }
                        else
                        {
                            textBoxLog.AppendText("File/Directory does not exist");
                            //errorMessageXML.Text = ("File/Directory does not exist");
                        }
                    }
                }
                else if (ext == ".txt" || ext == ".log" || regex.IsMatch(ext))
                {
                    //textBoxLog.AppendText(Path.GetExtension(s).ToLower());
                    int node;
                    if (textBoxRepLog.Text.Split().Length == 2)
                    {
                        Int32.TryParse(textBoxRepLog.Text.Split()[1], out node);
                    }
                    else
                    {
                        textBoxLog.AppendText("Введите ноду в RepLog\r\n\r\n");
                        continue;
                    }

                    CheckBoxScheduler.Checked = true;
                    string[] log = new string[1];
                    log[0] = SchedulerForDistr(s, node.ToString());
                    AddText(log);
                    OpenFile(Directory.GetCurrentDirectory() + @"\tmplog.txt");
                }
                else
                {
                    textBoxLog.AppendText("Некорректный файл\r\n");
                }
                textBoxLog.AppendText("\r\n");
            }
        }


        private void TextBoxXML_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            DragAndDrop(files);

        }
        private void TextBoxXML_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                e.Effect = DragDropEffects.All;
            }
        }


        private void TabContr_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            DragAndDrop(files);
        }

        private void TabContr_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                e.Effect = DragDropEffects.All;
            }
        }



        private void TextBoxLog_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            DragAndDrop(files);

        }

        private void TextBoxLog_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                e.Effect = DragDropEffects.All;
            }
        }

        private void buttonLogSwitch_Click(object sender, EventArgs e)
        {
            dataGridViewLog.Visible = !dataGridViewLog.Visible;
        }

        private void ButtonUpdate_Click(object sender, EventArgs e)
        {
            string current = Assembly.GetEntryAssembly().Location;
            string tmp = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\" + Assembly.GetEntryAssembly().GetName().Name + ".bak";
            string webResource = @"https://www.dropbox.com/s/olcepppwgl41acm/FindSync.exe?dl=1";
            //if (File.Exists(s + ".bak")) File.Delete(s + ".bak");
            textBoxLog.AppendText(current + "\r\n" + tmp);
            if (System.IO.File.Exists(tmp)) System.IO.File.Delete(tmp);
            System.IO.File.Move(current, tmp);
            WebClient myWebClient = new WebClient();
            try
            {
                myWebClient.DownloadFile(webResource, "fs.exe");
            }
            catch
            {
                if (System.IO.File.Exists(Directory.GetCurrentDirectory() + @"\fs.exe"))
                {
                    System.IO.File.Delete(Directory.GetCurrentDirectory() + @"\fs.exe");
                }
                System.IO.File.Move(tmp, current);
                return;
            }

            System.IO.File.Move(Directory.GetCurrentDirectory() + @"\fs.exe", current);

            System.Windows.Forms.Application.Restart();

        }

        private void dataGridViewLog_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //dataGridViewLog[e.ColumnIndex, e.RowIndex].Value = "kek";
            if(dataGridViewLog[3, e.RowIndex].Value != null)
            {
                string path = dataGridViewLog[3, e.RowIndex].Value.ToString();
                if (System.IO.File.Exists(path))
                {
                    OpenFile(path);
                }
            }
            
        }
    }
}