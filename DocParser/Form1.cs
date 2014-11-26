using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
namespace WordToExcel
{
    public partial class Form1 : Form
    {
        Thread thd;
        int currentCount, totalCount;
        // СОЦ НАЙМ -> АКТ -> ОСТАЛЬНОЕ
        public Form1()
        {
            InitializeComponent();
        }

        void ParseFiles()
        {
            List<string> errorFiles = new List<string>();
            var dir = "D:\\DocsToParse\\Output";
            var allDirs = Directory.GetFileSystemEntries(dir);
            var outputDir = dir;//Path.Combine(dir, "Output");
            //if (!Directory.Exists(outputDir))                        // common processing
            //    Directory.CreateDirectory(outputDir);
            totalCount = allDirs.Count();
            currentCount = 0;
            foreach (var entry in allDirs)
            {
                if (entry.EndsWith("Output") || entry.EndsWith("Processed") || !entry.EndsWith(".doc") || Path.GetFileName(entry).StartsWith("~$"))
                    continue;

                Logger.LogI(string.Format("--- Processing directory: {0} ---", entry));


                if (CheckFile(entry))
                    errorFiles.Add(entry);

                //var filesInDirList = Directory.GetFiles(entry).ToList();
                //filesInDirList.RemoveAll(f => Path.GetFileName(f).StartsWith("~$"));       // common processing
                //filesInDirList = ProcessFilesOrder(filesInDirList);
                string outputName = string.Empty;
                //if (filesInDirList[0] == null)
                //{
                //    Logger.LogE(string.Empty, new Exception("FILE ENTRY IS MISSING"));     // common processing
                //    continue;
                //}
                //Entry outputEntry = GetInfoFromWordFile(Path.Combine(entry, filesInDirList[0]));
                //var supplName = outputEntry.DocNumber + "_" + entry.Substring(entry.LastIndexOf("\\") + 1);
                outputName = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(entry)).Replace("/", "-");// supplname for common processing

                //MergeFiles(filesInDirList, entry, outputName); // common processing
                
                //ConvertFile(entry, outputName); // create xml only
                
                currentCount++; 
                UpdateProgressBar();
            }
            EnableButton();
        }

        private bool CheckFile(string entry)
        {
            bool res = false;

            string fileText = ReadFileText(entry);

            string docNumFromLoan = string.Empty;
            string docNumFromAct = string.Empty;
            string docNumFromRest = string.Empty;

            return res;
        }

        private void EnableButton()
        {
            if (this.InvokeRequired)
                this.Invoke(new MethodInvoker(EnableButton));
            else
                button1.Enabled = true;
        }

        private void UpdateProgressBar()
        {
            if (this.InvokeRequired)
                this.Invoke(new MethodInvoker(UpdateProgressBar));
            else
                progressBar1.Value = 100 * currentCount / totalCount;
        }

        private void ConvertFile(string fileName, string savePath)
        {
            Word.Application objWordApp = new Word.Application();
            Word.Document objWordDoc = new Word.Document();
            object missing = Type.Missing;
            object fileFormatXML = Word.WdSaveFormat.wdFormatXML;
            object objCreateDoc;

                try {
                    objWordDoc = objWordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                    objWordDoc.Activate();
                    objWordApp.Visible = false;

                    objWordApp.Selection.InsertFile(fileName, ref missing, true, ref missing, ref missing);

                    objCreateDoc = savePath + ".xml";
                    objWordApp.ActiveDocument.SaveAs(ref objCreateDoc, ref fileFormatXML, ref missing,
                                                    ref missing, ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing, ref missing);

                    objWordApp.ActiveDocument.Close(ref missing, ref missing, ref missing);
                    objWordApp.Quit();
                }
                catch { }
        }

        private void MergeFiles(List<string> filesInDirList, string entry, string outputName)
        {
            for (int i = 0; i < filesInDirList.Count; i++)
            {
                filesInDirList[i] = Path.Combine(entry, filesInDirList[i]);
            }

            Word.Application objWordApp = new Word.Application();
            Word.Document objWordDoc = new Word.Document();
            object missing = Type.Missing;
            object objCreateDoc = outputName.Replace(".", " ");
            object objPageBreak = Word.WdBreakType.wdPageBreak;
            object fileFormatDoc = Word.WdSaveFormat.wdFormatDocument;
            object fileFormatXML = Word.WdSaveFormat.wdFormatXML;

            try
            {
                objWordDoc = objWordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                objWordDoc.Activate();
                objWordApp.Visible = false;

                foreach (var fl in filesInDirList)
                {
                    objWordApp.Selection.InsertFile(fl, ref missing, true, ref missing, ref missing);
                    objWordApp.Selection.InsertBreak(ref objPageBreak);
                }
                objWordApp.ActiveDocument.SaveAs(ref objCreateDoc, ref fileFormatDoc, ref missing,
                                                 ref missing, ref missing, ref missing, ref missing,
                                                 ref missing, ref missing, ref missing, ref missing);
                objWordApp.ActiveDocument.SaveAs(ref objCreateDoc, ref fileFormatXML, ref missing,
                                                ref missing, ref missing, ref missing, ref missing,
                                                ref missing, ref missing, ref missing, ref missing);

                objWordApp.ActiveDocument.Close(ref missing, ref missing, ref missing);
                objWordApp.Quit();
            }
            catch (Exception ex)
            {
                Logger.LogE(string.Empty, ex);
                objWordApp.Quit();
            }
        }
              
        private List<string> ProcessFilesOrder(List<string> filesInDir)
        {
            try
            {
                List<string> res = new List<string>();
                var first = filesInDir.First(f => Path.GetFileName(f).ToLower().StartsWith("договор") || Path.GetFileName(f).ToLower().StartsWith("соц") || Path.GetFileName(f).ToLower().StartsWith("типовой")|| Path.GetFileName(f).ToLower().StartsWith("утвержден"));
                var seconds = filesInDir.FindAll(f => Path.GetFileName(f).ToLower().StartsWith("акт"));
                var others = filesInDir.FindAll(f => !Path.GetFileName(f).ToLower().StartsWith("договор") && !Path.GetFileName(f).ToLower().StartsWith("соц") && !Path.GetFileName(f).ToLower().StartsWith("акт") && !Path.GetFileName(f).ToLower().StartsWith("типовой") && !Path.GetFileName(f).ToLower().StartsWith("утвержден"));

                res.Add(first);
                res.AddRange(seconds);
                res.AddRange(others);

                if (res.Count == 0)
                    res.AddRange(filesInDir);

                return res;
            }
            catch (Exception ex)
            {
                Logger.LogE(string.Empty, ex);
                return filesInDir;
            }
        }

        private Entry GetInfoFromWordFile(string entry)
        {
            try
            {
                var result = new Entry();

                result.DocText = ReadFileText(entry);
                string fileText = result.DocText;

                var type = Constants.type1;
                for (int i = 0; i < 3; i++)
                {
                    if (fileText.IndexOf(type) == -1)
                    {
                        if (i == 0)
                            type = Constants.type2;
                        else if (i == 1)
                            type = Constants.type3;
                        else
                            type = Constants.type4;
                    }
                    else
                        break;
                }                

                var index = fileText.IndexOf(type) > -1 ? fileText.IndexOf(type) + type.Length + 1 : -1;
                if (index == -1)
                    result.DocNumber = "XX";
                else
                {
                    fileText = fileText.Substring(index);
                    var gIngex = fileText.IndexOf("г");
                    var docNum = fileText.Substring(0, gIngex);
                    result.DocNumber = docNum.Replace("\r", string.Empty).Replace("\n", string.Empty).Trim();
                }

                type = Constants.type5;
                index = fileText.IndexOf(type) > -1 ? fileText.IndexOf(type) + type.Length + 1 : -1;
                if (index == -1)
                    result.ActNumber = "XX";
                else
                {
                    fileText = fileText.Substring(index);
                    var gIngex = fileText.IndexOf("от");
                    var docNum = fileText.Substring(0, gIngex);
                    result.ActNumber = docNum.Replace("\r", string.Empty).Replace("\n", string.Empty).Trim();
                }

                return result;
            }
            catch (Exception ex)
            {
                Logger.LogE(string.Empty, ex);
                return null;
            }
        }

        private string ReadFileText(string entry)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                object miss = System.Reflection.Missing.Value;
                object path = entry;
                object readOnly = true;
                Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
                string totaltext = "";
                for (int i = 0; i < docs.Paragraphs.Count; i++)
                {
                    totaltext += " \r\n " + docs.Paragraphs[i + 1].Range.Text.ToString();
                }
                docs.Close();
                word.Quit();

                return totaltext;
            }
            catch (Exception ex)
            {
                Logger.LogE(string.Empty, ex);
                return null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            thd = new Thread(new ThreadStart(ParseFiles));
            thd.Start();

            this.UseWaitCursor = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {            
            MessageBox.Show("Если выполнене программы не закончено, не забудьте убить все процессы WINWORD.EXE в системе!");
            try
            {
                if (thd.IsAlive)
                    thd.Abort();
            }
            catch { }
        }
    }
}
