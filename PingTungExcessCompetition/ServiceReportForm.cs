using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using System.IO;
using Aspose.Words;
using FISCA.UDT;
using PingTungExcessCompetition.DAO;
using K12.Data;

namespace PingTungExcessCompetition
{
    public partial class ServiceReportForm : BaseForm
    {
        public Configure _Configure { get; private set; }

        BackgroundWorker bgWorkerReport = new BackgroundWorker();

        List<string> ClassIDList = new List<string>();

        AccessHelper _accessHelper = new AccessHelper();
        public ServiceReportForm()
        {
            InitializeComponent();
            bgWorkerReport.DoWork += BgWorkerReport_DoWork;
            bgWorkerReport.ProgressChanged += BgWorkerReport_ProgressChanged;
            bgWorkerReport.RunWorkerCompleted += BgWorkerReport_RunWorkerCompleted;
            bgWorkerReport.WorkerReportsProgress = true;
        }

        public void SetClassIDs(List<string> classIDs)
        {
            ClassIDList = classIDs;
        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            try
            {
                Document doc = (Document)e.Result;

                UserControlEnable(true);

                #region 儲存檔案

                string reportName = "屏東免試入學-班級服務表現";

                string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                path = Path.Combine(path, reportName + ".doc");

                if (File.Exists(path))
                {
                    int i = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                        if (!File.Exists(newPath))
                        {
                            path = newPath;
                            break;
                        }
                    }
                }

                Document document = new Document();
                try
                {
                    if (doc != null)
                        document = doc;
                    document.Save(path, SaveFormat.Doc);
                    System.Diagnostics.Process.Start(path);
                }
                catch
                {
                    System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = reportName + ".doc";
                    sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);

                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤," + ex.Message);
            }
        }

        private void BgWorkerReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("班級服務表現資料產生中...", e.ProgressPercentage);
        }

        private void BgWorkerReport_DoWork(object sender, DoWorkEventArgs e)
        {

            DataTable dtTable = new DataTable();
            bgWorkerReport.ReportProgress(1);

            Document docTemplate = _Configure.Template;
            if (docTemplate == null)
                docTemplate = new Document(new MemoryStream(Properties.Resources.屏東班級服務表現樣板));


            // 校名
            string SchoolName = K12.Data.School.ChineseName;

            // 取得班導師
            Dictionary<string, string> ClassTeacherNameDict = QueryData.GetClassTeacherNameDictByClassID(ClassIDList);

            // 產生合併欄位
            dtTable.Columns.Add("學校名稱");
            dtTable.Columns.Add("班級名稱");
            dtTable.Columns.Add("學年度");

            for (int studIdx = 1; studIdx <= 100; studIdx++)
            {
                dtTable.Columns.Add("座號" + studIdx);
                dtTable.Columns.Add("姓名" + studIdx);
                dtTable.Columns.Add("服務表現項目7上" + studIdx);
                dtTable.Columns.Add("服務表現項目7下" + studIdx);
                dtTable.Columns.Add("服務表現項目8上" + studIdx);
                dtTable.Columns.Add("服務表現項目8下" + studIdx);
                dtTable.Columns.Add("服務表現項目9上" + studIdx);
                dtTable.Columns.Add("服務表現積分" + studIdx);
                dtTable.Columns.Add("服務表現備註" + studIdx);
            }

            // 取得班級學生相關資料
            Dictionary<string, List<StudentInfo>> ClassStudentDict = QueryData.GetClassStudentDict(ClassIDList);
            Dictionary<string, ClassRecord> classRecDict = new Dictionary<string, ClassRecord>();
            List<ClassRecord> claRecList = Class.SelectByIDs(ClassIDList);
            foreach (ClassRecord data in claRecList)
            {
                if (!classRecDict.ContainsKey(data.ID))
                    classRecDict.Add(data.ID, data);
            }

            //StreamWriter sw1 = new StreamWriter(Application.StartupPath + "\\合併欄位.txt");
            //StringBuilder sb1 = new StringBuilder();
            //foreach (DataColumn dc in dtTable.Columns)
            //    sb1.AppendLine(dc.Caption);

            //sw1.Write(sb1.ToString());
            //sw1.Close();

            // 排序後班級
            foreach (string class_id in ClassTeacherNameDict.Keys)
            {
                DataRow row = dtTable.NewRow();
                row["學校名稱"] = SchoolName;
                if (classRecDict.ContainsKey(class_id))
                {
                    row["班級名稱"] = classRecDict[class_id].Name;
                }

                int sc;
                if (int.TryParse(K12.Data.School.DefaultSchoolYear, out sc))
                {
                    row["學年度"] = sc + 1;
                }
                if (ClassStudentDict.ContainsKey(class_id))
                {
                    int studIdx = 1;
                    foreach (StudentInfo si in ClassStudentDict[class_id])
                    {

                        row["座號" + studIdx] = si.SeatNo;
                        row["姓名" + studIdx] = si.StudentName;
                        row["服務表現項目7上" + studIdx] = string.Join("\n", si.ServiceItem_7a.ToArray());
                        row["服務表現項目7下" + studIdx] = string.Join("\n", si.ServiceItem_7b.ToArray());
                        row["服務表現項目8上" + studIdx] = string.Join("\n", si.ServiceItem_8a.ToArray());
                        row["服務表現項目8下" + studIdx] = string.Join("\n", si.ServiceItem_8b.ToArray());
                        row["服務表現項目9上" + studIdx] = string.Join("\n", si.ServiceItem_9a.ToArray());
                        row["服務表現積分" + studIdx] = si.ServiceScore;
                        row["服務表現備註" + studIdx] = string.Join("\n", si.ServiceMemo.ToArray());
                        studIdx++;
                    }
                }

                dtTable.Rows.Add(row);
            }

            //// debug
            //dtTable.TableName = "debug";
            //dtTable.WriteXml(Application.StartupPath + "\\debug.xml");

            Document doc = _Configure.Template;
            doc.MailMerge.Execute(dtTable);
            doc.MailMerge.DeleteFields();
            e.Result = doc;
            bgWorkerReport.ReportProgress(100);
        }

        private void LoadTemplate()
        {
            try
            {
                UserControlEnable(false);
                List<Configure> confList = _accessHelper.Select<Configure>();
                if (confList != null && confList.Count > 0)
                {
                    _Configure = confList[0];
                    _Configure.Decode();
                }
                else
                {
                    _Configure = new Configure();
                    _Configure.Name = "屏東免試入學-班級服務表現";
                    _Configure.Template = new Document(new MemoryStream(Properties.Resources.屏東班級服務表現樣板));
                    _Configure.Encode();

                }
                _Configure.Save();
                UserControlEnable(true);
            }
            catch (Exception ex)
            {
                MsgBox.Show("讀取樣板發生錯誤，" + ex.Message);
            }

        }

        private void ServiceReportForm_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;
            LoadTemplate();
        }

        private void UserControlEnable(bool value)
        {
            lnkChangeTemplate.Enabled = value;
            lnkViewMapColumns.Enabled = value;
            lnkViewTemplate.Enabled = value;
            btnSetCadreName.Enabled = value;
            btnPrint.Enabled = value;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            UserControlEnable(false);

            bgWorkerReport.RunWorkerAsync();
        }

        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = "屏東免試入學-班級服務表現";

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                if (_Configure.Template == null)
                    _Configure.Template = new Document(new MemoryStream(Properties.Resources.屏東班級服務表現樣板));

                _Configure.Template.Save(path, Aspose.Words.SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        _Configure.Template.Save(path, Aspose.Words.SaveFormat.Doc);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            lnkViewTemplate.Enabled = true;
            #endregion
        }

        private void lnkChangeTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            lnkChangeTemplate.Enabled = false;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    _Configure.Template = new Aspose.Words.Document(dialog.FileName);
                    List<string> fields = new List<string>(_Configure.Template.MailMerge.GetFieldNames());
                    _Configure.Encode();
                    _Configure.Save();

                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }
            lnkChangeTemplate.Enabled = true;
        }

        private void lnkViewMapColumns_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            FISCA.Presentation.MotherForm.SetStatusBarMessage("班級學期成績合併欄位總表產生中...");

            // 產生合併欄位總表
            lnkViewMapColumns.Enabled = false;
            Global.ExportMappingFieldWord();
            lnkViewMapColumns.Enabled = true;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void lnkDefault_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            lnkDefault.Enabled = false;
            string reportName = "屏東免試入學-班級服務表現";

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            Document DefDoc = null;
            try
            {
                DefDoc = new Document(new MemoryStream(Properties.Resources.屏東班級服務表現樣板));

                DefDoc.Save(path, Aspose.Words.SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        DefDoc.Save(path, Aspose.Words.SaveFormat.Doc);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            lnkDefault.Enabled = true;

        }

        private void btnSetCadreName_Click(object sender, EventArgs e)
        {
            btnSetCadreName.Enabled = false;

            setCadreNameForm scnf = new setCadreNameForm();
            scnf.ShowDialog();



            btnSetCadreName.Enabled = true;
        }
    }
}
