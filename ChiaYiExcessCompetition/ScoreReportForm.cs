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
using ChiaYiExcessCompetition.DAO;
using K12.Data;

namespace ChiaYiExcessCompetition
{
    public partial class ScoreReportForm : BaseForm
    {
        public Configure _Configure { get; private set; }

        BackgroundWorker bgWorkerReport = new BackgroundWorker();

        AccessHelper _accessHelper = new AccessHelper();

        List<string> StudentIDList = new List<string>();

        public ScoreReportForm()
        {
            InitializeComponent();
            bgWorkerReport.DoWork += BgWorkerReport_DoWork;
            bgWorkerReport.RunWorkerCompleted += BgWorkerReport_RunWorkerCompleted;
            bgWorkerReport.ProgressChanged += BgWorkerReport_ProgressChanged;
            bgWorkerReport.WorkerReportsProgress = true;
        }

        private void BgWorkerReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void BgWorkerReport_DoWork(object sender, DoWorkEventArgs e)
        {
            DataTable dtTable = new DataTable();
            bgWorkerReport.ReportProgress(1);

            // 讀取資料
            // 取得所選學生資料
            List<rptStudentInfo> StudentInfoList = QueryData.GetRptStudentInfoListByIDs(StudentIDList);

            // 取得學生學期成績
            StudentInfoList = QueryData.FillRptDomainScoreInfo(StudentIDList, StudentInfoList);

            // 取得獎懲紀錄
            StudentInfoList = QueryData.FillRptMeritDemeritInfo(StudentIDList, StudentInfoList, _Configure.EndDate);

            // 取得服務學習紀錄



            // 取得體適能紀錄

            Console.Write(1);

            // 整理資料，填入 DataTable


            // debug 
            dtTable.TableName = "debug";
            dtTable.WriteXml(Application.StartupPath + @"\debug.xml");





            List<Document> docList = new List<Document>();

            Document docTemplate = _Configure.Template;
            if (docTemplate == null)
                docTemplate = new Document(new MemoryStream(Properties.Resources.嘉義區成績冊樣板));

            // 每位學生一個 Word檔案
            #region 產生合併欄位

            dtTable.Columns.Add("學年度");
            dtTable.Columns.Add("學校名稱");
            dtTable.Columns.Add("班級");
            dtTable.Columns.Add("座號");
            dtTable.Columns.Add("姓名");
            dtTable.Columns.Add("扶助弱勢_身分");
            dtTable.Columns.Add("扶助弱勢_積分");
            dtTable.Columns.Add("均衡學習_健康與體育_七上分數");
            dtTable.Columns.Add("均衡學習_健康與體育_七下分數");
            dtTable.Columns.Add("均衡學習_健康與體育_八上分數");
            dtTable.Columns.Add("均衡學習_健康與體育_八下分數");
            dtTable.Columns.Add("均衡學習_健康與體育_九上分數");
            dtTable.Columns.Add("均衡學習_健康與體育_平均");
            dtTable.Columns.Add("均衡學習_藝術與人文_七上分數");
            dtTable.Columns.Add("均衡學習_藝術與人文_七下分數");
            dtTable.Columns.Add("均衡學習_藝術與人文_八上分數");
            dtTable.Columns.Add("均衡學習_藝術與人文_八下分數");
            dtTable.Columns.Add("均衡學習_藝術與人文_九上分數");
            dtTable.Columns.Add("均衡學習_藝術與人文_平均");
            dtTable.Columns.Add("均衡學習_綜合活動_七上分數");
            dtTable.Columns.Add("均衡學習_綜合活動_七下分數");
            dtTable.Columns.Add("均衡學習_綜合活動_八上分數");
            dtTable.Columns.Add("均衡學習_綜合活動_八下分數");
            dtTable.Columns.Add("均衡學習_綜合活動_九上分數");
            dtTable.Columns.Add("均衡學習_綜合活動_平均");
            dtTable.Columns.Add("均衡學習_積分");
            dtTable.Columns.Add("品德表現_獎懲_大功統計");
            dtTable.Columns.Add("品德表現_獎懲_小功統計");
            dtTable.Columns.Add("品德表現_獎懲_嘉獎統計");
            dtTable.Columns.Add("品德表現_獎懲_大過統計");
            dtTable.Columns.Add("品德表現_獎懲_小過統計");
            dtTable.Columns.Add("品德表現_獎懲_警告統計");
            dtTable.Columns.Add("品德表現_獎懲_銷過統計");
            dtTable.Columns.Add("品德表現_獎懲_積分");
            dtTable.Columns.Add("品德表現_服務學習_積分");
            dtTable.Columns.Add("品德表現_服務學習_校內時數統計");
            dtTable.Columns.Add("品德表現_服務學習_校外時數統計");
            dtTable.Columns.Add("品德表現_體適能_積分");

            for (int i = 1; i <= 100; i++)
            {
                dtTable.Columns.Add("品德表現_獎懲_獎懲日期" + i);
                dtTable.Columns.Add("品德表現_獎懲_學期" + i);
                dtTable.Columns.Add("品德表現_獎懲_獎懲事由" + i);
                dtTable.Columns.Add("品德表現_獎懲_大功" + i);
                dtTable.Columns.Add("品德表現_獎懲_小功" + i);
                dtTable.Columns.Add("品德表現_獎懲_嘉獎" + i);
                dtTable.Columns.Add("品德表現_獎懲_大過" + i);
                dtTable.Columns.Add("品德表現_獎懲_小過" + i);
                dtTable.Columns.Add("品德表現_獎懲_警告" + i);
                dtTable.Columns.Add("品德表現_獎懲_銷過" + i);


            }

            for (int i = 1; i <= 50; i++)
            {
                dtTable.Columns.Add("品德表現_服務學習_資料輸入日期" + i);
                dtTable.Columns.Add("品德表現_服務學習_校內外" + i);
                dtTable.Columns.Add("品德表現_服務學習_服務時數" + i);
                dtTable.Columns.Add("品德表現_服務學習_服務學習活動內容" + i);
                dtTable.Columns.Add("品德表現_服務學習_服務學習證明單位" + i);

                dtTable.Columns.Add("品德表現_體適能_檢測日期" + i);
                dtTable.Columns.Add("品德表現_體適能_年齡" + i);
                dtTable.Columns.Add("品德表現_體適能_性別" + i);
                dtTable.Columns.Add("品德表現_體適能_坐姿體前彎_成績" + i);
                dtTable.Columns.Add("品德表現_體適能_立定跳遠_成績" + i);
                dtTable.Columns.Add("品德表現_體適能_仰臥起坐_成績" + i);
                dtTable.Columns.Add("品德表現_體適能_公尺跑走_成績" + i);
                dtTable.Columns.Add("品德表現_體適能_坐姿體前彎_等級" + i);
                dtTable.Columns.Add("品德表現_體適能_立定跳遠_等級" + i);
                dtTable.Columns.Add("品德表現_體適能_仰臥起坐_等級" + i);
                dtTable.Columns.Add("品德表現_體適能_公尺跑走_等級" + i);
            }
            #endregion

            StreamWriter sw1 = new StreamWriter(Application.StartupPath + "\\合併欄位.txt");
            StringBuilder sb1 = new StringBuilder();
            foreach (DataColumn dc in dtTable.Columns)
                sb1.AppendLine(dc.Caption);

            sw1.Write(sb1.ToString());
            sw1.Close();

            // 校名
            string SchoolName = K12.Data.School.ChineseName;

        }

        public void SetStudentIDs(List<string> studIDs)
        {
            StudentIDList = studIDs;
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

        private void lnkDefault_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkDefault.Enabled = false;
            string reportName = "嘉義免試入學-成績冊";

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
                DefDoc = new Document(new MemoryStream(Properties.Resources.嘉義區成績冊樣板));

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

        private void lnkViewMapColumns_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("成績冊合併欄位總表產生中...");

            // 產生合併欄位總表
            lnkViewMapColumns.Enabled = false;
            Global.ExportMappingFieldWord();
            lnkViewMapColumns.Enabled = true;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
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

        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = "嘉義免試入學-成績冊";

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
                    _Configure.Template = new Document(new MemoryStream(Properties.Resources.嘉義區成績冊樣板));

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

        private void ScoreReportForm_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;
            LoadTemplate();
            this.Text = "成績冊(資料統計至" + _Configure.EndDate.Year + "年" + _Configure.EndDate.Month + "月" + _Configure.EndDate.Day + "日)";
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
                    if (_Configure.EndDate.Date.Year < DateTime.Now.Year)
                        _Configure.EndDate = new DateTime(DateTime.Now.Year, 4, 30);
                }
                else
                {
                    _Configure = new Configure();
                    _Configure.Name = "嘉義免試入學-成績冊";
                    _Configure.Template = new Document(new MemoryStream(Properties.Resources.嘉義區成績冊樣板));
                    if (_Configure.EndDate.Date.Year < DateTime.Now.Year)
                        _Configure.EndDate = new DateTime(DateTime.Now.Year, 4, 30);

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

        private void UserControlEnable(bool value)
        {
            lnkChangeTemplate.Enabled = value;
            lnkViewMapColumns.Enabled = value;
            lnkViewTemplate.Enabled = value;
            btnPrint.Enabled = value;
        }
    }
}
