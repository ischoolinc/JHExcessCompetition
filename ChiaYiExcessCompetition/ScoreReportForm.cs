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
using System.Xml.Linq;

namespace ChiaYiExcessCompetition
{
    public partial class ScoreReportForm : BaseForm
    {
        public Configure _Configure { get; private set; }

        BackgroundWorker bgWorkerReport = new BackgroundWorker();

        AccessHelper _accessHelper = new AccessHelper();

        List<string> StudentIDList = new List<string>();

        Dictionary<string, Document> StudentDocDict = new Dictionary<string, Document>();

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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("成績冊產生中...", e.ProgressPercentage);
        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");




        }

        private void BgWorkerReport_DoWork(object sender, DoWorkEventArgs e)
        {

            bgWorkerReport.ReportProgress(1);

            // 讀取資料
            // 取得所選學生資料
            List<rptStudentInfo> StudentInfoList = QueryData.GetRptStudentInfoListByIDs(StudentIDList);

            // 轉換各項類別對照值
            Dictionary<string, string> MappingTag1 = new Dictionary<string, string>();
            Dictionary<string, string> MappingTag2 = new Dictionary<string, string>();
            Dictionary<string, string> MappingTag3 = new Dictionary<string, string>();
            Dictionary<string, string> MappingTag4 = new Dictionary<string, string>();

            // 取得學生類別
            Dictionary<string, List<string>> StudentTagDict = QueryData.GetStudentTagName(StudentIDList);

            // 解析對照設定
            XElement elmRoot = XElement.Parse(_Configure.MappingContent);
            if (elmRoot != null)
            {
                foreach (XElement elm in elmRoot.Elements("Group"))
                {
                    string gpName = elm.Attribute("Name").Value;
                    if (gpName == "學生身分")
                    {
                        foreach (XElement elm1 in elm.Elements("Item"))
                        {
                            string tagName = elm1.Attribute("TagName").Value;
                            if (!MappingTag1.ContainsKey(tagName) && tagName.Length > 0)
                                MappingTag1.Add(tagName, elm1.Attribute("Code").Value);
                        }
                    }

                    if (gpName == "身心障礙")
                    {
                        foreach (XElement elm1 in elm.Elements("Item"))
                        {
                            string tagName = elm1.Attribute("TagName").Value;
                            if (!MappingTag2.ContainsKey(tagName) && tagName.Length > 0)
                                MappingTag2.Add(tagName, elm1.Attribute("Code").Value);
                        }
                    }

                    if (gpName == "學生報名身分設定")
                    {
                        foreach (XElement elm1 in elm.Elements("Item"))
                        {
                            string tagName = elm1.Attribute("TagName").Value;
                            if (!MappingTag3.ContainsKey(tagName) && tagName.Length > 0)
                                MappingTag3.Add(tagName, elm1.Attribute("Code").Value);
                        }
                    }

                    if (gpName == "失業勞工子女")
                    {
                        foreach (XElement elm1 in elm.Elements("Item"))
                        {
                            string tagName = elm1.Attribute("TagName").Value;
                            if (!MappingTag4.ContainsKey(tagName) && tagName.Length > 0)
                                MappingTag4.Add(tagName, elm1.Attribute("Code").Value);
                        }
                    }
                }
            }

            // 填入身心障礙生，計算體適能會用到
            foreach (rptStudentInfo si in StudentInfoList)
            {
                if (StudentTagDict.ContainsKey(si.StudentID))
                {
                    foreach (string tagName in StudentTagDict[si.StudentID])
                    {
                        if (MappingTag2.ContainsKey(tagName))
                        {
                            si.isSpecial = true;
                        }
                    }
                }
            }
            bgWorkerReport.ReportProgress(10);
            // 取得學生學期成績
            StudentInfoList = QueryData.FillRptDomainScoreInfo(StudentIDList, StudentInfoList);
            bgWorkerReport.ReportProgress(20);
            // 取得獎懲紀錄
            StudentInfoList = QueryData.FillRptMeritDemeritInfo(StudentIDList, StudentInfoList, _Configure.EndDate);

            // 取得服務學習紀錄
            StudentInfoList = QueryData.FillRptServiceInfo(StudentIDList, StudentInfoList, _Configure.EndDate);

            // 取得體適能紀錄
            StudentInfoList = QueryData.FillRptFitnessInfo(StudentIDList, StudentInfoList);

            // 取得低收入並填入
            StudentInfoList = QueryData.FillrptIncomeType(StudentIDList, StudentInfoList);
            
            bgWorkerReport.ReportProgress(50);

            StudentDocDict.Clear();

            // 整理資料，填入 DataTable
            foreach (rptStudentInfo si in StudentInfoList)
            {
                // 因為每位學生，所以用複製一份
                Document docTemplate = _Configure.Template.Clone();
                if (docTemplate == null)
                    docTemplate = new Document(new MemoryStream(Properties.Resources.嘉義區成績冊樣板));


                // 每位學生一個 Word檔案
                #region 產生合併欄位
                DataTable dtTable = new DataTable();

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
                }
                for (int i = 1; i <= 12; i++)
                {
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





                DataRow row = dtTable.NewRow();

                row["學年度"] = si.SchoolYear;
                row["學校名稱"] = K12.Data.School.ChineseName;
                row["班級"] = si.ClassName;
                row["座號"] = si.SeatNo;
                row["姓名"] = si.Name;

                if (si.IncomeType1)
                {
                    row["扶助弱勢_身分"] = "低收入戶";
                    row["扶助弱勢_積分"] = 1;
                }
                else
                {
                    row["扶助弱勢_身分"] = "";
                    row["扶助弱勢_積分"] = 0;
                }

                row["均衡學習_健康與體育_七上分數"] = "";
                row["均衡學習_健康與體育_七下分數"] = "";
                row["均衡學習_健康與體育_八上分數"] = "";
                row["均衡學習_健康與體育_八下分數"] = "";
                row["均衡學習_健康與體育_九上分數"] = "";
                row["均衡學習_健康與體育_平均"] = "";
                row["均衡學習_藝術與人文_七上分數"] = "";
                row["均衡學習_藝術與人文_七下分數"] = "";
                row["均衡學習_藝術與人文_八上分數"] = "";
                row["均衡學習_藝術與人文_八下分數"] = "";
                row["均衡學習_藝術與人文_九上分數"] = "";
                row["均衡學習_藝術與人文_平均"] = "";
                row["均衡學習_綜合活動_七上分數"] = "";
                row["均衡學習_綜合活動_七下分數"] = "";
                row["均衡學習_綜合活動_八上分數"] = "";
                row["均衡學習_綜合活動_八下分數"] = "";
                row["均衡學習_綜合活動_九上分數"] = "";
                row["均衡學習_綜合活動_平均"] = "";
                row["均衡學習_積分"] = si.DomainIScore;


                row["品德表現_獎懲_大功統計"] = si.MASum;
                row["品德表現_獎懲_小功統計"] = si.MBSum;
                row["品德表現_獎懲_嘉獎統計"] = si.MCSum;
                row["品德表現_獎懲_大過統計"] = si.DASum;
                row["品德表現_獎懲_小過統計"] = si.DBSum;
                row["品德表現_獎懲_警告統計"] = si.DCSum;
                row["品德表現_獎懲_銷過統計"] = si.MDCleanCount;
                row["品德表現_獎懲_積分"] = si.MDIScore;


                row["品德表現_服務學習_積分"] = si.ServiceIScore;
                row["品德表現_服務學習_校內時數統計"] = si.ServiceInHourCount;
                row["品德表現_服務學習_校外時數統計"] = si.ServiceOutHourCount;

                row["品德表現_體適能_積分"] = si.FitnessIScore;




                dtTable.Rows.Add(row);

                // debug 
                dtTable.TableName = "debug";
                dtTable.WriteXml(Application.StartupPath + @"\debug" + si.StudentID + ".xml");

                docTemplate.MailMerge.Execute(dtTable);
                docTemplate.MailMerge.RemoveEmptyParagraphs = true;
                docTemplate.MailMerge.DeleteFields();

                if (!StudentDocDict.ContainsKey(si.StudentID))
                    StudentDocDict.Add(si.StudentID, docTemplate);

            }


            bgWorkerReport.ReportProgress(90);


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
