﻿using System;
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

        List<string> errorMsgList = new List<string>();

        bool IsSingleFileSaved = true;

        Dictionary<string, Document> StudentDocDict = new Dictionary<string, Document>();
        Dictionary<string, string> StudentDocNameDict = new Dictionary<string, string>();
        Document _Doc = new Document();
        List<Document> _DocList = new List<Document>();

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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("超額比序均衡學習成績證明單產生中...", e.ProgressPercentage);
        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            if (errorMsgList.Count > 0)
            {
                MsgBox.Show("無法產生檔案，請到 教務作業>資料統計>報表>嘉義免試入學>送審用匯入檔，設定學生身分。");
                return;
            }


            string path = "";
            // 完成後開啟資料夾
            string p1 = "";


            if (IsSingleFileSaved)
            {
                if (StudentDocDict.Count > 0)
                    foreach (string sid in StudentDocDict.Keys)
                    {
                        try
                        {
                            Document document = StudentDocDict[sid];

                            #region 儲存檔案
                            string reportName = "嘉義免試入學-超額比序均衡學習成績證明單" + StudentDocNameDict[sid];
                            path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports\\嘉義免試入學超額比序均衡學習成績證明單");

                            p1 = path;
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
                                document.Save(path, SaveFormat.Doc);
                            }
                            catch (Exception ex)
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
                else
                {
                    return;
                }
            }
            else
            {
                try
                {

                    Document document = _Doc;

                    #region 儲存檔案
                    string reportName = "嘉義免試入學-超額比序均衡學習成績證明單";
                    path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports\\嘉義免試入學超額比序均衡學習成績證明單");

                    p1 = path;
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
                        document.Save(path, SaveFormat.Doc);
                    }
                    catch (Exception ex)
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

                _Doc = null;
                _DocList.Clear();
                GC.Collect();
            }
            if (p1 != "")
                System.Diagnostics.Process.Start(p1);
        }

        private void BgWorkerReport_DoWork(object sender, DoWorkEventArgs e)
        {
            errorMsgList.Clear();
            bgWorkerReport.ReportProgress(1);

            // 讀取資料
            // 取得所選學生資料
            List<rptStudentInfo> StudentInfoList = QueryData.GetRptStudentInfoListByIDs(StudentIDList);

            if (StudentInfoList.Count < 1)
            {
                MsgBox.Show("所選的學生皆非「在校生」，沒有可列印的資料。", "超額比序均衡學習成績證明單",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            // 轉換各項類別對照值
            Dictionary<string, string> MappingTag1 = new Dictionary<string, string>();
            Dictionary<string, string> MappingTag2 = new Dictionary<string, string>();
            Dictionary<string, string> MappingTag3 = new Dictionary<string, string>();
            Dictionary<string, string> MappingTag4 = new Dictionary<string, string>();

            // 取得學生類別
            Dictionary<string, List<string>> StudentTagDict = QueryData.GetStudentTagName(StudentIDList);

            try
            {
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
            }
            catch (Exception ex)
            {
                errorMsgList.Add(ex.Message);
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

            // 取得學生競賽資料並填入
            StudentInfoList = QueryData.FillStudentCompetitionPerformance(StudentIDList, StudentInfoList);


            bgWorkerReport.ReportProgress(50);

            StudentDocDict.Clear();
            StudentDocNameDict.Clear();

            List<string> domainNameList = new List<string>();
            domainNameList.Add("健康與體育");
            domainNameList.Add("藝術");
            domainNameList.Add("綜合活動");
            domainNameList.Add("科技");

            List<string> otherDomainNameList = new List<string>();
            otherDomainNameList.Add("語文");
            otherDomainNameList.Add("數學");
            otherDomainNameList.Add("自然科學");
            //otherDomainNameList.Add("科技");

            //競賽統計
            Dictionary<string, int> CompPerformanceCountDict = new Dictionary<string, int>();



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
                dtTable.Columns.Add("學號");
                dtTable.Columns.Add("扶助弱勢_身分");
                dtTable.Columns.Add("扶助弱勢_積分");

                foreach (string name in domainNameList)
                {
                    dtTable.Columns.Add("均衡學習_" + name + "_七上分數");
                    dtTable.Columns.Add("均衡學習_" + name + "_七下分數");
                    dtTable.Columns.Add("均衡學習_" + name + "_八上分數");
                    dtTable.Columns.Add("均衡學習_" + name + "_八下分數");
                    dtTable.Columns.Add("均衡學習_" + name + "_九上分數");
                    dtTable.Columns.Add("均衡學習_" + name + "_平均");
                }

                foreach (string name in otherDomainNameList)
                {
                    dtTable.Columns.Add("其他領域_" + name + "_七上分數");
                    dtTable.Columns.Add("其他領域_" + name + "_七下分數");
                    dtTable.Columns.Add("其他領域_" + name + "_八上分數");
                    dtTable.Columns.Add("其他領域_" + name + "_八下分數");
                    dtTable.Columns.Add("其他領域_" + name + "_九上分數");
                    dtTable.Columns.Add("其他領域_" + name + "_平均");
                }

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
                dtTable.Columns.Add("品德表現_服務學習_校內外時數統計");
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
                    dtTable.Columns.Add("品德表現_體適能_仰臥捲腹_成績" + i);
                    dtTable.Columns.Add("品德表現_體適能_漸速耐力跑_成績" + i);
                    dtTable.Columns.Add("品德表現_體適能_仰臥捲腹_等級" + i);
                    dtTable.Columns.Add("品德表現_體適能_漸速耐力跑_等級" + i);
                }
                #endregion

                // 加入競賽成績合併欄位
                for (int i = 1; i <= 30; i++)
                {
                    dtTable.Columns.Add("競賽成績_競賽層級" + i);
                    dtTable.Columns.Add("競賽成績_競賽性質" + i);
                    dtTable.Columns.Add("競賽成績_競賽名稱" + i);
                    dtTable.Columns.Add("競賽成績_得獎名次" + i);
                    dtTable.Columns.Add("競賽成績_證書日期" + i);
                    dtTable.Columns.Add("競賽成績_主辦單位" + i);
                }

                CompPerformanceCountDict.Clear();

                for (int i = 1; i <= 8; i++)
                {
                    dtTable.Columns.Add("競賽成績_縣市個人名次" + i);
                    dtTable.Columns.Add("競賽成績_縣市團體名次" + i);
                    dtTable.Columns.Add("競賽成績_全國個人名次" + i);
                    dtTable.Columns.Add("競賽成績_全國團體名次" + i);

                    CompPerformanceCountDict.Add("競賽成績_縣市個人名次" + i, 0);
                    CompPerformanceCountDict.Add("競賽成績_縣市團體名次" + i, 0);
                    CompPerformanceCountDict.Add("競賽成績_全國個人名次" + i, 0);
                    CompPerformanceCountDict.Add("競賽成績_全國團體名次" + i, 0);
                }

                dtTable.Columns.Add("競賽成績_縣市個人名次_其他");
                dtTable.Columns.Add("競賽成績_縣市團體名次_其他");
                dtTable.Columns.Add("競賽成績_全國個人名次_其他");
                dtTable.Columns.Add("競賽成績_全國團體名次_其他");
                CompPerformanceCountDict.Add("競賽成績_縣市個人名次_其他", 0);
                CompPerformanceCountDict.Add("競賽成績_縣市團體名次_其他", 0);
                CompPerformanceCountDict.Add("競賽成績_全國個人名次_其他", 0);
                CompPerformanceCountDict.Add("競賽成績_全國團體名次_其他", 0);

                dtTable.Columns.Add("競賽成績_競賽積分");
                dtTable.Columns.Add("成績_合計總分");

                //StreamWriter sw1 = new StreamWriter(Application.StartupPath + "\\合併欄位.txt");
                //StringBuilder sb1 = new StringBuilder();
                //foreach (DataColumn dc in dtTable.Columns)
                //    sb1.AppendLine(dc.Caption);

                //sw1.Write(sb1.ToString());
                //sw1.Close();





                DataRow row = dtTable.NewRow();

                row["學年度"] = si.SchoolYear;
                row["學校名稱"] = K12.Data.School.ChineseName;
                row["班級"] = si.ClassName;
                row["座號"] = si.SeatNo;
                row["姓名"] = si.Name;
                row["學號"] = si.StudentNumber;

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

                // 處理領域學期成績
                foreach (string name in domainNameList)
                {
                    if (si.DomainScoreInfoDict.ContainsKey(name))
                    {
                        row["均衡學習_" + name + "_七上分數"] = si.GetDomainSemsScore(name, "七上");
                        row["均衡學習_" + name + "_七下分數"] = si.GetDomainSemsScore(name, "七下");
                        row["均衡學習_" + name + "_八上分數"] = si.GetDomainSemsScore(name, "八上");
                        row["均衡學習_" + name + "_八下分數"] = si.GetDomainSemsScore(name, "八下");
                        row["均衡學習_" + name + "_九上分數"] = si.GetDomainSemsScore(name, "九上");
                        row["均衡學習_" + name + "_平均"] = si.DomainScoreInfoDict[name].AvgScore;
                    }
                }

                foreach (string name in otherDomainNameList)
                {
                    if (si.DomainScoreInfoDict.ContainsKey(name))
                    {
                        row["其他領域_" + name + "_七上分數"] = si.GetDomainSemsScore(name, "七上");
                        row["其他領域_" + name + "_七下分數"] = si.GetDomainSemsScore(name, "七下");
                        row["其他領域_" + name + "_八上分數"] = si.GetDomainSemsScore(name, "八上");
                        row["其他領域_" + name + "_八下分數"] = si.GetDomainSemsScore(name, "八下");
                        row["其他領域_" + name + "_九上分數"] = si.GetDomainSemsScore(name, "九上");
                        row["其他領域_" + name + "_平均"] = si.DomainScoreInfoDict[name].AvgScore;
                    }
                }

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
                row["品德表現_服務學習_校內外時數統計"] = si.ServiceInHourCount + si.ServiceOutHourCount;
                row["品德表現_體適能_積分"] = si.FitnessIScore;

                int idx = 1;
                // 處理獎懲明細
                foreach (rptMeritDemeritInfo md in si.MeritDemeritInfoList)
                {
                    row["品德表現_獎懲_獎懲日期" + idx] = md.OccurDate;
                    row["品德表現_獎懲_學期" + idx] = md.Semester;
                    row["品德表現_獎懲_獎懲事由" + idx] = md.Reason;
                    row["品德表現_獎懲_大功" + idx] = md.MA;
                    row["品德表現_獎懲_小功" + idx] = md.MB;
                    row["品德表現_獎懲_嘉獎" + idx] = md.MC;
                    row["品德表現_獎懲_大過" + idx] = md.DA;
                    row["品德表現_獎懲_小過" + idx] = md.DB;
                    row["品德表現_獎懲_警告" + idx] = md.DC;

                    if (md.Cleand == "是")
                        row["品德表現_獎懲_銷過" + idx] = "是";
                    else
                        row["品德表現_獎懲_銷過" + idx] = "--";

                    idx++;
                }

                idx = 1;
                // 處理服務學習
                foreach (rptServiceInfo sif in si.ServiceInfoList)
                {
                    row["品德表現_服務學習_資料輸入日期" + idx] = sif.OccurDate;
                    row["品德表現_服務學習_校內外" + idx] = sif.InternalOrExternal;
                    row["品德表現_服務學習_服務時數" + idx] = sif.Hours;
                    row["品德表現_服務學習_服務學習活動內容" + idx] = sif.Reason;
                    row["品德表現_服務學習_服務學習證明單位" + idx] = sif.Organizers;
                    idx++;
                }

                idx = 1;
                // 體適能
                foreach (rptFitnessInfo fi in si.FitnessInfoList)
                {
                    row["品德表現_體適能_檢測日期" + idx] = fi.TestDateStr;
                    row["品德表現_體適能_年齡" + idx] = fi.Age;
                    row["品德表現_體適能_性別" + idx] = si.Gender;
                    row["品德表現_體適能_坐姿體前彎_成績" + idx] = fi.Sit_and_reach;
                    row["品德表現_體適能_立定跳遠_成績" + idx] = fi.Standing_long_jump;
                    row["品德表現_體適能_仰臥起坐_成績" + idx] = fi.Sit_up;
                    row["品德表現_體適能_公尺跑走_成績" + idx] = fi.Cardiorespiratory;
                    row["品德表現_體適能_坐姿體前彎_等級" + idx] = fi.Sit_and_reach_degree;
                    row["品德表現_體適能_立定跳遠_等級" + idx] = fi.Standing_long_jump_degree;
                    row["品德表現_體適能_仰臥起坐_等級" + idx] = fi.Sit_up_degree;
                    row["品德表現_體適能_公尺跑走_等級" + idx] = fi.Cardiorespiratory_degree;

                    row["品德表現_體適能_仰臥捲腹_成績" + idx] = fi.Curl;
                    row["品德表現_體適能_仰臥捲腹_等級" + idx] = fi.CurlDegree;
                    row["品德表現_體適能_漸速耐力跑_成績" + idx] = fi.Pacer;
                    row["品德表現_體適能_漸速耐力跑_等級" + idx] = fi.PacerDegree;

                    idx++;
                }

                idx = 1;
                // 填入競賽基本資料
                foreach (rptCompPerformanceInfo cpi in si.CompPerformanceInfoList)
                {
                    row["競賽成績_競賽層級" + idx] = cpi.reward_level;
                    row["競賽成績_競賽性質" + idx] = cpi.habitude;
                    row["競賽成績_競賽名稱" + idx] = cpi.name;
                    row["競賽成績_得獎名次" + idx] = cpi.rank_name;
                    row["競賽成績_證書日期" + idx] = cpi.certificate_date;
                    row["競賽成績_主辦單位" + idx] = cpi.organizer;
                    idx++;
                }

                // 掃名次填入細項
                foreach (rptCompPerformanceInfo cpi in si.CompPerformanceInfoList)
                {
                    int n = 0;
                    int.TryParse(cpi.bt_rank_int, out n);


                    if (cpi.reward_level.Contains("國"))
                    {
                        // 全國
                        if (cpi.habitude == "個人賽")
                        {
                            if (n >= 1 && n <= 8)
                            {
                                CompPerformanceCountDict["競賽成績_全國個人名次" + n]++;
                            }
                            else
                            {
                                CompPerformanceCountDict["競賽成績_全國個人名次_其他"]++;
                            }
                        }

                        if (cpi.habitude == "團體賽")
                        {
                            if (n >= 1 && n <= 8)
                            {
                                CompPerformanceCountDict["競賽成績_全國團體名次" + n]++;
                            }
                            else
                            {
                                CompPerformanceCountDict["競賽成績_全國團體名次_其他"]++;
                            }
                        }
                    }
                    else
                    {
                        // 市
                        if (cpi.habitude == "個人賽")
                        {
                            if (n >= 1 && n <= 8)
                            {
                                CompPerformanceCountDict["競賽成績_縣市個人名次" + n]++;
                            }
                            else
                            {
                                CompPerformanceCountDict["競賽成績_縣市個人名次_其他"]++;
                            }
                        }

                        if (cpi.habitude == "團體賽")
                        {
                            if (n >= 1 && n <= 8)
                            {
                                CompPerformanceCountDict["競賽成績_縣市團體名次" + n]++;
                            }
                            else
                            {
                                CompPerformanceCountDict["競賽成績_縣市團體名次_其他"]++;
                            }
                        }
                    }
                }

                // 填入值
                foreach (string key in CompPerformanceCountDict.Keys)
                {
                    row[key] = CompPerformanceCountDict[key];
                }

                row["競賽成績_競賽積分"] = si.CompPerformanceScore;

                // 從報表樣板上檢查是由這幾個相加取得總積分：
                // 扶助弱勢_積分
                // 均衡學習_積分
                // 品德表現_獎懲_積分
                // 品德表現_服務學習_積分
                // 品德表現_體適能_積分
                // 競賽成績_競賽積分

                row["成績_合計總分"] = si.IncomeType1Score + si.DomainIScore + si.MDIScore + si.ServiceIScore + si.FitnessIScore + si.CompPerformanceScore;

                dtTable.Rows.Add(row);

                //// debug 
                //dtTable.TableName = "debug";
                //dtTable.WriteXml(Application.StartupPath + @"\debug" + si.StudentID + ".xml");

                docTemplate.MailMerge.Execute(dtTable);
                docTemplate.MailMerge.RemoveEmptyParagraphs = true;
                docTemplate.MailMerge.DeleteFields();

                if (!StudentDocDict.ContainsKey(si.StudentID))
                {
                    StudentDocDict.Add(si.StudentID, docTemplate);
                    StudentDocNameDict.Add(si.StudentID, si.ClassName + "班" + si.SeatNo + "號");
                }

                if (!IsSingleFileSaved)
                    _DocList.Add(docTemplate);
            }

            if (!IsSingleFileSaved)
            {
                _Doc.Sections.Clear();
                foreach (Document doc2 in _DocList)
                    _Doc.Sections.Add(_Doc.ImportNode(doc2.Sections[0], true));
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
            IsSingleFileSaved = ckbSingleFile.Checked;
            bgWorkerReport.RunWorkerAsync();
        }

        private void lnkDefault_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkDefault.Enabled = false;
            string reportName = "嘉義免試入學-超額比序均衡學習成績證明單";

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

        /// <summary>
        /// 合併欄位總表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lnkViewMapColumns_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("超額比序均衡學習成績證明單合併欄位總表產生中...");

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

        /// <summary>
        /// 檢視套印樣板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = "嘉義免試入學-超額比序均衡學習成績證明單";

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
            this.Text = "超額比序均衡學習成績證明單(資料統計至" + _Configure.EndDate.Year + "年" + _Configure.EndDate.Month + "月" + _Configure.EndDate.Day + "日)";
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
            lnkDefault.Enabled = value;
            btnPrint.Enabled = value;
        }
    }
}
