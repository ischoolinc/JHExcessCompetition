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
using Aspose.Cells;
using System.IO;
using FISCA.UDT;
using System.Xml.Linq;
using ChiaYiExcessCompetition.DAO;
using JHSchool.Data;

namespace ChiaYiExcessCompetition
{
    public partial class SubmitForReviewForm : BaseForm
    {
        BackgroundWorker bgWorkerExport = new BackgroundWorker();
        public Configure _Configure { get; private set; }
        AccessHelper _accessHelper = new AccessHelper();
        Dictionary<string, string> StudentCanSelectTagDict = new Dictionary<string, string>();


        public SubmitForReviewForm()
        {
            InitializeComponent();
            bgWorkerExport.DoWork += BgWorkerExport_DoWork;
            bgWorkerExport.RunWorkerCompleted += BgWorkerExport_RunWorkerCompleted;
            bgWorkerExport.ProgressChanged += BgWorkerExport_ProgressChanged;
            bgWorkerExport.WorkerReportsProgress = true;

        }

        private void BgWorkerExport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生中 ... ", e.ProgressPercentage);
        }

        private void BgWorkerExport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                try
                {
                    Workbook wb = (Workbook)e.Result;
                    if (wb != null)
                    {
                        string reportName = "嘉義免試入學-送審用匯入檔";

                        string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                        if (!Directory.Exists(path))
                            Directory.CreateDirectory(path);
                        path = Path.Combine(path, reportName + ".xlsx");

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

                        Workbook wbNew = new Workbook();
                        try
                        {

                            wbNew = wb;
                            wbNew.Save(path, SaveFormat.Xlsx);
                            System.Diagnostics.Process.Start(path);
                        }
                        catch
                        {
                            System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                            sd.Title = "另存新檔";
                            sd.FileName = reportName + ".xlsx";
                            sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
                            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                try
                                {
                                    wbNew.Save(sd.FileName, SaveFormat.Xlsx);

                                }
                                catch
                                {
                                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                    return;
                                }
                            }
                        }
                    }

                    UserControlEnable(true);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("發生錯誤：" + ex.Message);
                }
            }
            else
            {
                MsgBox.Show("發生錯誤：" + e.Error.Message);
            }
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void BgWorkerExport_DoWork(object sender, DoWorkEventArgs e)
        {
            bgWorkerExport.ReportProgress(1);
            try
            {
                // 取得預設樣板
                Workbook wb = new Workbook(new MemoryStream(Properties.Resources.Template));
                Worksheet wst = wb.Worksheets[0];


                // 取得學生基本資料
                List<StudentInfo> StudentInfoList = QueryData.GetStudentInfoList3();

                List<string> StudentIDList = new List<string>();
                foreach (StudentInfo si in StudentInfoList)
                {
                    StudentIDList.Add(si.StudentID);
                }

                // 取得地址資訊
                Dictionary<string, JHAddressRecord> AddressDict = new Dictionary<string, JHAddressRecord>();
                List<JHAddressRecord> tmpAddress = JHAddress.SelectByStudentIDs(StudentIDList);
                foreach (JHAddressRecord rec in tmpAddress)
                {
                    if (!AddressDict.ContainsKey(rec.RefStudentID))
                        AddressDict.Add(rec.RefStudentID, rec);
                }

                // 取得電話資料
                Dictionary<string, JHPhoneRecord> PhoneDict = new Dictionary<string, JHPhoneRecord>();
                List<JHPhoneRecord> tmpPhone = JHPhone.SelectByStudentIDs(StudentIDList);
                foreach (JHPhoneRecord rec in tmpPhone)
                {
                    if (!PhoneDict.ContainsKey(rec.RefStudentID))
                        PhoneDict.Add(rec.RefStudentID, rec);
                }

                bgWorkerExport.ReportProgress(20);

                // 取得監護人父母資訊
                Dictionary<string, JHParentRecord> ParentDict = new Dictionary<string, JHParentRecord>();
                List<JHParentRecord> tmpParent = JHParent.SelectByStudentIDs(StudentIDList);
                foreach (JHParentRecord rec in tmpParent)
                {
                    if (!ParentDict.ContainsKey(rec.RefStudentID))
                        ParentDict.Add(rec.RefStudentID, rec);
                }

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



                bgWorkerExport.ReportProgress(40);

                // 取得成績相關資料
                Dictionary<string, List<JHSemesterScoreRecord>> SemesterScoreRecordDict = new Dictionary<string, List<JHSemesterScoreRecord>>();
                List<JHSemesterScoreRecord> tmpSemsScore = JHSemesterScore.SelectByStudentIDs(StudentIDList);
                foreach (JHSemesterScoreRecord rec in tmpSemsScore)
                {
                    if (!SemesterScoreRecordDict.ContainsKey(rec.RefStudentID))
                        SemesterScoreRecordDict.Add(rec.RefStudentID, new List<JHSemesterScoreRecord>());

                    SemesterScoreRecordDict[rec.RefStudentID].Add(rec);
                }

                // 取得功過紀錄
                // 功過對照表
                JHMeritDemeritReduceRecord DemeritReduceRecord = JHMeritDemeritReduce.Select();
                // 懲
                Dictionary<string, List<JHDemeritRecord>> DemeritRecordDict = new Dictionary<string, List<JHDemeritRecord>>();
                List<JHDemeritRecord> tmpDemeritRecord = JHDemerit.SelectByStudentIDs(StudentIDList);
                foreach (JHDemeritRecord rec in tmpDemeritRecord)
                {
                    if (rec.Cleared == "是")
                        continue;

                    if (rec.OccurDate > _Configure.EndDate)
                        continue;
                    else
                    {
                        if (!DemeritRecordDict.ContainsKey(rec.RefStudentID))
                            DemeritRecordDict.Add(rec.RefStudentID, new List<JHDemeritRecord>());

                        DemeritRecordDict[rec.RefStudentID].Add(rec);
                    }

                }
                // 獎
                Dictionary<string, List<JHMeritRecord>> MeritRecordDict = new Dictionary<string, List<JHMeritRecord>>();
                List<JHMeritRecord> tmpMeritRecord = JHMerit.SelectByStudentIDs(StudentIDList);
                foreach (JHMeritRecord rec in tmpMeritRecord)
                {
                    if (rec.OccurDate > _Configure.EndDate)
                        continue;
                    else
                    {
                        if (!MeritRecordDict.ContainsKey(rec.RefStudentID))
                            MeritRecordDict.Add(rec.RefStudentID, new List<JHMeritRecord>());

                        MeritRecordDict[rec.RefStudentID].Add(rec);
                    }

                }

                // 取得服務學習時數
                StudentInfoList = QueryData.FillServiceLearn(StudentIDList, StudentInfoList,_Configure.EndDate);

                // 填入中低收入戶
                StudentInfoList = QueryData.FillIncomeType(StudentIDList, StudentInfoList);


                // 取得學生體適能資料並填入,嘉義版不卡日期，日期傳入不會限制
                StudentInfoList = QueryData.FillStudentFitness(StudentIDList, StudentInfoList, _Configure.EndDate);

                // 取得競賽總積分並填入學生資料
                StudentInfoList = QueryData.FillStudentCompetitionPerformanceSum(StudentIDList, StudentInfoList);

                // 填入 Excel 資料
                int wstRIdx = 1;
                bgWorkerExport.ReportProgress(70);

                foreach (StudentInfo si in StudentInfoList)
                {

                    // 考區代碼 0,嘉義區 10
                    wst.Cells[wstRIdx, 0].PutValue(10);

                    // 集報單位代碼 1
                    wst.Cells[wstRIdx, 1].PutValue(K12.Data.School.Code);

                    // 序號 2
                    wst.Cells[wstRIdx, 2].PutValue(wstRIdx);

                    // 學號 3
                    wst.Cells[wstRIdx, 3].PutValue(si.StudentNumber);

                    // 班級 4
                    wst.Cells[wstRIdx, 4].PutValue(si.ClassName);

                    // 座號 5
                    wst.Cells[wstRIdx, 5].PutValue(si.SeatNo);

                    // 學生姓名 6
                    wst.Cells[wstRIdx, 6].PutValue(si.StudentName);

                    // 身分證統一編號 7
                    wst.Cells[wstRIdx, 7].PutValue(si.IDNumber);

                    // 非中華民國身分證號 8
                    if (si.isTaiwanID)
                    {
                        wst.Cells[wstRIdx, 8].PutValue("");
                    }
                    else
                    {
                        wst.Cells[wstRIdx, 8].PutValue("V");
                    }

                    // 性別 9
                    wst.Cells[wstRIdx, 9].PutValue(si.GenderCode);

                    // 出生年(民國年) 10
                    wst.Cells[wstRIdx, 10].PutValue(si.BirthYear);

                    // 出生月 11
                    wst.Cells[wstRIdx, 11].PutValue(si.BirthMonth);

                    // 出生日 12
                    wst.Cells[wstRIdx, 12].PutValue(si.BirthDay);

                    // 畢業學校代碼 13
                    wst.Cells[wstRIdx, 13].PutValue(K12.Data.School.Code);

                    // 畢業年(民國年) 14
                    int gyear;
                    if (int.TryParse(K12.Data.School.DefaultSchoolYear, out gyear))
                    {
                        wst.Cells[wstRIdx, 14].PutValue(gyear + 1);
                    }

                    // 畢肄業 15
                    wst.Cells[wstRIdx, 15].PutValue(1);


                    wst.Cells[wstRIdx, 16].PutValue(0);
                    wst.Cells[wstRIdx, 17].PutValue(0);
                    wst.Cells[wstRIdx, 18].PutValue(0);
                    wst.Cells[wstRIdx, 22].PutValue(0);

                    if (StudentTagDict.ContainsKey(si.StudentID))
                    {
                        foreach (string tagName in StudentTagDict[si.StudentID])
                        {
                            if (MappingTag1.ContainsKey(tagName))
                            {
                                // 學生身分 16                              
                                wst.Cells[wstRIdx, 16].PutValue(MappingTag1[tagName]);
                            }

                            if (MappingTag2.ContainsKey(tagName))
                            {
                                si.isSpecial = true;
                                // 身心障礙 18                               
                                wst.Cells[wstRIdx, 18].PutValue(MappingTag2[tagName]);
                            }

                            if (MappingTag3.ContainsKey(tagName))
                            {
                                // 學生報名身分 17                               
                                wst.Cells[wstRIdx, 17].PutValue(MappingTag3[tagName]);

                            }
                            if (MappingTag4.ContainsKey(tagName))
                            {
                                // 失業勞工子女 22                              
                                wst.Cells[wstRIdx, 22].PutValue(MappingTag4[tagName]);
                            }
                        }

                    }


                    // 就學區 19,不處理

                    // 低收入戶 20
                    if (si.incomeType1)
                        wst.Cells[wstRIdx, 20].PutValue(1);
                    else
                        wst.Cells[wstRIdx, 20].PutValue(0);

                    // 中低收入戶 21
                    if (si.incomeType2)
                        wst.Cells[wstRIdx, 21].PutValue(1);
                    else
                        wst.Cells[wstRIdx, 21].PutValue(0);



                    // 資料授權 23
                    wst.Cells[wstRIdx, 23].PutValue(0);

                    string parentName = "";
                    // 家長姓名 24
                    if (ParentDict.ContainsKey(si.StudentID))
                    {

                        if (!string.IsNullOrWhiteSpace(ParentDict[si.StudentID].CustodianName))
                        {
                            parentName = ParentDict[si.StudentID].CustodianName;
                        }
                        else if (!string.IsNullOrWhiteSpace(ParentDict[si.StudentID].FatherName))
                        {
                            parentName = ParentDict[si.StudentID].FatherName;
                        }
                        else if (!string.IsNullOrWhiteSpace(ParentDict[si.StudentID].MotherName))
                        {
                            parentName = ParentDict[si.StudentID].MotherName;
                        }
                        else
                        {

                        }
                        wst.Cells[wstRIdx, 24].PutValue(parentName);
                    }


                    // 市內電話 25
                    // 市內電話分機 26
                    // 行動電話 27
                    if (PhoneDict.ContainsKey(si.StudentID))
                    {
                        wst.Cells[wstRIdx, 25].PutValue(PhoneDict[si.StudentID].Contact.Replace("-", "").Replace(")", "").Replace("(", ""));
                        wst.Cells[wstRIdx, 27].PutValue(PhoneDict[si.StudentID].Cell.Replace("-", "").Replace(")", "").Replace("(", ""));
                    }


                    // 郵遞區號 28
                    if (AddressDict.ContainsKey(si.StudentID))
                    {

                        if (AddressDict[si.StudentID].MailingZipCode != null)
                        {
                            string zipCode = AddressDict[si.StudentID].MailingZipCode;

                            if (zipCode.Length >= 3)
                                zipCode = zipCode.Substring(0, 3);

                            wst.Cells[wstRIdx, 28].PutValue(zipCode);
                        }

                        // 通訊地址 29
                        wst.Cells[wstRIdx, 29].PutValue(AddressDict[si.StudentID].MailingCounty + AddressDict[si.StudentID].MailingTown + AddressDict[si.StudentID].MailingDistrict + AddressDict[si.StudentID].MailingArea + AddressDict[si.StudentID].MailingDetail);

                    }

                    // 計算分數
                    if (SemesterScoreRecordDict.ContainsKey(si.StudentID))
                    {
                        si.CalcSemsScore5(SemesterScoreRecordDict[si.StudentID]);                      
                    }

                    // 健康與體育 30
                    if (si.isDomainHelPass)
                        wst.Cells[wstRIdx, 30].PutValue(1);
                    else
                        wst.Cells[wstRIdx, 30].PutValue(0);

                    // 藝術與人文 31
                    if (si.isDoaminArtPass)
                        wst.Cells[wstRIdx, 31].PutValue(1);
                    else
                        wst.Cells[wstRIdx, 31].PutValue(0);

                    // 綜合活動 32
                    if (si.isDomainActPass)
                        wst.Cells[wstRIdx, 32].PutValue(1);
                    else
                        wst.Cells[wstRIdx, 32].PutValue(0);

                    // 品德表現 33
                    List<JHDemeritRecord> recD;
                    List<JHMeritRecord> recM;

                    if (DemeritRecordDict.ContainsKey(si.StudentID))
                        recD = DemeritRecordDict[si.StudentID];
                    else
                        recD = new List<JHDemeritRecord>();

                    if (MeritRecordDict.ContainsKey(si.StudentID))
                        recM = MeritRecordDict[si.StudentID];
                    else
                        recM = new List<JHMeritRecord>();

                    si.CalcDemeritMemeritScore(recD, recM, DemeritReduceRecord);
                    wst.Cells[wstRIdx, 33].PutValue(si.MeritDemeritScore);

                    // 服務學習 34
                    wst.Cells[wstRIdx, 34].PutValue(si.ServiceLearnScore);

                    // 體適能 35
                    si.CalcFitnessScore();
                    wst.Cells[wstRIdx, 35].PutValue(si.FitnessScore);

                    // 競賽表現 36,使用者自行處理
                    if (si.CompPerfSum.HasValue)
                    {
                        wst.Cells[wstRIdx, 36].PutValue(si.CompPerfSum.Value);
                    }
              

                    // 不處理
                    // 家長意見_高中 37
                    // 家長意見_高職 38
                    // 導師意見_高中 39
                    // 導師意見_高職 40
                    // 輔導教師意見_高中 41
                    // 輔導教師意見_高職 42


                    wstRIdx++;
                }

                bgWorkerExport.ReportProgress(100);

                e.Result = wb;


            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            // 取得畫面上日期
            _Configure.EndDate = dtDate.Value;
            if (SaveConfig())
            {
                UserControlEnable(false);
                bgWorkerExport.RunWorkerAsync();
            }
        }


        private void SubmitForReviewForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;

            UserControlEnable(false);

            LoadTemplate();
            if (_Configure.EndDate.Date.Year < DateTime.Now.Year)
                _Configure.EndDate = new DateTime(DateTime.Now.Year, 4, 30);
            dtDate.Value = _Configure.EndDate;

            LoadConfig();
            LoadStudentTag();
            LoadDataGridViewValue();
            UserControlEnable(true);
        }



        private void btnSave_Click(object sender, EventArgs e)
        {
            UserControlEnable(false);
            if (SaveConfig())
            {
                MsgBox.Show("儲存完成。");
            }
            UserControlEnable(true);
        }


        private void LoadConfig()
        {
            dgData1.Rows.Clear();
            dgData2.Rows.Clear();
            dgData3.Rows.Clear();

            if (_Configure != null)
            {
                try
                {
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
                                    int rowIdx = dgData1.Rows.Add();
                                    dgData1.Rows[rowIdx].Cells[colCode1.Index].Value = elm1.Attribute("Code").Value;
                                    dgData1.Rows[rowIdx].Cells[colItem1.Index].Value = elm1.Attribute("Name").Value;
                                }

                            }

                            if (gpName == "身心障礙")
                            {
                                foreach (XElement elm1 in elm.Elements("Item"))
                                {
                                    int rowIdx = dgData2.Rows.Add();
                                    dgData2.Rows[rowIdx].Cells[colCode1.Index].Value = elm1.Attribute("Code").Value;
                                    dgData2.Rows[rowIdx].Cells[colItem1.Index].Value = elm1.Attribute("Name").Value;
                                }
                            }

                            if (gpName == "學生報名身分設定")
                            {
                                foreach (XElement elm1 in elm.Elements("Item"))
                                {
                                    int rowIdx = dgData3.Rows.Add();
                                    dgData3.Rows[rowIdx].Cells[colCode1.Index].Value = elm1.Attribute("Code").Value;
                                    dgData3.Rows[rowIdx].Cells[colItem1.Index].Value = elm1.Attribute("Name").Value;
                                }
                            }


                            cboSelectTag4.Text = "";

                        }
                    }
                }
                catch (Exception ex)
                {
                    MsgBox.Show("解析設定檔失敗," + ex.Message);
                }
            }

        }


        /// <summary>
        /// 載入學生可選類別
        /// </summary>
        private void LoadStudentTag()
        {
            StudentCanSelectTagDict = QueryData.GetStudentAllTag();

            DataTable dd1 = new DataTable();
            DataTable dd2 = new DataTable();
            DataTable dd3 = new DataTable();
            DataTable dd4 = new DataTable();

            cboSelectTag4.Items.Clear();
            DataGridViewComboBoxColumn cboItem1 = new DataGridViewComboBoxColumn();
            cboItem1.Name = "colStudTag1";
            cboItem1.Width = 150;
            cboItem1.HeaderText = "學生類別";

            DataGridViewComboBoxColumn cboItem2 = new DataGridViewComboBoxColumn();
            cboItem2.Name = "colStudTag2";
            cboItem2.Width = 150;
            cboItem2.HeaderText = "學生類別";

            DataGridViewComboBoxColumn cboItem3 = new DataGridViewComboBoxColumn();
            cboItem3.Name = "colStudTag3";
            cboItem3.Width = 150;
            cboItem3.HeaderText = "學生類別";


            cboSelectTag4.Items.Add("");
            dd1.Columns.Add("VALUE");
            dd1.Columns.Add("ITEM");
            dd2.Columns.Add("VALUE");
            dd2.Columns.Add("ITEM");
            dd3.Columns.Add("VALUE");
            dd3.Columns.Add("ITEM");
            dd4.Columns.Add("VALUE");
            dd4.Columns.Add("ITEM");

            List<string> selectItems = new List<string>();
            selectItems.Add("");

            foreach (string name in StudentCanSelectTagDict.Keys)
            {
                selectItems.Add(name);
            }


            foreach (string name in selectItems)
            {
                DataRow dr1 = dd1.NewRow();
                dr1["VALUE"] = name;
                dr1["ITEM"] = name;
                dd1.Rows.Add(dr1);

                DataRow dr2 = dd2.NewRow();
                dr2["VALUE"] = name;
                dr2["ITEM"] = name;
                dd2.Rows.Add(dr2);

                DataRow dr3 = dd3.NewRow();
                dr3["VALUE"] = name;
                dr3["ITEM"] = name;
                dd3.Rows.Add(dr3);

                DataRow dr4 = dd4.NewRow();
                dr4["VALUE"] = name;
                dr4["ITEM"] = name;
                dd4.Rows.Add(dr4);
            }

            cboItem1.DataSource = dd1;
            cboItem1.DisplayMember = "ITEM";
            cboItem1.ValueMember = "VALUE";

            cboItem2.DataSource = dd2;
            cboItem2.DisplayMember = "ITEM";
            cboItem2.ValueMember = "VALUE";

            cboItem3.DataSource = dd3;
            cboItem3.DisplayMember = "ITEM";
            cboItem3.ValueMember = "VALUE";

            cboSelectTag4.DataSource = dd4;

            cboSelectTag4.DisplayMember = "ITEM";
            cboSelectTag4.ValueMember = "VALUE";

            dgData1.Columns.Add(cboItem1);
            dgData2.Columns.Add(cboItem2);
            dgData3.Columns.Add(cboItem3);
        }

        private void LoadDataGridViewValue()
        {
            if (_Configure != null)
            {
                try
                {
                    XElement elmRoot = XElement.Parse(_Configure.MappingContent);
                    if (elmRoot != null)
                    {
                        foreach (XElement elm in elmRoot.Elements("Group"))
                        {
                            string gpName = elm.Attribute("Name").Value;
                            if (gpName == "學生身分")
                            {
                                foreach (DataGridViewRow drv in dgData1.Rows)
                                {
                                    foreach (XElement elm1 in elm.Elements("Item"))
                                    {
                                        if (drv.Cells[colItem1.Index].Value.ToString() == elm1.Attribute("Name").Value)
                                        {
                                            if (StudentCanSelectTagDict.ContainsKey(elm1.Attribute("TagName").Value))
                                                drv.Cells[2].Value = elm1.Attribute("TagName").Value;
                                            else
                                                drv.Cells[2].Value = "";
                                            break;
                                        }
                                    }
                                }

                            }

                            if (gpName == "身心障礙")
                            {
                                foreach (DataGridViewRow drv in dgData2.Rows)
                                {
                                    foreach (XElement elm1 in elm.Elements("Item"))
                                    {
                                        if (drv.Cells[colItem1.Index].Value.ToString() == elm1.Attribute("Name").Value)
                                        {
                                            if (StudentCanSelectTagDict.ContainsKey(elm1.Attribute("TagName").Value))
                                                drv.Cells[2].Value = elm1.Attribute("TagName").Value;
                                            else
                                                drv.Cells[2].Value = "";
                                            break;
                                        }
                                    }
                                }

                            }

                            if (gpName == "學生報名身分設定")
                            {
                                foreach (DataGridViewRow drv in dgData3.Rows)
                                {
                                    foreach (XElement elm1 in elm.Elements("Item"))
                                    {
                                        if (drv.Cells[colItem1.Index].Value.ToString() == elm1.Attribute("Name").Value)
                                        {
                                            if (StudentCanSelectTagDict.ContainsKey(elm1.Attribute("TagName").Value))
                                                drv.Cells[2].Value = elm1.Attribute("TagName").Value;
                                            else
                                                drv.Cells[2].Value = "";
                                            break;
                                        }
                                    }
                                }
                            }

                            if (gpName == "失業勞工子女")
                            {
                                foreach (XElement elm1 in elm.Elements("Item"))
                                {
                                    if (StudentCanSelectTagDict.ContainsKey(elm1.Attribute("TagName").Value))
                                        cboSelectTag4.Text = elm1.Attribute("TagName").Value;
                                    else
                                        cboSelectTag4.Text = "";
                                    break;
                                }
                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    MsgBox.Show("解析設定檔失敗," + ex.Message);
                }
            }
        }


        /// <summary>
        /// 儲存設定檔
        /// </summary>
        private bool SaveConfig()
        {
            try
            {
                // 收集資料回存
                XElement elmRoot = XElement.Parse(_Configure.MappingContent);
                if (elmRoot != null)
                {
                    foreach (XElement elm in elmRoot.Elements("Group"))
                    {
                        string gpName = elm.Attribute("Name").Value;
                        if (gpName == "學生身分")
                        {

                            foreach (DataGridViewRow drv in dgData1.Rows)
                            {
                                if (drv.IsNewRow)
                                    continue;

                                foreach (XElement elm1 in elm.Elements("Item"))
                                {
                                    string name = elm1.Attribute("Name").Value;
                                    if (name == drv.Cells[colItem1.Index].Value.ToString())
                                    {
                                        if (drv.Cells[2].Value == null)
                                        {
                                            elm1.SetAttributeValue("TagName", "");
                                        }
                                        else
                                        {
                                            elm1.SetAttributeValue("TagName", drv.Cells[2].Value.ToString());
                                        }
                                        break;
                                    }
                                }
                            }
                        }

                        if (gpName == "身心障礙")
                        {
                            foreach (DataGridViewRow drv in dgData2.Rows)
                            {
                                if (drv.IsNewRow)
                                    continue;
                                foreach (XElement elm1 in elm.Elements("Item"))
                                {
                                    string name = elm1.Attribute("Name").Value;
                                    if (name == drv.Cells[colItem1.Index].Value.ToString())
                                    {
                                        if (drv.Cells[2].Value == null)
                                        {
                                            elm1.SetAttributeValue("TagName", "");
                                        }
                                        else
                                        {
                                            elm1.SetAttributeValue("TagName", drv.Cells[2].Value.ToString());
                                        }
                                        break;
                                    }
                                }
                            }
                        }

                        if (gpName == "學生報名身分設定")
                        {
                            foreach (DataGridViewRow drv in dgData3.Rows)
                            {
                                if (drv.IsNewRow)
                                    continue;
                                foreach (XElement elm1 in elm.Elements("Item"))
                                {
                                    string name = elm1.Attribute("Name").Value;
                                    if (name == drv.Cells[colItem1.Index].Value.ToString())
                                    {
                                        if (drv.Cells[2].Value == null)
                                        {
                                            elm1.SetAttributeValue("TagName", "");
                                        }
                                        else
                                        {
                                            elm1.SetAttributeValue("TagName", drv.Cells[2].Value.ToString());
                                        }
                                        break;
                                    }
                                }
                            }
                        }

                        if (gpName == "失業勞工子女")
                        {
                            foreach (XElement elm1 in elm.Elements("Item"))
                            {
                                if (cboSelectTag4.Text == null)
                                {
                                    elm1.SetAttributeValue("TagName", "");
                                }
                                else
                                {
                                    elm1.SetAttributeValue("TagName", cboSelectTag4.Text);
                                }
                                break;
                            }
                        }
                    }
                }
                _Configure.MappingContent = elmRoot.ToString();
                _Configure.Save();

                return true;

            }
            catch (Exception ex)
            {
                MsgBox.Show("寫入設定失敗," + ex.Message);
                return false;
            }
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
                    if (_Configure.EndDate == null)
                        _Configure.EndDate = new DateTime(DateTime.Now.Year, 4, 30);
                    else
                    {
                        if (_Configure.EndDate.Date.Year < DateTime.Now.Year)
                            _Configure.EndDate = new DateTime(DateTime.Now.Year, 4, 30);
                    }

                    if (string.IsNullOrEmpty(_Configure.MappingContent))
                    {
                        _Configure.MappingContent = DefaultConfigXML().ToString();
                    }
                }
                else
                {
                    _Configure = new Configure();
                    _Configure.Name = "嘉義免試入學-班級服務表現";
                    if (_Configure.EndDate == null)
                        _Configure.EndDate = new DateTime(DateTime.Now.Year, 4, 30);

                    if (string.IsNullOrEmpty(_Configure.MappingContent))
                    {
                        _Configure.MappingContent = DefaultConfigXML().ToString();
                    }
                }

                _Configure.Save();

                UserControlEnable(true);
            }
            catch (Exception ex)
            {
                MsgBox.Show("讀取樣板發生錯誤，" + ex.Message);
            }

        }


        /// <summary>
        /// 預設對照XML
        /// </summary>
        /// <returns></returns>
        private XElement DefaultConfigXML()
        {
            XElement elmRoot = new XElement("MappingConfig");
            List<string> gList = new List<string>();
            gList.Add("學生身分");
            gList.Add("身心障礙");
            gList.Add("學生報名身分設定");
            gList.Add("失業勞工子女");

            // 0 一般生
            Dictionary<string, string> g1ItemList = new Dictionary<string, string>();
            g1ItemList.Add("原住民", "1");
            g1ItemList.Add("派外人員子女", "2");
            g1ItemList.Add("蒙藏生", "3");
            g1ItemList.Add("回國僑生", "4");
            g1ItemList.Add("港澳生", "5");
            g1ItemList.Add("退伍軍人", "6");
            g1ItemList.Add("境外優秀科學技術人才子女", "7");

            // 0	無
            Dictionary<string, string> g2ItemList = new Dictionary<string, string>();
            g2ItemList.Add("智能障礙", "1");
            g2ItemList.Add("視覺障礙", "2");
            g2ItemList.Add("聽覺障礙", "3");
            g2ItemList.Add("語言障礙", "4");
            g2ItemList.Add("肢體障礙", "5");
            g2ItemList.Add("腦性麻痺", "6");
            g2ItemList.Add("身體病弱", "7");
            g2ItemList.Add("情緒行為障礙", "8");
            g2ItemList.Add("學習障礙", "9");
            g2ItemList.Add("多重障礙", "A");
            g2ItemList.Add("自閉症", "B");
            g2ItemList.Add("發展遲緩", "C");
            g2ItemList.Add("其他障礙", "D");

            // 0 一般生
            Dictionary<string, string> g3ItemList = new Dictionary<string, string>();
            g3ItemList.Add("身障生", "1");
            g3ItemList.Add("原住民(有認證)", "2");
            g3ItemList.Add("原住民(無認證)", "3");
            g3ItemList.Add("蒙藏生", "4");
            g3ItemList.Add("外派子女25%", "5");
            g3ItemList.Add("外派子女15%", "6");
            g3ItemList.Add("外派子女10%", "7");
            g3ItemList.Add("退伍軍人25%", "8");
            g3ItemList.Add("退伍軍人20%", "9");
            g3ItemList.Add("退伍軍人15%", "A");
            g3ItemList.Add("退伍軍人10%", "B");
            g3ItemList.Add("退伍軍人5%", "C");
            g3ItemList.Add("退伍軍人3%", "D");
            g3ItemList.Add("優秀子女25%", "E");
            g3ItemList.Add("優秀子女15%", "F");
            g3ItemList.Add("優秀子女10%", "G");
            g3ItemList.Add("僑生", "H");


            foreach (string gname in gList)
            {
                XElement g1 = new XElement("Group");
                g1.SetAttributeValue("Name", gname);

                if (gname == "學生身分")
                {
                    foreach (string na in g1ItemList.Keys)
                    {
                        XElement elm = new XElement("Item");
                        elm.SetAttributeValue("Name", na);
                        elm.SetAttributeValue("TagName", "");
                        elm.SetAttributeValue("Code", g1ItemList[na]);
                        g1.Add(elm);
                    }
                }

                if (gname == "身心障礙")
                {
                    foreach (string na in g2ItemList.Keys)
                    {
                        XElement elm = new XElement("Item");
                        elm.SetAttributeValue("Name", na);
                        elm.SetAttributeValue("TagName", "");
                        elm.SetAttributeValue("Code", g2ItemList[na]);
                        g1.Add(elm);
                    }
                }
                if (gname == "學生報名身分設定")
                {
                    foreach (string na in g3ItemList.Keys)
                    {
                        XElement elm = new XElement("Item");
                        elm.SetAttributeValue("Name", na);
                        elm.SetAttributeValue("TagName", "");
                        elm.SetAttributeValue("Code", g3ItemList[na]);
                        g1.Add(elm);
                    }
                }

                if (gname == "失業勞工子女")
                {
                    XElement elm = new XElement("Item");
                    elm.SetAttributeValue("Name", "是");
                    elm.SetAttributeValue("TagName", "");
                    elm.SetAttributeValue("Code", "1");
                    g1.Add(elm);
                }

                elmRoot.Add(g1);
            }



            return elmRoot;
        }

        private void UserControlEnable(bool value)
        {
            dgData1.Enabled = dgData2.Enabled = dgData3.Enabled = cboSelectTag4.Enabled = dtDate.Enabled = btnSave.Enabled = btnExport.Enabled = value;
        }
    }
}
