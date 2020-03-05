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
using Aspose.Cells;
using System.IO;
using FISCA.UDT;
using System.Xml.Linq;
using PingTungExcessCompetition.DAO;
using JHSchool.Data;

namespace PingTungExcessCompetition
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
                        string reportName = "屏東免試入學-送審用匯入檔";

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


                // 取得語言認證學生id
                List<string> hasLanguageCertificateIDList = QueryData.GetLanguageCertificate(StudentIDList);

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

                    if (rec.RegisterDate.HasValue)
                    {
                        if (rec.RegisterDate.Value > _Configure.EndDate)
                            continue;
                        else
                        {
                            if (!DemeritRecordDict.ContainsKey(rec.RefStudentID))
                                DemeritRecordDict.Add(rec.RefStudentID, new List<JHDemeritRecord>());

                            DemeritRecordDict[rec.RefStudentID].Add(rec);
                        }
                    }
                }
                // 獎
                Dictionary<string, List<JHMeritRecord>> MeritRecordDict = new Dictionary<string, List<JHMeritRecord>>();
                List<JHMeritRecord> tmpMeritRecord = JHMerit.SelectByStudentIDs(StudentIDList);
                foreach (JHMeritRecord rec in tmpMeritRecord)
                {
                    if (rec.RegisterDate.HasValue)
                    {
                        if (rec.RegisterDate.Value > _Configure.EndDate)
                            continue;
                        else
                        {
                            if (!MeritRecordDict.ContainsKey(rec.RefStudentID))
                                MeritRecordDict.Add(rec.RefStudentID, new List<JHMeritRecord>());

                            MeritRecordDict[rec.RefStudentID].Add(rec);
                        }
                    }
                }
                // 填入 Excel 資料
                int wstRIdx = 1;

                int d18 = 0, d19 = 0, d31 = 0, d32 = 0, d33 = 0, d34 = 0;



                bgWorkerExport.ReportProgress(70);

                // 幹部限制
                List<string> CadreName1 = Global.GetCadreName1();

                foreach (StudentInfo si in StudentInfoList)
                {

                    // 考區代碼 0， 12/屏東考區
                    wst.Cells[wstRIdx, 0].PutValue(12);
                    // 集報單位代碼 1,學校代碼
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

                    // 性別 8
                    wst.Cells[wstRIdx, 8].PutValue(si.GenderCode);

                    // 出生年(民國年) 9
                    wst.Cells[wstRIdx, 9].PutValue(si.BirthYear);
                    // 出生月 10
                    wst.Cells[wstRIdx, 10].PutValue(si.BirthMonth);
                    // 出生日 11
                    wst.Cells[wstRIdx, 11].PutValue(si.BirthDay);
                    // 畢業學校代碼 12
                    wst.Cells[wstRIdx, 12].PutValue(K12.Data.School.Code);

                    // 畢業年(民國年) 13
                    wst.Cells[wstRIdx, 13].PutValue(K12.Data.School.DefaultSchoolYear);

                    // 畢肄業 14
                    wst.Cells[wstRIdx, 14].PutValue(1);


                    // 就學區 17

                    // 低收入戶 18
                    d18 = 0;
                    wst.Cells[wstRIdx, 18].PutValue(d18);

                    // 中低收入戶 19
                    d19 = 0;
                    wst.Cells[wstRIdx, 19].PutValue(d19);

                    wst.Cells[wstRIdx, 15].PutValue("0");
                    wst.Cells[wstRIdx, 16].PutValue("0");
                    wst.Cells[wstRIdx, 20].PutValue("0");
                    wst.Cells[wstRIdx, 28].PutValue("0");

                    if (StudentTagDict.ContainsKey(si.StudentID))
                    {
                        foreach (string tagName in StudentTagDict[si.StudentID])
                        {
                            if (MappingTag1.ContainsKey(tagName))
                            {
                                // 學生身分 15                                
                                wst.Cells[wstRIdx, 15].PutValue(MappingTag1[tagName]);
                            }

                            if (MappingTag2.ContainsKey(tagName))
                            {
                                // 身心障礙 16                               
                                wst.Cells[wstRIdx, 16].PutValue(MappingTag2[tagName]);
                            }

                            if (MappingTag3.ContainsKey(tagName))
                            {
                                // 學生報名身分 28                               
                                wst.Cells[wstRIdx, 28].PutValue(MappingTag3[tagName]);

                            }
                            if (MappingTag4.ContainsKey(tagName))
                            {
                                // 失業勞工子女 20                              
                                wst.Cells[wstRIdx, 20].PutValue(MappingTag4[tagName]);
                            }
                        }

                    }



                    // 資料授權 21
                    wst.Cells[wstRIdx, 21].PutValue(0);

                    string parentName = "";


                    // 家長姓名 22
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
                        wst.Cells[wstRIdx, 22].PutValue(parentName);
                    }


                    // 市內電話 23
                    // 行動電話 24
                    if (PhoneDict.ContainsKey(si.StudentID))
                    {
                        wst.Cells[wstRIdx, 23].PutValue(PhoneDict[si.StudentID].Contact);
                        wst.Cells[wstRIdx, 24].PutValue(PhoneDict[si.StudentID].Cell);
                    }


                    // 郵遞區號 25
                    if (AddressDict.ContainsKey(si.StudentID))
                    {
                        wst.Cells[wstRIdx, 25].PutValue(AddressDict[si.StudentID].MailingZipCode);

                        // 通訊地址 26
                        wst.Cells[wstRIdx, 26].PutValue(AddressDict[si.StudentID].MailingCounty + AddressDict[si.StudentID].MailingTown + AddressDict[si.StudentID].MailingDistrict + AddressDict[si.StudentID].MailingArea + AddressDict[si.StudentID].MailingDetail);

                    }
                    // 非中華民國身分證號 27



                    // 市內電話分機 29

                    // 均衡學習 30
                    // 計算分數
                    if (SemesterScoreRecordDict.ContainsKey(si.StudentID))
                    {
                        si.CalcSemsScore5(SemesterScoreRecordDict[si.StudentID]);
                        // 成績滿5學期才顯示
                        if (si.hasSemester5Score)
                            wst.Cells[wstRIdx, 30].PutValue(si.Semester5Score);
                    }

                    // 服務表現 31
                    si.CalcCadreScore(CadreName1);
                    wst.Cells[wstRIdx, 31].PutValue(si.ServiceScore);

                    // 品德表現 32
                    if (DemeritRecordDict.ContainsKey(si.StudentID))
                    {
                        if (MeritRecordDict.ContainsKey(si.StudentID))
                        {
                            si.CalcDemeritMemeritScore(DemeritRecordDict[si.StudentID], MeritRecordDict[si.StudentID], DemeritReduceRecord);
                        }else
                        {
                            si.CalcDemeritMemeritScore(DemeritRecordDict[si.StudentID], new List<JHMeritRecord>(), DemeritReduceRecord);
                        }
                        wst.Cells[wstRIdx, 32].PutValue(si.MeritDemeritScore);
                    }
                    else
                    {
                        // 沒有懲戒
                        wst.Cells[wstRIdx, 32].PutValue(10);
                    }
                   

                    // 競賽表現 33
                    d33 = 0;
                    wst.Cells[wstRIdx, 33].PutValue(d33);

                    // 體適能 34
                    d34 = 0;
                    wst.Cells[wstRIdx, 34].PutValue(d34);

                    // 本土語言認證 35
                    if (hasLanguageCertificateIDList.Contains(si.StudentID))
                    {
                        wst.Cells[wstRIdx, 35].PutValue(1);
                    }
                    else
                    {
                        wst.Cells[wstRIdx, 35].PutValue(0);
                    }



                    // 36~39 系統無法提供先空
                    // 適性發展_高中 36
                    // 適性發展_高職 37
                    // 適性發展_綜合高中 38
                    // 適性發展_五專 39

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
            dtDate.Value = _Configure.EndDate;
            LoadStudentTag();
            LoadConfig();

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
                                    dgData1.Rows[rowIdx].Cells[2].Value = elm1.Attribute("TagName").Value;
                                }
                            }

                            if (gpName == "身心障礙")
                            {
                                foreach (XElement elm1 in elm.Elements("Item"))
                                {
                                    int rowIdx = dgData2.Rows.Add();
                                    dgData2.Rows[rowIdx].Cells[colCode1.Index].Value = elm1.Attribute("Code").Value;
                                    dgData2.Rows[rowIdx].Cells[colItem1.Index].Value = elm1.Attribute("Name").Value;
                                    dgData2.Rows[rowIdx].Cells[2].Value = elm1.Attribute("TagName").Value;
                                }
                            }

                            if (gpName == "學生報名身分設定")
                            {
                                foreach (XElement elm1 in elm.Elements("Item"))
                                {
                                    int rowIdx = dgData3.Rows.Add();
                                    dgData3.Rows[rowIdx].Cells[colCode1.Index].Value = elm1.Attribute("Code").Value;
                                    dgData3.Rows[rowIdx].Cells[colItem1.Index].Value = elm1.Attribute("Name").Value;
                                    dgData3.Rows[rowIdx].Cells[2].Value = elm1.Attribute("TagName").Value;
                                }
                            }

                            if (gpName == "失業勞工子女")
                            {
                                foreach (XElement elm1 in elm.Elements("Item"))
                                {
                                    cboSelectTag4.Text = elm1.Attribute("TagName").Value;
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
        /// 載入學生可選類別
        /// </summary>
        private void LoadStudentTag()
        {
            StudentCanSelectTagDict = QueryData.GetStudentAllTag();

            cboSelectTag4.Items.Clear();
            DataGridViewComboBoxColumn cboItem1 = new DataGridViewComboBoxColumn();
            cboItem1.Name = "colStudTag1";
            cboItem1.Width = 150;
            cboItem1.HeaderText = "學生類別";

            DataGridViewComboBoxColumn cboItem2 = new DataGridViewComboBoxColumn();
            cboItem2.Name = "colStudTag1";
            cboItem2.Width = 150;
            cboItem2.HeaderText = "學生類別";

            DataGridViewComboBoxColumn cboItem3 = new DataGridViewComboBoxColumn();
            cboItem3.Name = "colStudTag1";
            cboItem3.Width = 150;
            cboItem3.HeaderText = "學生類別";

            foreach (string name in StudentCanSelectTagDict.Values)
            {
                cboItem1.Items.Add(name);
                cboItem2.Items.Add(name);
                cboItem3.Items.Add(name);

                cboSelectTag4.Items.Add(name);
            }

            dgData1.Columns.Add(cboItem1);
            dgData2.Columns.Add(cboItem2);
            dgData3.Columns.Add(cboItem3);
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
                                        elm1.SetAttributeValue("TagName", drv.Cells[2].Value.ToString());
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
                                        elm1.SetAttributeValue("TagName", drv.Cells[2].Value.ToString());
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
                                        elm1.SetAttributeValue("TagName", drv.Cells[2].Value.ToString());
                                    }
                                }
                            }
                        }

                        if (gpName == "失業勞工子女")
                        {
                            foreach (XElement elm1 in elm.Elements("Item"))
                            {
                                elm1.SetAttributeValue("TagName", cboSelectTag4.Text);

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
                    _Configure.Name = "屏東免試入學-班級服務表現";
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
