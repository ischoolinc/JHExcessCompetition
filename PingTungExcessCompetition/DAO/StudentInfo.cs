using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JHSchool.Data;

namespace PingTungExcessCompetition.DAO
{
    public class StudentInfo
    {
        public string ClassID { get; set; }

        public string StudentID { get; set; }

        public string StudentName { get; set; }

        public string ClassName { get; set; }

        public string SeatNo { get; set; }

        public string StudentNumber { get; set; }

        public string IDNumber { get; set; }

        /// <summary>
        /// 出生年
        /// </summary>
        public string BirthYear { get; set; }
        /// <summary>
        /// 出生月
        /// </summary>
        public string BirthMonth { get; set; }
        /// <summary>
        /// 出生日
        /// </summary>
        public string BirthDay { get; set; }

        /// <summary>
        /// 報考男女代碼 男1，女 2
        /// </summary>
        public string GenderCode { get; set; }


        public bool HasScore1 = false;
        public bool HasScore2 = false;
        public bool HasScore3 = false;


        /// <summary>
        /// 學習歷程
        /// </summary>
        public List<SemsHistoryInfo> SemsHistoryInfoList = new List<SemsHistoryInfo>();

        /// <summary>
        /// 幹部紀錄
        /// </summary>
        public List<CadreInfo> CadreInfoList = new List<CadreInfo>();

        /// <summary>
        /// 一上服務內容
        /// </summary>
        public List<string> ServiceItem_7a = new List<string>();
        /// <summary>
        /// 一下服務內容
        /// </summary>
        public List<string> ServiceItem_7b = new List<string>();
        /// <summary>
        /// 二上服務內容
        /// </summary>
        public List<string> ServiceItem_8a = new List<string>();
        /// <summary>
        /// 二下服務內容
        /// </summary>
        public List<string> ServiceItem_8b = new List<string>();
        /// <summary>
        /// 三上服務內容
        /// </summary>
        public List<string> ServiceItem_9a = new List<string>();

        /// <summary>
        /// 服務備註
        /// </summary>
        public List<string> ServiceMemo = new List<string>();

        /// <summary>
        /// 是否身心手冊
        /// </summary>
        public bool isSpecial = false;

        /// <summary>
        /// 服務積分
        /// </summary>
        public int ServiceScore = 0;

        /// <summary>
        /// 均衡學習分數
        /// </summary>
        public int Semester5Score = 0;

        /// <summary>
        /// 品德表現分數
        /// </summary>
        public int MeritDemeritScore = 0;

        /// <summary>
        /// 體適能分數
        /// </summary>
        public int FitnessScore = 0;


        public bool hasSemester5Score = false;

        /// <summary>
        /// 計算服務表現成績
        /// </summary>
        public void CalcCadreScore(List<string> CadreNameFilter)
        {
            HasScore1 = HasScore2 = HasScore3 = false;
            // 109 學年度屏東
            // 1.擔任班級幹部任滿 1 學期，經考核表現優良者得 3 分。
            // 2.社團社長任滿 1 學期，經考核表現優良者得 2 分。
            // 3.特殊服務表現：每班 0~4 位，任滿1 學期，經考核表現優良者得 2 分。

            //  社團幹部
            //  學校幹部
            //  班級幹部



            // 取得幹部紀錄比對並放入相關資料
            foreach (SemsHistoryInfo shi in SemsHistoryInfoList)
            {
                if (shi.GradeYear == "1" || shi.GradeYear == "7")
                {
                    if (shi.Semester == "1")
                    {
                        foreach (CadreInfo ci in CadreInfoList)
                        {
                            if (ci.SchoolYear == shi.SchoolYear && ci.Semester == shi.Semester)
                            {
                                //
                                if (ci.ReferenceType == "班級幹部" && ci.CadreName == "特殊服務表現")
                                {
                                    // HasScore3 = true;
                                    ServiceScore += 2;
                                    ServiceItem_7a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部" && CadreNameFilter.Contains(ci.CadreName))
                                {
                                    //  HasScore1 = true;
                                    ServiceScore += 3;
                                    ServiceItem_7a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    //  HasScore2 = true;
                                    ServiceScore += 2;
                                    ServiceItem_7a.Add(ci.CadreName);
                                }
                                else
                                {

                                }
                            }
                        }
                    }

                    if (shi.Semester == "2")
                    {
                        foreach (CadreInfo ci in CadreInfoList)
                        {
                            if (ci.SchoolYear == shi.SchoolYear && ci.Semester == shi.Semester)
                            {
                                if (ci.ReferenceType == "班級幹部" && ci.CadreName == "特殊服務表現")
                                {
                                    // HasScore3 = true;
                                    ServiceScore += 2;
                                    ServiceItem_7b.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部" && CadreNameFilter.Contains(ci.CadreName))
                                {
                                    //  HasScore1 = true;
                                    ServiceScore += 3;
                                    ServiceItem_7b.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    // HasScore2 = true;
                                    ServiceScore += 2;
                                    ServiceItem_7b.Add(ci.CadreName);
                                }
                                else
                                {

                                }
                            }
                        }
                    }
                }

                if (shi.GradeYear == "2" || shi.GradeYear == "8")
                {
                    if (shi.Semester == "1")
                    {
                        foreach (CadreInfo ci in CadreInfoList)
                        {
                            if (ci.SchoolYear == shi.SchoolYear && ci.Semester == shi.Semester)
                            {
                                if (ci.ReferenceType == "班級幹部" && ci.CadreName == "特殊服務表現")
                                {
                                    // HasScore3 = true;
                                    ServiceScore += 2;
                                    ServiceItem_8a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部" && CadreNameFilter.Contains(ci.CadreName))
                                {
                                    // HasScore1 = true;
                                    ServiceScore += 3;
                                    ServiceItem_8a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    // HasScore2 = true;
                                    ServiceScore += 2;
                                    ServiceItem_8a.Add(ci.CadreName);
                                }
                                else
                                {

                                }
                            }
                        }
                    }

                    if (shi.Semester == "2")
                    {
                        foreach (CadreInfo ci in CadreInfoList)
                        {
                            if (ci.SchoolYear == shi.SchoolYear && ci.Semester == shi.Semester)
                            {
                                if (ci.ReferenceType == "班級幹部" && ci.CadreName == "特殊服務表現")
                                {
                                    // HasScore3 = true;
                                    ServiceScore += 2;
                                    ServiceItem_8b.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部" && CadreNameFilter.Contains(ci.CadreName))
                                {
                                    //HasScore1 = true;
                                    ServiceScore += 3;
                                    ServiceItem_8b.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    //  HasScore2 = true;
                                    ServiceScore += 2;
                                    ServiceItem_8b.Add(ci.CadreName);
                                }
                                else
                                {

                                }
                            }
                        }
                    }

                }

                if (shi.GradeYear == "3" || shi.GradeYear == "9")
                {
                    if (shi.Semester == "1")
                    {
                        foreach (CadreInfo ci in CadreInfoList)
                        {
                            if (ci.SchoolYear == shi.SchoolYear && ci.Semester == shi.Semester)
                            {
                                if (ci.ReferenceType == "班級幹部" && ci.CadreName == "特殊服務表現")
                                {
                                    //HasScore3 = true;
                                    ServiceScore += 2;
                                    ServiceItem_9a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部" && CadreNameFilter.Contains(ci.CadreName))
                                {
                                    // HasScore1 = true;
                                    ServiceScore += 3;
                                    ServiceItem_9a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    // HasScore2 = true;
                                    ServiceScore += 2;
                                    ServiceItem_9a.Add(ci.CadreName);
                                }
                                else
                                {

                                }
                            }
                        }
                    }
                }
            }

            // 因規定上限10分
            if (ServiceScore > 10)
                ServiceScore = 10;

            //// 3 分 班級幹部
            //if (HasScore1)
            //    ServiceScore += 3;

            //// 2 分 社團幹部 社長
            //if (HasScore2)
            //    ServiceScore += 2;

            //// 2 分 班級幹部 特殊服務表現
            //if (HasScore3)
            //    ServiceScore += 2;
        }

        /// <summary>
        /// 計算均衡學習
        /// </summary>
        /// <param name="SemsScoreList"></param>
        public void CalcSemsScore5(List<JHSemesterScoreRecord> SemsScoreList)
        {
            Semester5Score = 0;
            hasSemester5Score = true;
            decimal score1 = 0, score2 = 0, score3 = 0;

            List<string> tmpList = new List<string>();
            foreach (SemsHistoryInfo sh in SemsHistoryInfoList)
            {
                if (sh.GradeYear != "3" && sh.Semester != "2")
                    tmpList.Add(sh.SchoolYear + "_" + sh.Semester);
            }

            // 學期歷程未滿5學期
            if (tmpList.Count != 5)
                hasSemester5Score = false;

            // 成績未滿5學期
            foreach (JHSemesterScoreRecord semsRec in SemsScoreList)
            {
                string key = semsRec.SchoolYear + "_" + semsRec.Semester;
                if (!tmpList.Contains(key))
                    hasSemester5Score = false;
            }


            // 1. 本項基本條件為國中健康體育、藝術人文、綜合活動等三個領域五學期平均成績達及格者。
            // 2.符合基本條件單領域五學期平均成績達及格以上者，計 3 分。
            // 3.符合基本條件 2 領域五學期平均成績達及格以上者，計 6 分。
            // 4.符合基本條件 3 領域五學期平均成
            //績達及格以上者，計 9 分。"

            foreach (JHSemesterScoreRecord semsRec in SemsScoreList)
            {
                foreach (string dname in semsRec.Domains.Keys)
                {
                    if (dname == "健康體育")
                    {
                        if (semsRec.Domains[dname].Score.HasValue)
                            score1 += semsRec.Domains[dname].Score.Value;
                    }

                    if (dname == "藝術人文")
                    {
                        if (semsRec.Domains[dname].Score.HasValue)
                            score2 += semsRec.Domains[dname].Score.Value;
                    }

                    if (dname == "綜合活動")
                    {
                        if (semsRec.Domains[dname].Score.HasValue)
                            score3 += semsRec.Domains[dname].Score.Value;
                    }
                }
            }

            if ((score1 / 5) >= 60)
                Semester5Score += 3;

            if ((score2 / 5) >= 60)
                Semester5Score += 3;

            if ((score3 / 5) >= 60)
                Semester5Score += 3;


        }

        /// <summary>
        /// 計算品德表現
        /// </summary>
        public void CalcDemeritMemeritScore(List<JHDemeritRecord> recD, List<JHMeritRecord> recM, JHMeritDemeritReduceRecord mdr)
        {
            int SumRecD = 0, SumRecM = 0;


            int da = 3, db = 3, ma = 3, mb = 3;

            // 功過換算
            if (mdr.DemeritAToDemeritB.HasValue)
                da = mdr.DemeritAToDemeritB.Value;

            if (mdr.DemeritBToDemeritC.HasValue)
                db = mdr.DemeritBToDemeritC.Value;

            if (mdr.MeritAToMeritB.HasValue)
                ma = mdr.MeritAToMeritB.Value;

            if (mdr.MeritBToMeritC.HasValue)
                mb = mdr.MeritBToMeritC.Value;

            foreach (JHDemeritRecord rec in recD)
            {
                int b1 = 0, c1 = 0; ;
                if (rec.DemeritA.HasValue)
                    b1 = da * rec.DemeritA.Value;
                if (rec.DemeritB.HasValue)
                {
                    b1 += rec.DemeritB.Value;
                }

                if (rec.DemeritC.HasValue)
                    c1 = rec.DemeritC.Value;

                c1 += b1 * db;   // 都換成警告
                SumRecD += c1;
            }

            //foreach (JHMeritRecord rec in recM)
            //{
            //    int b1 = 0, c1 = 0; ;
            //    if (rec.MeritA.HasValue)
            //        b1 = ma * rec.MeritA.Value;
            //    if (rec.MeritB.HasValue)
            //    {
            //        b1 += rec.MeritB.Value;
            //    }

            //    if (rec.MeritC.HasValue)
            //        c1 = rec.MeritC.Value;

            //    c1 += b1 * mb;   // 都換成獎勵
            //    SumRecM += c1;
            //}

            // 功過相抵 (//使用這提到不功過相抵)  sum = SumRecD - SumRecM;
            int sum = SumRecD;

            if (sum < 1)
                MeritDemeritScore = 10;
            else if (sum < db)
                MeritDemeritScore = 7;
            else if (sum < (db * 2))
                MeritDemeritScore = 4;
            else
                MeritDemeritScore = 0;
        }

        /// <summary>
        /// 體適能用
        /// </summary>
        public List<string> sit_and_reach_degreeList = new List<string>();
        public List<string> standing_long_jump_degreeList = new List<string>();
        public List<string> sit_up_degreeList = new List<string>();
        public List<string> cardiorespiratory_degreeList = new List<string>();

        /// <summary>
        /// 計算體適能分數
        /// </summary>
        public void CalcFitnessScore()
        {
            int ItemCount = 0, passCount = 0;


            // 符合達標字串
            List<string> passStringList = new List<string>();
            passStringList.Add("金牌");
            passStringList.Add("銀牌");
            passStringList.Add("銅牌");
            passStringList.Add("中等");
            passStringList.Add("免測");


            if (isSpecial)
            {// 有身心手冊 8分
                FitnessScore = 8;
            }
            else
            {
                FitnessScore = 0;

                if (sit_and_reach_degreeList.Count > 0)
                    ItemCount += 1;
                if (standing_long_jump_degreeList.Count > 0)
                    ItemCount += 1;
                if (sit_up_degreeList.Count > 0)
                    ItemCount += 1;
                if (cardiorespiratory_degreeList.Count > 0)
                    ItemCount += 1;

                // 檢查四項資料
                foreach (string name in passStringList)
                {
                    if (sit_and_reach_degreeList.Contains(name))
                    {
                        passCount += 1;
                        break;
                    }
                }

                foreach (string name in passStringList)
                {
                    if (standing_long_jump_degreeList.Contains(name))
                    {
                        passCount += 1;
                        break;
                    }
                }

                foreach (string name in passStringList)
                {
                    if (sit_up_degreeList.Contains(name))
                    {
                        passCount += 1;
                        break;
                    }
                }

                foreach (string name in passStringList)
                {
                    if (cardiorespiratory_degreeList.Contains(name))
                    {
                        passCount += 1;
                        break;
                    }
                }

                if (ItemCount == 4)
                {
                    // 有四項但都未達標 2 分
                    FitnessScore = 2;

                    if (passCount == 1)
                        FitnessScore = 4;

                    if (passCount == 2)
                        FitnessScore = 6;

                    if (passCount == 3)
                        FitnessScore = 8;

                    if (passCount == 4)
                        FitnessScore = 10;

                }
            }
        }

    }
}
