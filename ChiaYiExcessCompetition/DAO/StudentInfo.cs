using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JHSchool.Data;

namespace ChiaYiExcessCompetition.DAO
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
        /// 台灣身分證
        /// </summary>
        public bool isTaiwanID { get; set; }

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

        // 健康與體育
        public bool isDomainHelPass = false;
        // 綜合活動
        public bool isDomainActPass = false;
        // 藝術與人文//藝術
        public bool isDoaminArtPass = false;
        /// <summary>
        /// 科技
        /// </summary>
        public bool isDoaminTechPass = false;

        /// <summary>
        /// 服務學習時數分數
        /// </summary>
        public int ServiceLearnScore = 0;

        /// <summary>
        /// 學校代碼
        /// </summary>
        public string SchoolCode = "";

        /// <summary>
        /// 學習歷程
        /// </summary>
        public List<SemsHistoryInfo> SemsHistoryInfoList = new List<SemsHistoryInfo>();


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

        /// <summary>
        /// 體適能是否加分
        /// </summary>
        public bool isAddFitnessScore = false;

        /// <summary>
        /// 競賽總積分
        /// </summary>
        public decimal? CompPerfSum = null;


        /// <summary>
        /// 是否低輸入
        /// </summary>
        public bool incomeType1 = false;

        /// <summary>
        /// 是否中低收入
        /// </summary>
        public bool incomeType2 = false;


        public bool hasSemester5Score = false;


        public int CompetitionScore = 0;

        public List<int> CompetitionScoreD = new List<int>();



        /// <summary>
        /// 計算均衡學習
        /// </summary>
        /// <param name="SemsScoreList"></param>
        public void CalcSemsScore5(List<JHSemesterScoreRecord> SemsScoreList)
        {
            Semester5Score = 0;
            hasSemester5Score = true;
            decimal score1 = 0, score2 = 0, score3 = 0, score4=0;

            isDomainHelPass = false;
            isDomainActPass = false;
            isDoaminArtPass = false;
            isDoaminTechPass = false;

            List<string> tmpList = new List<string>();
            foreach (SemsHistoryInfo sh in SemsHistoryInfoList)
            {
                if (sh.GradeYear == "3" && sh.Semester == "2")
                    continue;

                tmpList.Add(sh.SchoolYear + "_" + sh.Semester);
            }

            // 學期歷程未滿5學期
            if (tmpList.Count < 5)
                hasSemester5Score = false;

            int sescoreCount = 0;
            // 成績未滿5學期
            foreach (JHSemesterScoreRecord semsRec in SemsScoreList)
            {
                string key = semsRec.SchoolYear + "_" + semsRec.Semester;
                if (tmpList.Contains(key))
                {
                    sescoreCount += 1;
                }
            }

            if (tmpList.Count != sescoreCount)
                hasSemester5Score = false;

            // 1. 本項基本條件為國中健康體育、藝術人文、綜合活動等三個領域五學期平均成績達及格者。
            // 2.符合基本條件單領域五學期平均成績達及格以上者，計 3 分。
            // 3.符合基本條件 2 領域五學期平均成績達及格以上者，計 6 分。
            // 4.符合基本條件 3 領域五學期平均成
            //績達及格以上者，計 9 分。"
            // 佳樺提供需求需要使用擇優成績，平均計算到整數四捨五入。

            //2023-01-04 加入科技 改上限12分
            //https://3.basecamp.com/4399967/buckets/15765350/todos/5679930879
            foreach (JHSemesterScoreRecord semsRec in SemsScoreList)
            {
                string key = semsRec.SchoolYear + "_" + semsRec.Semester;
                if (tmpList.Contains(key))
                {
                    foreach (string dname in semsRec.Domains.Keys)
                    {
                        if (dname == "健康體育" || dname == "健康與體育")
                        {
                            if (semsRec.Domains[dname].Score.HasValue)
                                score1 += semsRec.Domains[dname].Score.Value;
                        }

                        if (dname == "藝術")
                        {
                            if (semsRec.Domains[dname].Score.HasValue)
                                score2 += semsRec.Domains[dname].Score.Value;
                        }

                        if (dname == "綜合活動")
                        {
                            if (semsRec.Domains[dname].Score.HasValue)
                                score3 += semsRec.Domains[dname].Score.Value;
                        }

                        if (dname == "科技")
                        {
                            if (semsRec.Domains[dname].Score.HasValue)
                                score4 += semsRec.Domains[dname].Score.Value;
                        }
                    }
                }
            }

            score1 = Math.Round(score1 / 5, 0, MidpointRounding.AwayFromZero);
            score2 = Math.Round(score2 / 5, 0, MidpointRounding.AwayFromZero);
            score3 = Math.Round(score3 / 5, 0, MidpointRounding.AwayFromZero);
            score4 = Math.Round(score4 / 5, 0, MidpointRounding.AwayFromZero);

            if (score1 >= 60)
            {
                Semester5Score += 3;
                isDomainHelPass = true;
            }

            if (score2 >= 60)
            {
                Semester5Score += 3;
                isDoaminArtPass = true;
            }

            if (score3 >= 60)
            {
                Semester5Score += 3;
                isDomainActPass = true;
            }

            if (score4 >= 60)
            {
                Semester5Score += 3;
                isDoaminTechPass = true;
            }
            
        }

        /// <summary>
        /// 計算品德表現
        /// </summary>
        public void CalcDemeritMemeritScore(List<JHDemeritRecord> recD, List<JHMeritRecord> recM, JHMeritDemeritReduceRecord mdr)
        {
            int SumRecD = 0, SumRecM = 0;
            // 沒有記錄預設 6 分
            MeritDemeritScore = 6;

            // 使用公布計算方式，不使用系統內功過換算
            int da = 3, db = 3, ma = 3, mb = 3;

            //// 功過換算
            //if (mdr.DemeritAToDemeritB.HasValue)
            //    da = mdr.DemeritAToDemeritB.Value;

            //if (mdr.DemeritBToDemeritC.HasValue)
            //    db = mdr.DemeritBToDemeritC.Value;

            //if (mdr.MeritAToMeritB.HasValue)
            //    ma = mdr.MeritAToMeritB.Value;

            //if (mdr.MeritBToMeritC.HasValue)
            //    mb = mdr.MeritBToMeritC.Value;

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

            foreach (JHMeritRecord rec in recM)
            {
                int b1 = 0, c1 = 0; ;
                if (rec.MeritA.HasValue)
                    b1 = ma * rec.MeritA.Value;
                if (rec.MeritB.HasValue)
                {
                    b1 += rec.MeritB.Value;
                }

                if (rec.MeritC.HasValue)
                    c1 = rec.MeritC.Value;

                c1 += b1 * mb;   // 都換成獎勵
                SumRecM += c1;
            }

            //if (StudentID == "309")
            //{
            //    Console.Write("");
            //}


            // 功過相抵   
            int sum = SumRecM - SumRecD;

            if (sum >= 0)
            {
                MeritDemeritScore += sum;
            }
            else if (sum > -9)
            {
                // 懲戒未達大過
                MeritDemeritScore = 3;
            }
            else
            {
                MeritDemeritScore = 0;
            }

            // 最高 12 分
            if (MeritDemeritScore > 12)
                MeritDemeritScore = 12;

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
            int score = 0;

            // 嘉義版最高9分，每項獲得中等，給3分，最高9分，四項都有銅牌以上加1分

            // 符合達標字串
            List<string> passStringList = new List<string>();
            passStringList.Add("金牌");
            passStringList.Add("銀牌");
            passStringList.Add("銅牌");
            passStringList.Add("中等");
            passStringList.Add("免測");


            if (isSpecial)
            {// 有身心手冊 9分
                FitnessScore = 9;
            }
            else
            {
                FitnessScore = 0;

                // 檢查四項資料
                foreach (string name in passStringList)
                {
                    if (sit_and_reach_degreeList.Contains(name))
                    {
                        score += 3;
                        break;
                    }
                }

                foreach (string name in passStringList)
                {
                    if (standing_long_jump_degreeList.Contains(name))
                    {
                        score += 3;
                        break;
                    }
                }

                foreach (string name in passStringList)
                {
                    if (sit_up_degreeList.Contains(name))
                    {
                        score += 3;
                        break;
                    }
                }

                foreach (string name in passStringList)
                {
                    if (cardiorespiratory_degreeList.Contains(name))
                    {
                        score += 3;
                        break;
                    }
                }

                if (score > 9)
                    score = 9;

                // 檢查每次是否都有銅牌
                if (isAddFitnessScore)
                    score += 1;

                FitnessScore = score;

            }
        }


    }
}
