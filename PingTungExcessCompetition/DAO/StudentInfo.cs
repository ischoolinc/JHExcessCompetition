﻿using System;
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
        /// 服務積分
        /// </summary>
        public int ServiceScore = 0;

        /// <summary>
        /// 均衡學習分數
        /// </summary>
        public int Semester5Score = 0;


        /// <summary>
        /// 計算成績
        /// </summary>
        public void CalcScore(List<string> CadreNameFilter)
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

        public void CalcSemsScore5(List<JHSemesterScoreRecord> SemsScoreList)
        {
            Semester5Score = 0;
            
            // 1. 本項基本條件為國中健康體育、藝術人文、綜合活動等三個領域五學期平均成績達及格者。
            // 2.符合基本條件單領域五學期平均成績達及格以上者，計 3 分。
            // 3.符合基本條件 2 領域五學期平均成績達及格以上者，計 6 分。
            // 4.符合基本條件 3 領域五學期平均成
            //績達及格以上者，計 9 分。"

            foreach(JHSemesterScoreRecord semsRec in SemsScoreList)
            {

            }


        }


    }
}
