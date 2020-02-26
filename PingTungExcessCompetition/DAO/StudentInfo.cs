using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PingTungExcessCompetition.DAO
{
    public class StudentInfo
    {
        public string ClassID { get; set; }

        public string StudentID { get; set; }

        public string StudentName { get; set; }

        public string SeatNo { get; set; }

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
        /// 計算成績
        /// </summary>
        public void CalcScore()
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
                                if (ci.ReferenceType == "班級幹部" && ci.CadreName == "特殊服務表現")
                                {
                                    HasScore3 = true;
                                    ServiceItem_7a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部")
                                {
                                    HasScore1 = true;
                                    ServiceItem_7a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    HasScore2 = true;
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
                                    HasScore3 = true;
                                    ServiceItem_7b.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部")
                                {
                                    HasScore1 = true;
                                    ServiceItem_7b.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    HasScore2 = true;
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
                                    HasScore3 = true;
                                    ServiceItem_8a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部")
                                {
                                    HasScore1 = true;
                                    ServiceItem_8a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    HasScore2 = true;
                                    ServiceItem_8a.Add(ci.CadreName);
                                }
                                else
                                {

                                }
                            }
                        }
                    }

                    if(shi.Semester == "2")
                    {
                        foreach (CadreInfo ci in CadreInfoList)
                        {
                            if (ci.SchoolYear == shi.SchoolYear && ci.Semester == shi.Semester)
                            {
                                if (ci.ReferenceType == "班級幹部" && ci.CadreName == "特殊服務表現")
                                {
                                    HasScore3 = true;
                                    ServiceItem_8b.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部")
                                {
                                    HasScore1 = true;
                                    ServiceItem_8b.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    HasScore2 = true;
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
                                    HasScore3 = true;
                                    ServiceItem_9a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "班級幹部")
                                {
                                    HasScore1 = true;
                                    ServiceItem_9a.Add(ci.CadreName);
                                }
                                else if (ci.ReferenceType == "社團幹部" && ci.CadreName == "社長")
                                {
                                    HasScore2 = true;
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

            // 3 分 班級幹部
            if (HasScore1)
                ServiceScore += 3;

            // 2 分 社團幹部 社長
            if (HasScore2)
                ServiceScore += 2;

            // 2 分 班級幹部 特殊服務表現
            if (HasScore3)
                ServiceScore += 2;
        }
    }
}
