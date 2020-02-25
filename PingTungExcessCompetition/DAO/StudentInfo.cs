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
        /// 服務積分
        /// </summary>
        public int ServiceScore = 0;
        
        /// <summary>
        /// 計算成績
        /// </summary>
        public void CalcScore()
        {
            // 3 分
            if (HasScore1)
                ServiceScore += 3;

            // 2 分
            if (HasScore2)
                ServiceScore += 2;

             // 2 分
            if (HasScore1)
                ServiceScore += 2;
        }
    }
}
