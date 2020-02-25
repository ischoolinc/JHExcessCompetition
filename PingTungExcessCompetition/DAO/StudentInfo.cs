using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PingTungExcessCompetition.DAO
{
    public class StudentInfo
    {
        public string StudentID { get; set; }

        public string StudentName { get; set; }

        public string SeatNo { get; set; }

        /// <summary>
        /// 一上服務內容
        /// </summary>
        public StringBuilder ServiceItem_1a = new StringBuilder();
        /// <summary>
        /// 一下服務內容
        /// </summary>
        public StringBuilder ServiceItem_1b = new StringBuilder();
        /// <summary>
        /// 二上服務內容
        /// </summary>
        public StringBuilder ServiceItem_2a = new StringBuilder();
        /// <summary>
        /// 二下服務內容
        /// </summary>
        public StringBuilder ServiceItem_2b = new StringBuilder();
        /// <summary>
        /// 三上服務內容
        /// </summary>
        public StringBuilder ServiceItem_3a = new StringBuilder();

        /// <summary>
        /// 服務備註
        /// </summary>
        public StringBuilder ServiceMemo = new StringBuilder();

        /// <summary>
        /// 服務積分
        /// </summary>
        public int ServiceScore = 0;
    }
}
