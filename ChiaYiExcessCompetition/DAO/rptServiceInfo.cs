using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChiaYiExcessCompetition.DAO
{
    /// <summary>
    /// 成績冊用服務學習
    /// </summary>
    public class rptServiceInfo
    {
        public string StudentID { get; set; }

        /// <summary>
        /// 品德表現_服務學習_資料輸入日期
        /// </summary>
        public string OccurDate { get; set; }

        /// <summary>
        /// 品德表現_服務學習_校內外
        /// </summary>
        public string InternalOrExternal { get; set; }

        /// <summary>
        /// 品德表現_服務學習_服務時數
        /// </summary>
        public decimal Hours { get; set; }

        /// <summary>
        /// 品德表現_服務學習_服務學習活動內容
        /// </summary>
        public string Reason { get; set; }

        /// <summary>
        /// 品德表現_服務學習_服務學習證明單位
        /// </summary>
        public string Organizers { get; set; }



    }
}
