using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChiaYiExcessCompetition.DAO
{
    public class rptCompPerformanceInfo
    {
        public string StudentID { get; set; }

        /// <summary>
        /// 競賽名稱
        /// </summary>
        public string name { get; set; }

        /// <summary>
        /// 競賽範圍
        /// </summary>
        public string reward_level { get; set; }

        /// <summary>
        /// 競賽性質(個人賽或團體賽)
        /// </summary>
        public string habitude { get; set; }

        /// <summary>
        /// 主辦單位
        /// </summary>
        public string organizer { get; set; }

        /// <summary>
        /// 證書字號
        /// </summary>
        public string certificate_number { get; set; }

        /// <summary>
        /// 證書日期
        /// </summary>
        public string certificate_date { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        public string remarks { get; set; }

        /// <summary>
        /// 名次
        /// </summary>
        public string rank_name { get; set; }

        /// <summary>
        /// 積分
        /// </summary>
        public decimal bt_integral { get; set; }

        /// <summary>
        /// 名次順序
        /// </summary>
        public string bt_rank_int { get; set; }

    }
}
