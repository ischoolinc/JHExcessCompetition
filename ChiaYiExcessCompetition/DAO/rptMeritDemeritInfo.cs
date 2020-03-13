using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChiaYiExcessCompetition.DAO
{
    /// <summary>
    /// 成績冊用獎懲
    /// </summary>
    public class rptMeritDemeritInfo
    {
        public string StudentID { get; set; }
        /// <summary>
        /// 品德表現_獎懲_獎懲日期
        /// </summary>
        public string OccurDate { get; set; }

        /// <summary>
        /// 品德表現_獎懲_學期
        /// </summary>
        public string Semester { get; set; }

        /// <summary>
        /// 品德表現_獎懲_獎懲事由
        /// </summary>
        public string Reason { get; set; }

        /// <summary>
        /// 品德表現_獎懲_大功
        /// </summary>
        public string MA { get; set; }

        /// <summary>
        /// 品德表現_獎懲_小功
        /// </summary>
        public string MB { get; set; }

        /// <summary>
        /// 品德表現_獎懲_嘉獎
        /// </summary>
        public string MC { get; set; }

        /// <summary>
        /// 品德表現_獎懲_大過
        /// </summary>
        public string DA { get; set; }

        /// <summary>
        /// 品德表現_獎懲_小過
        /// </summary>
        public string DB { get; set; }

        /// <summary>
        /// 品德表現_獎懲_警告
        /// </summary>
        public string DC { get; set; }

        /// <summary>
        /// 品德表現_獎懲_銷過
        /// </summary>
        public string Cleand { get; set; }
    }
}
