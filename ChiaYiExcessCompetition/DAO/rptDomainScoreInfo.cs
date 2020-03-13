using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChiaYiExcessCompetition.DAO
{
    /// <summary>
    /// 成績冊用領域成績
    /// </summary>
    public class rptDomainScoreInfo
    {
        public string StudentID { get; set; }
        public string Name { get; set; }

        public Dictionary<string, decimal> ScoreDict = new Dictionary<string, decimal>();

        public decimal AvgScore = 0;

    }
}
