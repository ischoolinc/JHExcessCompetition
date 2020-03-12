using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PingTungExcessCompetition.DAO
{
    /// <summary>
    /// 幹部紀錄
    /// </summary>
    public class CadreInfo
    {

        public string StudentID { get; set; }
        public string SchoolYear { get; set; }
        public string Semester { get; set; }
        public string ReferenceType { get; set; }
        public string CadreName { get; set; }

        /// <summary>
        /// 說明欄位
        /// </summary>
        public string Text { get; set; }
    }
}
