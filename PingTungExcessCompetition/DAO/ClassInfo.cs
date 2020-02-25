using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PingTungExcessCompetition.DAO
{
    public class ClassInfo
    {
        public string ClassID { get; set; }
        public string ClassName { get; set; }

        public int DisplayOrder { get; set; }

        List<StudentInfo> StudentInfoList = new List<StudentInfo>();
    }
}
