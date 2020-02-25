using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using FISCA.Permission;
using FISCA.Presentation;

namespace PingTungExcessCompetition
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {

            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["班級", "資料統計"];
            rbItem1["報表"]["屏東免試入學"]["班級服務表現"].Enable = UserAcl.Current["CDBF4D69-AD83-46AB-8F2B-7DEE4ADD03AE"].Executable;
            rbItem1["報表"]["屏東免試入學"]["班級服務表現"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {

                    ServiceReportForm srf = new ServiceReportForm();
                    srf.SetClassIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
                    srf.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選班級");
                    return;
                }

            };

            // 屏東免試入學-班級服務表現
            Catalog catalog1 = RoleAclSource.Instance["班級"]["功能按鈕"];
            catalog1.Add(new RibbonFeature("CDBF4D69-AD83-46AB-8F2B-7DEE4ADD03AE", "屏東免試入學-班級服務表現"));


        }
    }
}
