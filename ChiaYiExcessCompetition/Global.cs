using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Aspose.Words;

namespace ChiaYiExcessCompetition
{
    public class Global
    {
        public const string _UDTTableName = "ischool.ChiaYiExcessCompetition.configure";

        public static string _ProjectName = "嘉義區超額比序均衡學習成績證明單";

        public static string _DefaultConfTypeName = "預設設定檔";

        public static void ExportMappingFieldWord()
        {

            string inputReportName = "嘉義區超額比序均衡學習成績證明單合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            Document tempDoc = new Document(new MemoryStream(Properties.Resources.嘉義區成績冊合併欄位總表));
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(tempDoc);
            builder.MoveToDocumentEnd();
            builder.Writeln();


            builder.StartTable();
            builder.InsertCell(); builder.Write("學年度");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "學年度" + " \\* MERGEFORMAT ", "«" + "學年度" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("學校名稱");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "學校名稱" + " \\* MERGEFORMAT ", "«" + "學校名稱" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("班級");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "班級" + " \\* MERGEFORMAT ", "«" + "班級" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("座號");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "座號" + " \\* MERGEFORMAT ", "«" + "座號" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("姓名");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "姓名" + " \\* MERGEFORMAT ", "«" + "姓名" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("學號");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "學號" + " \\* MERGEFORMAT ", "«" + "學號" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("扶助弱勢身分");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "扶助弱勢_身分" + " \\* MERGEFORMAT ", "«" + "扶助弱勢身分" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("扶助弱勢積分");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "扶助弱勢_積分" + " \\* MERGEFORMAT ", "«" + "扶助弱勢積分" + "»");
            builder.EndRow();

            builder.EndTable();

            List<string> domainNameList = new List<string>();
            domainNameList.Add("健康與體育");
            domainNameList.Add("藝術");
            domainNameList.Add("綜合活動");
            domainNameList.Add("科技");

            builder.Writeln("");
            builder.Writeln("均衡學習-領域成績");
            builder.StartTable();
            builder.InsertCell(); builder.Write("領域");
            builder.InsertCell(); builder.Write("七上");
            builder.InsertCell(); builder.Write("七下");
            builder.InsertCell(); builder.Write("八上");
            builder.InsertCell(); builder.Write("八下");
            builder.InsertCell(); builder.Write("九上");
            builder.InsertCell(); builder.Write("平均");
            builder.EndRow();

            foreach (string dName in domainNameList)
            {
                builder.InsertCell();
                builder.Write(dName);

                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "均衡學習_" + dName + "_七上分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "均衡學習_" + dName + "_七下分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "均衡學習_" + dName + "_八上分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "均衡學習_" + dName + "_八下分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "均衡學習_" + dName + "_九上分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "均衡學習_" + dName + "_平均" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.EndRow();
            }
            builder.EndTable();
            builder.Writeln();

            List<string> otherDomainNameList = new List<string>();
            otherDomainNameList.Add("語文");
            otherDomainNameList.Add("數學");
            otherDomainNameList.Add("自然科學");
            //otherDomainNameList.Add("科技");

            builder.Writeln("其他領域成績");
            builder.StartTable();
            builder.InsertCell(); builder.Write("領域");
            builder.InsertCell(); builder.Write("七上");
            builder.InsertCell(); builder.Write("七下");
            builder.InsertCell(); builder.Write("八上");
            builder.InsertCell(); builder.Write("八下");
            builder.InsertCell(); builder.Write("九上");
            builder.InsertCell(); builder.Write("平均");
            builder.EndRow();

            foreach (string dName in otherDomainNameList)
            {
                builder.InsertCell();
                builder.Write(dName);

                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "其他領域_" + dName + "_七上分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "其他領域_" + dName + "_七下分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "其他領域_" + dName + "_八上分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "其他領域_" + dName + "_八下分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "其他領域_" + dName + "_九上分數" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "其他領域_" + dName + "_平均" + " \\* MERGEFORMAT ", "«" + "DS" + "»");
                builder.EndRow();
            }
            builder.EndTable();
            builder.Writeln();

            builder.Writeln("獎懲統計");
            builder.StartTable();
            builder.InsertCell(); builder.Write("大功");
            builder.InsertCell(); builder.Write("小功");
            builder.InsertCell(); builder.Write("嘉獎");
            builder.InsertCell(); builder.Write("大過");
            builder.InsertCell(); builder.Write("小過");
            builder.InsertCell(); builder.Write("警告");
            builder.InsertCell(); builder.Write("銷過");
            builder.EndRow();

            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_大功統計" + " \\*MERGEFORMAT ", "«" + "MD" + "»");
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_小功統計" + " \\*MERGEFORMAT ", "«" + "MD" + "»");
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_嘉獎統計" + " \\*MERGEFORMAT ", "«" + "MD" + "»");
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_大過統計" + " \\*MERGEFORMAT ", "«" + "MD" + "»");
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_小過統計" + " \\*MERGEFORMAT ", "«" + "MD" + "»");
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_警告統計" + " \\*MERGEFORMAT ", "«" + "MD" + "»");
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_銷過統計" + " \\*MERGEFORMAT ", "«" + "MD" + "»");

            builder.EndRow();
            builder.EndTable();

            builder.Writeln();
            builder.StartTable();
            builder.InsertCell(); builder.Write("均衡學習積分");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "均衡學習_積分" + " \\* MERGEFORMAT ", "«" + "均衡學習積分" + "»");
            builder.EndRow();
            builder.InsertCell(); builder.Write("品德表現獎懲積分");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "品德表現_獎懲_積分" + " \\* MERGEFORMAT ", "«" + "品德表現獎懲積分" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("品德表現服務學習積分");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "品德表現_服務學習_積分" + " \\* MERGEFORMAT ", "«" + "品德表現服務學習積分" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("品德表現_服務學習校內時數統計");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "品德表現_服務學習_校內時數統計" + " \\* MERGEFORMAT ", "«" + "SC" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("品德表現_服務學習校外時數統計");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "品德表現_服務學習_校外時數統計" + " \\* MERGEFORMAT ", "«" + "SC" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("品德表現_服務學習校內外時數統計");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "品德表現_服務學習_校內外時數統計" + " \\* MERGEFORMAT ", "«" + "SC" + "»");
            builder.EndRow();

            builder.InsertCell(); builder.Write("品德表現_體適能積分");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "品德表現_體適能_積分" + " \\* MERGEFORMAT ", "«" + "品德表現_體適能積分" + "»");
            builder.EndRow();
            builder.EndTable();

            builder.Writeln();
            builder.Writeln("品德表現_獎懲明細");
            builder.StartTable();
            builder.InsertCell(); builder.Write("獎懲日期");
            builder.InsertCell(); builder.Write("學期");
            builder.InsertCell(); builder.Write("獎懲事由");
            builder.InsertCell(); builder.Write("大功");
            builder.InsertCell(); builder.Write("小功");
            builder.InsertCell(); builder.Write("嘉獎");
            builder.InsertCell(); builder.Write("大過");
            builder.InsertCell(); builder.Write("小過");
            builder.InsertCell(); builder.Write("警告");
            builder.InsertCell(); builder.Write("銷過");
            builder.EndRow();

            for (int i = 1; i <= 50; i++)
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_獎懲日期" + i + " \\* MERGEFORMAT ", "«" + "MD" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_學期" + i + " \\* MERGEFORMAT ", "«" + "MD" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_獎懲事由" + i + " \\* MERGEFORMAT ", "«" + "MD" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_大功" + i + " \\* MERGEFORMAT ", "" + "D" + i + "");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_小功" + i + " \\* MERGEFORMAT ", "" + "D" + i + "");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_嘉獎" + i + " \\* MERGEFORMAT ", "" + "D" + i + "");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_大過" + i + " \\* MERGEFORMAT ", "" + "D" + i + "");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_小過" + i + " \\* MERGEFORMAT ", "" + "D" + i + "");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_警告" + i + " \\* MERGEFORMAT ", "" + "D" + i + "");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_獎懲_銷過" + i + " \\* MERGEFORMAT ", "" + "D" + i + "");
                builder.EndRow();
            }
            builder.EndTable();

            builder.Writeln();
            builder.Writeln("品德表現_服務學習明細");
            builder.StartTable();
            builder.InsertCell(); builder.Write("資料輸入日期");
            builder.InsertCell(); builder.Write("校內外");
            builder.InsertCell(); builder.Write("服務時數");
            builder.InsertCell(); builder.Write("服務學習活動內容");
            builder.InsertCell(); builder.Write("服務學習證明單位");
            builder.EndRow();

            for (int i = 1; i <= 30; i++)
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_服務學習_資料輸入日期" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_服務學習_校內外" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_服務學習_服務時數" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_服務學習_服務學習活動內容" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_服務學習_服務學習證明單位" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.EndRow();
            }
            builder.EndTable();

            builder.Writeln();
            builder.Writeln("品德表現_體適能明細");
            builder.StartTable();
            builder.InsertCell(); builder.Write("檢測日期");
            builder.InsertCell(); builder.Write("年齡");
            builder.InsertCell(); builder.Write("性別");
            builder.InsertCell(); builder.Write("坐姿體前彎_成績");
            builder.InsertCell(); builder.Write("坐姿體前彎_等級");
            builder.InsertCell(); builder.Write("立定跳遠_成績");
            builder.InsertCell(); builder.Write("立定跳遠_等級");
            builder.InsertCell(); builder.Write("仰臥起坐_成績");
            builder.InsertCell(); builder.Write("仰臥起坐_等級");
            builder.InsertCell(); builder.Write("公尺跑走_成績");
            builder.InsertCell(); builder.Write("公尺跑走_等級");
            builder.InsertCell(); builder.Write("仰臥捲腹_成績");
            builder.InsertCell(); builder.Write("仰臥捲腹_等級");
            builder.InsertCell(); builder.Write("漸速耐力跑_成績");
            builder.InsertCell(); builder.Write("漸速耐力跑_等級");
            builder.EndRow();

            for (int i = 1; i <= 10; i++)
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_檢測日期" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_年齡" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_性別" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_坐姿體前彎_成績" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_坐姿體前彎_等級" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_立定跳遠_成績" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_立定跳遠_等級" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_仰臥起坐_成績" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_仰臥起坐_等級" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_公尺跑走_成績" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");

                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_公尺跑走_等級" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");


                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_仰臥捲腹_成績" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");

                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_仰臥捲腹_等級" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");

                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_漸速耐力跑_成績" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");

                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "品德表現_體適能_漸速耐力跑_等級" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");

                builder.EndRow();
            }

            builder.EndTable();


            builder.Writeln();
            builder.Writeln("競賽成績");
            builder.StartTable();
            builder.InsertCell(); builder.Write("競賽層級");
            builder.InsertCell(); builder.Write("競賽性質");
            builder.InsertCell(); builder.Write("競賽名稱");
            builder.InsertCell(); builder.Write("得獎名次");
            builder.InsertCell(); builder.Write("證書日期");
            builder.InsertCell(); builder.Write("主辦單位");
            builder.EndRow();

            for (int i = 1; i <= 20; i++)
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_競賽層級" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_競賽性質" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_競賽名稱" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_得獎名次" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_證書日期" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_主辦單位" + i + " \\* MERGEFORMAT ", "«" + "D" + i + "»");
                builder.EndRow();
            }
            builder.EndTable();

            builder.Writeln();
            builder.Writeln("競賽成績-統計");
            builder.Writeln("縣市個人");
            builder.StartTable();
            for (int i = 1; i <= 8; i++)
            {
                builder.InsertCell(); builder.Write("名次" + i);
            }
            builder.InsertCell(); builder.Write("其他");
            builder.EndRow();
            for (int i = 1; i <= 8; i++)
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_縣市個人名次" + i + " \\* MERGEFORMAT ", "" + "C" + "");
            }
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_縣市個人名次_其他" + " \\* MERGEFORMAT ", "" + "C" + "");
            builder.EndRow();
            builder.EndTable();

            builder.Writeln("縣市團體");
            builder.StartTable();
            for (int i = 1; i <= 8; i++)
            {
                builder.InsertCell(); builder.Write("名次" + i);
            }
            builder.InsertCell(); builder.Write("其他");
            builder.EndRow();
            for (int i = 1; i <= 8; i++)
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_縣市團體名次" + i + " \\* MERGEFORMAT ", "" + "C" + "");
            }
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_縣市團體名次_其他" + " \\* MERGEFORMAT ", "" + "C" + "");
            builder.EndRow();
            builder.EndTable();

            builder.Writeln("全國個人");
            builder.StartTable();
            for (int i = 1; i <= 8; i++)
            {
                builder.InsertCell(); builder.Write("名次" + i);
            }
            builder.InsertCell(); builder.Write("其他");
            builder.EndRow();
            for (int i = 1; i <= 8; i++)
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_全國個人名次" + i + " \\* MERGEFORMAT ", "" + "C" + "");
            }
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_全國個人名次_其他" + " \\* MERGEFORMAT ", "" + "C" + "");
            builder.EndRow();
            builder.EndTable();

            builder.Writeln("全國團體");
            builder.StartTable();
            for (int i = 1; i <= 8; i++)
            {
                builder.InsertCell(); builder.Write("名次" + i);
            }
            builder.InsertCell(); builder.Write("其他");
            builder.EndRow();
            for (int i = 1; i <= 8; i++)
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_全國團體名次" + i + " \\* MERGEFORMAT ", "" + "C" + "");
            }
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_全國團體名次_其他" + " \\* MERGEFORMAT ", "" + "C" + "");
            builder.EndRow();
            builder.EndTable();

            builder.Writeln("");
            builder.StartTable();
            builder.InsertCell(); builder.Write("競賽成績_競賽積分");
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "競賽成績_競賽積分" + " \\* MERGEFORMAT ", "«" + "S" + "»");
            builder.EndRow();
            builder.EndTable();

            builder.Writeln("");
            builder.Writeln("成績-合計總分");
            builder.StartTable();
            builder.InsertCell(); builder.Write("成績_合計總分");
            builder.InsertCell(); builder.InsertField("MERGEFIELD " + "成績_合計總分" + " \\* MERGEFORMAT ", "«" + "S" + "»");
            builder.EndRow();
            builder.EndTable();



            try
            {

                tempDoc.Save(path, SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        tempDoc.Save(sd.FileName, SaveFormat.Doc);

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }

        }
    }
}
