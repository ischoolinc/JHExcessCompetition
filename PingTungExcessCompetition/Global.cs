using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Aspose.Words;

namespace PingTungExcessCompetition
{
    public class Global
    {
        public const string _UDTTableName = "ischool.PingTungExcessCompetition.configure";

        public static string _ProjectName = "屏東免試入學-班級服務表現";

        public static string _DefaultConfTypeName = "預設設定檔";

        /// <summary>
        /// 取得幹部限制
        /// </summary>
        /// <returns></returns>
        //public static List<string> GetCadreName1()
        //{
        //    List<string> CadreName1 = new List<string>();
        //    CadreName1.Add("班長");
        //    CadreName1.Add("副班長");
        //    CadreName1.Add("學藝股長");
        //    CadreName1.Add("風紀股長");
        //    CadreName1.Add("衛生股長");
        //    CadreName1.Add("服務股長");
        //    CadreName1.Add("總務股長");
        //    CadreName1.Add("事務股長");
        //    CadreName1.Add("康樂股長");
        //    CadreName1.Add("體育股長");
        //    CadreName1.Add("輔導股長");
        //    return CadreName1;
        //}



        public static void ExportMappingFieldWord()
        {
            string inputReportName = "屏東免試入學班級服務表現合併欄位總表";
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

            Document tempDoc = new Document(new MemoryStream(Properties.Resources.屏東班級服務表現合併欄位總表));
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(tempDoc);
            builder.MoveToDocumentEnd();
            builder.Writeln();


            builder.StartTable();
            builder.InsertCell(); builder.Write("學校名稱");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "學校名稱" + " \\* MERGEFORMAT ", "«" + "學校名稱" + "»");
            builder.EndRow();
            builder.InsertCell(); builder.Write("班級名稱");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "班級名稱" + " \\* MERGEFORMAT ", "«" + "班級名稱" + "»");
            builder.EndRow();
            builder.InsertCell(); builder.Write("學年度");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD " + "學年度" + " \\* MERGEFORMAT ", "«" + "學年度" + "»");
            builder.EndRow();
            builder.EndTable();

            builder.Writeln();
            builder.Writeln();

            List<string> tmpList = new List<string>();
            tmpList.Add("服務表現項目7上");
            tmpList.Add("服務表現項目7下");
            tmpList.Add("服務表現項目8上");
            tmpList.Add("服務表現項目8下");
            tmpList.Add("服務表現項目9上");
            tmpList.Add("服務表現積分");
            tmpList.Add("服務表現備註");

            builder.StartTable();
            builder.Writeln("服務表現合併欄位");
            builder.InsertCell(); builder.Write("座號");
            builder.InsertCell(); builder.Write("姓名");
            foreach (string name in tmpList)
            {
                string n1 = name.Replace("服務表現", "");
                builder.InsertCell(); builder.Write(n1);
            }

            builder.EndRow();

            for (int studIdx = 1; studIdx <= 50; studIdx++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "座號" + studIdx + " \\* MERGEFORMAT ", "«" + "座" + studIdx + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "姓名" + studIdx + " \\* MERGEFORMAT ", "«" + "姓" + studIdx + "»");

                foreach (string name in tmpList)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + name + studIdx + " \\* MERGEFORMAT ", "«" + "C" + studIdx + "»");
                }

                builder.EndRow();
            }

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