using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using System.IO;
using System.Xml.Linq;

namespace PingTungExcessCompetition
{
    [FISCA.UDT.TableName(Global._UDTTableName)]
    public class Configure : ActiveRecord
    {
        public Configure()
        {

        }

        /// <summary>
        /// 設定檔名稱
        /// </summary>
        [FISCA.UDT.Field]
        public string Name { get; set; }

        /// <summary>
        /// 列印樣板
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream { get; set; }
        public Aspose.Words.Document Template { get; set; }

        /// <summary>
        /// 統計至日期
        /// </summary>
        [FISCA.UDT.Field]
        public DateTime EndDate { get; set; }

        /// <summary>
        /// 對照表設定
        /// </summary>
        [FISCA.UDT.Field]
        public string MappingContent { get; set; }

        /// <summary>
        /// 在儲存前，把資料填入儲存欄位中
        /// </summary>
        public void Encode()
        {
            MemoryStream stream = new MemoryStream();
            this.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
            this.TemplateStream = Convert.ToBase64String(stream.ToArray());
        }
        /// <summary>
        /// 在資料取出後，把資料從儲存欄位轉換至資料欄位
        /// </summary>
        public void Decode()
        {
            this.Template = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream)));
        }

        /// <summary>
        /// 幹部名稱
        /// </summary>
        [FISCA.UDT.Field]
        public string CadreNames { get; set; }


        public List<string> LoadCareNames()
        {
            List<string> CadreNameList = new List<string>();
            bool isLoadDefault = true;
            try
            {
                if (!string.IsNullOrWhiteSpace(CadreNames))
                {

                    XElement elmRoot = XElement.Parse(CadreNames);
                    foreach (XElement elm in elmRoot.Elements("Item"))
                    {
                        if (elm.Attribute("Name") != null)
                        {
                            CadreNameList.Add(elm.Attribute("Name").Value);
                        }
                    }
                    isLoadDefault = false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            if (isLoadDefault)
            {
                CadreNameList.Add("班長");
                CadreNameList.Add("副班長");
                CadreNameList.Add("學藝股長");
                CadreNameList.Add("風紀股長");
                CadreNameList.Add("衛生股長");
                CadreNameList.Add("服務股長");
                CadreNameList.Add("總務股長");
                CadreNameList.Add("事務股長");
                CadreNameList.Add("康樂股長");
                CadreNameList.Add("體育股長");
                CadreNameList.Add("輔導股長");
                CadreNameList.Add("環保股長");
                CadreNameList.Add("資訊股長");
            }
            return CadreNameList;
        }
    }
}
