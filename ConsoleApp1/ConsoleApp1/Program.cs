using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ConsoleApp1
{
    /// <summary>
    /// 读取excel文件
    /// </summary>
    public class Program
    {
        static void Main(string[] args)
        {
            XmlConfigurator.Configure();
            ILog log = LogManager.GetLogger("TestLogging");//获取一个日志记录器
            log.Info("程序运行开始记录日志");

            DirectoryInfo d = Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent;
            try
            {
                #region 读取配置文件
                XmlDocument xml = new XmlDocument();
                string siteconfigpath = d.FullName + "\\" + ConfigurationManager.AppSettings["siteconfig"];
                xml.Load(siteconfigpath);

                XmlNode xn = xml.SelectSingleNode("siteconfig");

                if ("TRUE".Equals(xn["addExpertData"].InnerText.ToUpper()))
                {
                    log.Info("读取Excel数据文件");
                    xn["addExpertData"].InnerXml = "false";
                    xml.Save(siteconfigpath);
                    #region 读取专家数据
                    DataSet ds = ExcelToDS(d.FullName + "\\" + ConfigurationManager.AppSettings["excelfile"]);
                    foreach (DataRow item in ds.Tables["table1"].Rows)
                    {
                        if (!String.IsNullOrWhiteSpace(item["专家id"].ToString()))
                        {
                            log.Info("专家id：" + item["专家id"]);
                            log.Info("专家名字：" + item["专家名字"]);
                            log.Info("专家所在地区：" + item["专家所在地区"]);
                            log.Info("专家所在医院：" + item["专家所在医院"]);
                        }
                    }
                    #endregion
                }
                else
                {
                    log.Info("不读取Excel数据文件");
                }
                #endregion
            }
            catch (Exception e)
            {
                log.Error(e.Message);
            }
            log.Info("程序结束");

            Console.ReadKey();
        }
        /// <summary>
        /// Excel表格转为DataSet
        /// 
        /// </summary>
        /// <param name="Path"></param>
        /// <returns></returns>
        public static DataSet ExcelToDS(string Path)
        {
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=Excel 12.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select distinct * from [Sheet1$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            return ds;
        }
    }
}
