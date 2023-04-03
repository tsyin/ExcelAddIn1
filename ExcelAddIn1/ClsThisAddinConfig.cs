using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;

namespace Config
{
    /// <summary>
    /// 该插件配置文件读写类：读写配置信息
    /// </summary>
    public class ClsThisAddinConfig : ClsBaseConfig
    {
        static string strPath1 = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
        private static ClsThisAddinConfig _Instance = null;
        //构造函数
        public ClsThisAddinConfig(string strPath)
        {
            ConfigPath = strPath;
            ConfigName = "Config.xml";
            RootNodeName = "Config";
        }

        public static ClsThisAddinConfig GetInstance() 
        {
            if (_Instance == null)
                _Instance = new ClsThisAddinConfig(strPath1);
            return _Instance;

        }
        public T readEle<T>(string attribute,T defaultValue)
        {
            T message;
            ClsThisAddinConfig temp = GetInstance();
            message = temp.ReadConfig("Navgater", attribute, defaultValue);
            return message;
        }

        public void writeEle<T>(string attribute,T defaultValue)
        {
            ClsThisAddinConfig temp = GetInstance();
            temp.WriteConfig("Navgater", attribute,defaultValue.ToString());
        }
    }
}