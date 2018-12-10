using System.Collections.Generic;

namespace BPParse
{
    public class Config
    {
        public static string excelDir = "";
        
    }

    public class JsonObjectConfig
    {
        public static Dictionary<string, string> jsonConfigDic = new Dictionary<string, string>
        {
            {"pay", "BPSetting.BPPay"}
        };
    }
}