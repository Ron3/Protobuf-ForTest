using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;


namespace BPSetting
{
    public class BPObject : object
    {
        public virtual void fromExcelRow(IRow rowObj)
        {
            
        }

        public virtual void fromExcelRow(IRow rowObj, string[] titleArray)
        {
            
        }
    }
}

