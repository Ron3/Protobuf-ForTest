using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace BPSetting
{
    public class BPPay : BPObject
    {
        // 编号
        public int PayID{get; set;}

        // productId
        public string ProductId{get; set;}

        /// <summary>
        /// 构造函数
        /// </summary>
        public BPPay()
        {

        }

        /// <summary>
        /// 解析出值.然后赋值给对应的属性
        /// </summary>
        /// <param name="rowObj"></param>
        /// <param name="titleArray"></param>
        public override void fromExcelRow(IRow rowObj, string[] titleArray)
        {
            if(rowObj == null){
                return;
            }

            // 解释出每一个列
            for (int colIndex = 0; colIndex < titleArray.Length; ++colIndex)
            {
                ICell cellObj = rowObj.GetCell(colIndex);
                if(cellObj == null)
                    continue;

                string strTitle = titleArray[colIndex];
                Dictionary<string, object> dic = new Dictionary<string, object>();
                dic.Add(strTitle, cellObj.ToString());
                
                // 这里可以直接序列化标题 + 值了.不过所有序列化的值都是字符串
                // string jsonData = JsonConvert.SerializeObject(dic);
                // Console.WriteLine("dic ==> " + jsonData);
                

                // switch(cellObj.CellType)
                // {
                //     case CellType.Unknown:
                //     {
                //         // Console.WriteLine("读取到一个字段未知类型异常.进程退出.标题===> " + titleArray[colIndex]);        
                //         break;
                //     }
                    
                //     case CellType.Numeric:
                //     {
                //         break;
                //     }
                    
                //     case CellType.String:
                //         break;

                //     case CellType.Formula:
                //         break;

                //     case CellType.Blank:
                //         break;

                //     case CellType.Boolean:
                //         break;

                //     case CellType.Error:
                //         break;

                //     default:
                //         break;
                // }

                // string val = GetCellString(rowObj.GetCell(colIndex));
                // Console.WriteLine("val ==> " + val);
            }
        }
    }
}

