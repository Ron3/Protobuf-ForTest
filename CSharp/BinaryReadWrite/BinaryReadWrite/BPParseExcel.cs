using System;
using System.Diagnostics;
using System.ComponentModel;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Newtonsoft.Json;


namespace BPParse
{
    public class BPPasrseExcel
    {
        public List<string> pathArray = null;


        #region 基础私有方法

        /// <summary>
        /// 在配置文件里指定的目录下.找出所有的Excel文件
        /// </summary>
        private void _FindAllExcelFile()
        {
            this.pathArray = new List<string>();
            
            // TODO Ron. 先测试
            this.pathArray.Add("./pay.xlsx");
        }


        /// <summary>
        /// 解析Excel
        /// </summary>
        private void _ParseAllExcel()
        {
            if(this.pathArray == null){
                return;
            }

            // 解释单个excel文
            foreach (string path in this.pathArray){
                this._ParseSingleExcel(path);
            }
        }



        /// <summary>
        /// 解释单个excel文件
        /// </summary>
        /// <param name="path"></param>
        private void _ParseSingleExcel(string path)
        {
            if(path == null || path == string.Empty){
                return;
            }

            XSSFWorkbook excelBook;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                excelBook = new XSSFWorkbook(file);
                var enumerator = excelBook.GetEnumerator();
                while(enumerator.MoveNext())
                {
                    ISheet sheetObj = (ISheet)enumerator.Current;
                    this._ParseSingleSheet(sheetObj);
                }
            }
        }


        /// <summary>
        /// 解析单张表
        /// </summary>
        /// <param name="sheetObj"></param>
        private void _ParseSingleSheet(ISheet sheetObj)
        {
            if(sheetObj == null){
                return;
            }

            Console.WriteLine("Parse sheet ==> " + sheetObj.SheetName);

            // 找到这个表的标题.
            List<string> titleList = new List<String>();
            IRow titleRow = sheetObj.GetRow(0);
            int cellCount = titleRow.LastCellNum;
            for(int i = 0; i < cellCount; ++i)
            {
                string title = GetCellString(titleRow.GetCell(i));
                if(this._IsValid(title) == false){
                    break;
                }

                Console.WriteLine("Title Val ==> " + title);
                titleList.Add(title);
            }

            // 转成数组.方便后面操作
            string[] titleArray = titleList.ToArray();

            // 接在开始解析每一行的数据
            Console.WriteLine("sheetObj.LastRowNum ==> " + sheetObj.LastRowNum);
            for(int rowIndex = 1; rowIndex < sheetObj.LastRowNum; ++rowIndex)
            {
                IRow rowObj = sheetObj.GetRow(rowIndex);
                if(rowObj == null){
                    continue;
                }

                // 解释成一个object
                this._ParseToObject(rowObj, sheetObj.SheetName, titleArray);

                // 得到一行json数据
                // this._WriteToJson(rowObj, sheetObj.SheetName, titleArray);

                // // 这里取要注意.就是取标题头的.
                // for (int colIndex = 0; colIndex < titleList.Count; ++colIndex)
                // {
                //     ICell cellObj = rowObj.GetCell(colIndex);
                //     if(cellObj == null)
                //         continue;

                //     this._WriteToJson(cellObj, sheetObj.SheetName, titleArray);

                //     // string val = GetCellString(rowObj.GetCell(colIndex));
                //     // Console.WriteLine("val ==> " + val);
                // }
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private static string Convert(string type, string value)
        {
            switch (type)
            {
                case "int[]":
                case "int32[]":
                case "long[]":
                    return $"[{value}]";
                case "string[]":
                    return $"[{value}]";
                case "int":
                case "int32":
                case "int64":
                case "long":
                case "float":
                case "double":
                    return value;
                case "string":
                    return $"\"{value}\"";
                default:
                    throw new Exception($"不支持此类型: {type}");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="i"></param>
        /// <param name="j"></param>
        /// <returns></returns>
        private static string GetCellString(ISheet sheet, int i, int j)
        {
            return sheet.GetRow(i)?.GetCell(j)?.ToString() ?? "";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        private static string GetCellString(IRow row, int i)
        {
            return row?.GetCell(i)?.ToString() ?? "";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static string GetCellString(ICell cell)
        {
            return cell?.ToString() ?? "";
        }


        /// <summary>
        /// 判定这个标题是否有效
        /// </summary>
        /// <param name="title"></param>
        /// <returns></returns>
        private bool _IsValid(string title)
        {
            if(title.Length > 0){
                return true;
            }

            return false;
        }


        /// <summary>
        /// 解释成一个对象.
        /// </summary>
        /// <param name="rowObj"></param>
        /// <param name="sheetName"></param>
        /// <param name="titleArray"></param>
        /// <returns></returns>
        private object _ParseToObject(IRow rowObj, string sheetName, string[] titleArray)
        {
            if(rowObj == null){
                return null;
            }

            // 1, 找到对应的类名
            string className;
            if(JsonObjectConfig.jsonConfigDic.TryGetValue(sheetName, out className) == false){
                this._Exit("找不到对应的json解析类.");
                return null;
            }

            // 2, 创建出对应的类
            Type t = Type.GetType(className);
            BPSetting.BPObject obj = (BPSetting.BPObject)System.Activator.CreateInstance(t);
            obj.fromExcelRow(rowObj, titleArray);

            // 3, 
            return obj;
        }

        /**
        /// <summary>
        /// 写json格式
        /// </summary>
        /// <param name="rowObj"></param>
        /// <param name="sheetName"></param>
        private string _WriteToJson(IRow rowObj, string sheetName, string[] titleArray)
        {
            if(rowObj == null){
                return null;
            }

            // 1, 找到对应的类名
            string className;
            if(JsonObjectConfig.jsonConfigDic.TryGetValue(sheetName, out className) == false){
                this._Exit("找不到对应的json解析类.");
                return null;
            }

            // 2, 然后通过反射
            Type t = Type.GetType(className);
            Console.WriteLine("t ==> " + t);

            var obj = System.Activator.CreateInstance(t);
            BPSetting.BPPay payObj = (BPSetting.BPPay)obj;
            payObj.PayID = 1;

            string jsonData = JsonConvert.SerializeObject(payObj);
            Console.WriteLine("jsonData ==> " + jsonData);

            return "";
            
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

            // return null;
        }
         */

        /// <summary>
        /// 退出
        /// </summary>
        /// <param name="errMsg"></param>
        private void _Exit(string errMsg="")
        {
            if(errMsg.Length > 0){
                Console.WriteLine("[发生错误.停止解析] errMsg ==> " + errMsg);
            }
            
            Process.GetCurrentProcess().Kill();
        }

        #endregion


        /// <summary>
        /// 总入口
        /// </summary>
        public void Start()
        {
            this._FindAllExcelFile();
            this._ParseAllExcel();
        }
    }


}
