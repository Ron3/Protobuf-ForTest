using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;


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

            // 接在开始解析每一行的数据
            Console.WriteLine("sheetObj.LastRowNum ==> " + sheetObj.LastRowNum);
            for(int rowIndex = 1; rowIndex < sheetObj.LastRowNum; ++rowIndex)
            {
                IRow rowObj = sheetObj.GetRow(rowIndex);
                if(rowObj == null){
                    break;
                }

                for (int colIndex = 0; colIndex < titleList.Count; ++colIndex)
                {
                    string val = GetCellString(rowObj.GetCell(colIndex));
                    Console.WriteLine("val ==> " + val);
                }
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
