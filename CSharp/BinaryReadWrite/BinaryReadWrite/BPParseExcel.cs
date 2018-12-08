using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


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
