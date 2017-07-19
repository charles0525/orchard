using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Data;
using System.IO;

namespace Acai.NewsStatistics.Common
{
    public class ExcelHelper
    {
        public static readonly ExcelHelper instance = new ExcelHelper();

        IWorkbook workbook;

        #region Excel导出

        /// <summary>
        /// For DataTable Excel 导出
        /// </summary>
        /// <param name="dtTable">导出表</param>
        /// <param name="dicHeadInfo">顶部统计信息</param>
        /// <param name="listColumns">报表自定义列</param>
        /// <param name="fileName">导出文件名称</param>
        /// <param name="sheetName">Sheet 名称</param>
        /// <param name="isAjax">isAjax=true 返回文件地址，否则输出流</param>
        public string ExportFromTable(DataTable dtTable, Dictionary<string, string> dicHeadInfo = null, List<ExcelColumns> listColumns = null, string fileName = "", string sheetName = "", bool isAjax = false)
        {
            if (dtTable == null || listColumns == null || listColumns.Count <= 0)
                return "";

            if (string.IsNullOrEmpty(sheetName))
                sheetName = "Sheet1";

            if (string.IsNullOrEmpty(fileName))
                fileName = DateTime.Now.ToString("yyyyMMddHHmmss") + new Random().Next(1000, 9999).ToString();

            workbook = new XSSFWorkbook(); //HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(sheetName);
            sheet.DefaultColumnWidth = 16;

            ICellStyle cellStyle = null;//单元格样式
            ICellStyle columnStyle = null;//报表列样式
            ICellStyle headStyle = null;//顶部统计信息样式
            SetCellStyle(out cellStyle, out columnStyle, out headStyle);

            string mergedCells = string.Empty;//纵向合并列集合
            int colIndex = 0, rowIndex = 0;//全局列索引，行索引
            int columnsRowIndex = 0;//标题列行索引
            ICell icell = null;

            #region 顶部统计等信息

            //顶部统计信息
            if (dicHeadInfo != null && dicHeadInfo.Count > 0)
            {
                IRow cellHeader = sheet.CreateRow(rowIndex);
                cellHeader.Height = 500;
                foreach (var col in dicHeadInfo)
                {
                    if (col.Value == "000")//换行
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, listColumns.Count - 1));//合并
                        rowIndex++;
                        colIndex = 0;
                        cellHeader = sheet.CreateRow(rowIndex);
                        cellHeader.Height = 500;
                    }
                    else
                    {
                        icell = cellHeader.CreateCell(colIndex);
                        icell.SetCellValue(col.Value);
                        icell.CellStyle = headStyle;
                        colIndex++;
                    }
                }
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, listColumns.Count - 1));
                rowIndex++;//行索引+1
            }

            #endregion

            #region 创建标题列

            IRow cellColumn = sheet.CreateRow(rowIndex);
            cellColumn.Height = 350;
            colIndex = 0;
            columnsRowIndex = rowIndex;//获取标题列所在行索引
            foreach (ExcelColumns col in listColumns)
            {
                icell = cellColumn.CreateCell(colIndex);
                icell.SetCellValue(col.FieldTitle);
                icell.CellStyle = columnStyle;
                colIndex++;
                if (col.IsMerged)//如果合并，记录下需要合并的列
                    mergedCells += col.FieldTitle + ",";
            }
            /*
             第一个参数表示要冻结的列数；
            第二个参数表示要冻结的行数；
            第三个参数表示右边区域可见的首列序号，从1开始计算；
            第四个参数表示下边区域可见的首行序号，也是从1开始计算；
             */
            sheet.CreateFreezePane(0, rowIndex + 1, 0, 1);
            rowIndex++;//行索引+1

            #endregion

            #region 创建单元格，填充内容

            //创建单元格
            string strValue = string.Empty;
            Dictionary<string, decimal> dicBottomInfo = null;
            foreach (DataRow row in dtTable.Rows)
            {
                IRow cellRow = sheet.CreateRow(rowIndex);
                cellRow.Height = 100 * 3;
                colIndex = 0;
                foreach (ExcelColumns col in listColumns)
                {
                    //判断自定义列 FieldName 是否与表字段名一致，不一致则单元格赋予空值
                    if (dtTable.Columns[col.FieldName] != null)
                    {
                        //数据格式化
                        if (col.FieldType == "Int")//整型
                            strValue = string.Format("{0:0.#}", string.IsNullOrEmpty(row[col.FieldName].ToString()) ? 0 : row[col.FieldName]);
                        else if (col.FieldType == "Dec")//小数
                            strValue = string.Format("{0:0.00}", string.IsNullOrEmpty(row[col.FieldName].ToString()) ? 0 : row[col.FieldName]);
                        else if (col.FieldType == "Date")//日期
                            strValue = string.Format("{0:yyyy年MM月dd日}", row[col.FieldName]);
                        else if (col.FieldType == "DateTime")//时间
                            strValue = string.Format("{0:yyyy年MM月dd日 HH:mm:ss}", row[col.FieldName]);
                        else
                            strValue = row[col.FieldName].ToString();

                        //通过委托调用方法
                        if (col.Action != null)
                            strValue = col.Action(strValue);

                        //记录统计列信息
                        if (col.IsAccount)
                        {
                            if (dicBottomInfo == null)
                                dicBottomInfo = new Dictionary<string, decimal>();
                            //如果集合不包含该列则添加,以Excel标题列为键值
                            if (!dicBottomInfo.ContainsKey(col.FieldTitle))
                                dicBottomInfo.Add(col.FieldTitle, decimal.Parse(string.IsNullOrEmpty(strValue) ? "0" : strValue));
                            else//如果存在则累加
                                dicBottomInfo[col.FieldTitle] += decimal.Parse(string.IsNullOrEmpty(strValue) ? "0" : strValue);
                        }
                    }
                    else
                        strValue = string.Empty;

                    icell = cellRow.CreateCell(colIndex);
                    icell.SetCellValue(strValue);
                    icell.CellStyle = cellStyle;
                    colIndex++;
                }
                rowIndex++;
            }

            //纵向合并
            if (!string.IsNullOrEmpty(mergedCells))
                MergedCells(sheet, mergedCells, columnsRowIndex);

            #endregion

            #region 导出

            string tmpPath = "/upload/export/tmp";
            string tmpFile = "";
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Write(stream);
                if (isAjax)
                {
                    if (!Directory.Exists(HttpContext.Current.Server.MapPath(tmpPath)))
                        Directory.CreateDirectory(HttpContext.Current.Server.MapPath(tmpPath));
                    tmpFile = string.Format("{0}\\{1}.xlsx", HttpContext.Current.Server.MapPath(tmpPath), fileName);

                    var buf = stream.ToArray();
                    using (var fs = new FileStream(tmpFile, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(buf, 0, buf.Length);
                        fs.Flush();
                        fs.Close();
                        fs.Dispose();
                    }
                }
                else
                {
                    string userAgent = System.Web.HttpContext.Current.Request.UserAgent.ToUpper();
                    if (userAgent.IndexOf("FIREFOX") < 0) fileName = System.Web.HttpContext.Current.Server.UrlEncode(fileName);
                    else fileName = System.Web.HttpContext.Current.Server.UrlDecode(fileName);
                    var buf = stream.ToArray();//stream.GetBuffer()

                    HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
                    //如果去掉.xls 将导致火狐下 导出文件后缀丢失
                    HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}.xlsx", fileName));
                    HttpContext.Current.Response.Clear();
                    HttpContext.Current.Response.BinaryWrite(buf);
                    HttpContext.Current.Response.Flush();
                    HttpContext.Current.Response.Close();
                }

                stream.Close();
                stream.Dispose();
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return string.Format("{0}/{1}.xlsx", tmpPath, fileName);

            #endregion
        }


        /// <summary>
        /// 设置报表单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="cellStyle"></param>
        /// <param name="columnStyle"></param>
        private void SetCellStyle(out ICellStyle cellStyle, out ICellStyle columnStyle, out ICellStyle headStyle)
        {
            cellStyle = columnStyle = headStyle = null;
            if (workbook == null) return;

            //单元格样式 
            cellStyle = workbook.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Left;
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.LeftBorderColor = cellStyle.RightBorderColor
                = cellStyle.TopBorderColor = cellStyle.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");//以文本形式存储

            //列样式
            columnStyle = workbook.CreateCellStyle();
            columnStyle.CloneStyleFrom(cellStyle);
            columnStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            columnStyle.FillPattern = FillPattern.SolidForeground;
            //columnStyle.WrapText = true;//自动换行
            columnStyle.Alignment = HorizontalAlignment.Center;
            columnStyle.VerticalAlignment = VerticalAlignment.Center;

            //顶部统计样式
            headStyle = workbook.CreateCellStyle();
            headStyle.CloneStyleFrom(columnStyle);
            headStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
            headStyle.Alignment = HorizontalAlignment.Left;

            //DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            //dsi.Company = "xxxx.com";
            //workbook.DocumentSummaryInformation = dsi;
            //SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            //si.Subject = "xxxx.com";
            //workbook.SummaryInformation = si;
        }

        /// <summary>
        /// 单元格纵向合并
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="mergedCells">合并的列集合</param>
        /// <param name="headRowIndex">报表标题列 行索引</param>
        private void MergedCells(ISheet sheet, string mergedCells, int columnsRowIndex)
        {
            if (string.IsNullOrEmpty(mergedCells)) return;

            if (!mergedCells.StartsWith(",")) mergedCells = "," + mergedCells;
            if (!mergedCells.EndsWith(",")) mergedCells += ",";

            //需要合并列数量
            int iNumA = 0, iNumB = mergedCells.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Length;
            int colIndex = 0, rowIndex = 0, rowIndexMin = 0;
            string strValue = string.Empty;

            //获取标题行
            IRow rowHead = null;
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                if (rowIndex == columnsRowIndex)
                {
                    rowHead = (HSSFRow)rows.Current;
                    break;
                }
                rowIndex++;
            }
            rowIndex = 0;
            rowHead.GetCell(0);
            IRow irow = null;//内容行
            for (colIndex = 0; colIndex < rowHead.LastCellNum; colIndex++) //foreach (ICell cell in rowHead.Cells)//遍历列
            {
                if (iNumB == iNumA) break;//避免不必要的遍历
                if (!mergedCells.Contains("," + rowHead.GetCell(colIndex).StringCellValue + ","))
                {
                    colIndex++;
                    continue;
                }
                rowIndexMin = columnsRowIndex;//标题列索引
                strValue = string.Empty;
                //直接从标题行下一行开始遍历
                for (rowIndex = columnsRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    irow = sheet.GetRow(rowIndex);//获取标题行下一行
                    if (rowIndexMin == columnsRowIndex)
                    {
                        strValue = irow.GetCell(colIndex).StringCellValue;
                        rowIndexMin = rowIndex;
                    }
                    else if (strValue != irow.GetCell(colIndex).StringCellValue)
                    {
                        if ((rowIndex - 1) >= rowIndexMin)
                            sheet.AddMergedRegion(new CellRangeAddress(rowIndexMin, rowIndex - 1, colIndex, colIndex));
                        strValue = irow.GetCell(colIndex).StringCellValue;
                        rowIndexMin = rowIndex;
                    }
                    else if (rowIndex == (sheet.LastRowNum))
                    {
                        if (strValue == irow.GetCell(colIndex).StringCellValue)
                        {
                            if ((rowIndex - 1) >= rowIndexMin)
                                sheet.AddMergedRegion(new CellRangeAddress(rowIndexMin, rowIndex, colIndex, colIndex));
                        }
                    }
                }
                //colIndex++;
                iNumA++;//合并列次数+1
            }
        }

        #endregion

        #region EXCEL导入

        /// <summary>
        /// 导入EXCEL数据
        /// </summary>
        /// <param name="listColumns">EXCEL 列对应表列集合</param>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        public DataTable ImportExcelToTable(List<ExcelColumns> listColumns, string filePath)
        {
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
                ISheet sheet = workbook.GetSheetAt(0);
                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

                DataTable dtTable = new DataTable();
                //添加表列
                foreach (ExcelColumns col in listColumns)
                {
                    dtTable.Columns.Add(col.FieldName);
                }

                int rowIndex = 0, colIndex = 0;
                IRow headRow = null;
                ICell icell = null;
                while (rows.MoveNext())
                {
                    if (rowIndex == 0)
                    {
                        headRow = (HSSFRow)rows.Current;//获取Excel 标题行
                        rowIndex++;
                        continue;
                    }
                    IRow row = (HSSFRow)rows.Current;
                    DataRow newRow = dtTable.NewRow();
                    colIndex = 0;
                    for (colIndex = 0; colIndex < headRow.LastCellNum; colIndex++) //foreach (ICell cell in headRow.Cells)
                    {
                        icell = headRow.GetCell(colIndex);
                        //只允许导出集合 listColumns 包含的列，且列队顺序一致
                        if (!icell.StringCellValue.Equals(listColumns[colIndex].FieldTitle))
                            newRow[colIndex] = row.Cells[colIndex].StringCellValue;
                        else
                            newRow[colIndex] = null;

                        colIndex++;
                    }

                    rowIndex++;
                    dtTable.Rows.Add(newRow);
                }

                return dtTable;
            }
        }

        #endregion
    }

    #region 报表自定义列

    /// <summary>
    /// 通过委托调用导出列处理方法（有且仅有一个string参数的方法，均可使用该委托，方法名不做限制）（导出使用）
    /// </summary>
    /// <param name="value"></param>
    public delegate string ColumnAction(string value);

    /// <summary>
    /// 通过委托调用导出列校验方法（导入使用）
    /// </summary>
    /// <param name="value">校验的内容</param>
    /// <param name="errMsg">错误信息</param>
    /// <returns></returns>
    public delegate string ColumnCheck(string value, out string errMsg);

    /// <summary>
    /// 导出报表自定义列
    /// </summary>
    public class ExcelColumns
    {
        /// <summary>
        /// 通过构造函数直接赋值
        /// </summary>
        /// <param name="fieldName">对应表列名</param>
        /// <param name="fieldTitle">对应Excel标题</param>
        public ExcelColumns(string fieldName = "", string fieldTitle = "")
        {
            this.FieldName = fieldName;
            this.FieldTitle = fieldTitle;
        }

        /// <summary>
        /// 对应表列名
        /// </summary>
        public string FieldName;

        /// <summary>
        /// 对应Excel标题
        /// </summary>
        public string FieldTitle;

        /// <summary>
        /// 单元格数据格式（Str:字符类型(默认)，Int:整数，Dec:小数类型10.00，Date:日期类型 2014年01月01日,DateTime:日期时间类型 2014年01月01日 20:13:14）
        /// </summary>
        public string FieldType;

        /// <summary>
        /// 是否纵向合并单元格（默认 false）
        /// </summary>
        public bool IsMerged;

        /// <summary>
        /// 是否需要统计 累加计算（默认 false）
        /// </summary>
        public bool IsAccount;

        /// <summary>
        /// 声明处理列委托对象（导出使用）
        /// </summary>
        public ColumnAction Action;

        /// <summary>
        /// 声明校验列委托对象（导入使用）
        /// </summary>
        /// <returns></returns>
        public ColumnCheck Check;
    }

    #endregion
}