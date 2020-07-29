using ConsoleNPOI.MyHelper.Model;
using Microsoft.Practices.EnterpriseLibrary.ExceptionHandling;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Crypto.Tls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.MyHelper
{
    public static class NpoiHelper
    {
        #region 匯出excel

        #region ListToExcel

        /// <summary>
        /// 透過List匯出Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="param"></param>
        public static void ExportExcel<T>(NpoiParam<T> param)
        {
            try
            {
                //依情況決定要建新的 Sheet 或是用舊的 (即來自範本) 主要由SheetName來決定有幾個Sheet
                ISheet[] sheets = GetSheet(param);

                for (int i = 0; i < sheets.Length; i++)
                {
                    if (param.Data[i].Any())
                        SetSheetValue(ref param, ref sheets[i], i);
                    else
                    {
                        // 若沒資料在起點寫入No Data !
                        sheets[i].CreateRow(0);
                        sheets[i].GetRow(0).CreateCell(0).SetCellValue("No Data !");
                    }
                }

                Export(param.Workbook, param.FileFullName);
            }
            catch (Exception ex)
            {
                GetCustomErrorCodeDescription(ex);
            }
        }

        /// <summary>
        /// 取得Sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="param"></param>
        /// <returns></returns>
        private static ISheet[] GetSheet<T>(NpoiParam<T> param)
        {
            ISheet[] sheets = new ISheet[param.SheetName.Length];
            for (int i = 0; i < param.SheetName.Length; i++)
            {
                // 在 workbook 中以 sheet name 尋找 是否找得到sheet
                if (param.Workbook.GetSheet(param.SheetName[i]) == null)
                {
                    ISheet sheetTmp = param.Workbook.CreateSheet(param.SheetName[i]);
                    sheetTmp = CreateColumn(param, i);

                    sheets[i] = sheetTmp;
                }
                else
                    sheets[i] = param.Workbook.GetSheet(param.SheetName[i]); // 找得到即為要塞值的目標
            }

            return sheets;
        }

        /// <summary>
        /// 建立表頭內容
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="param"></param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        private static ISheet CreateColumn<T>(NpoiParam<T> param, int sheetIndex)
        {
            var sheet = param.Workbook.GetSheet(param.SheetName[sheetIndex]);
            // 建立表頭Row
            sheet.CreateRow(0);
            ICellStyle columnStyle = GetBaseCellStyle(param.Workbook, param.HeaderFontStyle);

            if (param.ShowHeader)
            {
                for (int j = 0; j < param.ColumnMapping[sheetIndex].Count; j++)
                {
                    // 建立欄位
                    sheet.GetRow(0).CreateCell(j);
                    // 設定欄位Style
                    sheet.GetRow(0).GetCell(j).CellStyle = columnStyle;
                    // 給欄位值
                    sheet.GetRow(0).GetCell(j).SetCellValue(param.ColumnMapping[sheetIndex][j].ExcelColumnName);
                }
            }
            return sheet;
        }

        /// <summary>
        /// 取得基本style設定
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fontStyle"></param>
        /// <returns></returns>
        private static ICellStyle GetBaseCellStyle(IWorkbook workbook, FontStyle fontStyle)
        {
            //建立欄位Style物件
            ICellStyle cellStyle = workbook.CreateCellStyle();

            //畫線
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;

            cellStyle.Alignment = HorizontalAlignment.Center; //水平置中

            //預設字型大小
            IFont font = workbook.CreateFont();
            font.FontName = (fontStyle.FontName == null) ? "新細明體" : fontStyle.FontName;
            font.FontHeightInPoints = (fontStyle.FontHeightInPoints == null) ? (short)12 : fontStyle.FontHeightInPoints.Value;
            cellStyle.SetFont(font);

            return cellStyle;
        }

        /// <summary>
        /// 設定sheet裡面的資料
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="param"></param>
        /// <param name="sheet"></param>
        private static void SetSheetValue<T>(ref NpoiParam<T> param, ref ISheet sheet, int sheetIndex)
        {
            int line = 1;
            
            //根據標題列先處理所有Style (對Npoi來說 '創建'Style在workbook中是很慢的操作 作越少次越好 絕對不要foreach在塞每行列實際資料時重覆作 只通通在標題列做一次就好)
            ICellStyle[] cellStyleArr = InitialColumnStyle(param.Workbook, param.ColumnMapping, param.DataFontStyle, sheetIndex);

            foreach (var item in param.Data[sheetIndex])
            {
                sheet.CreateRow(line);

                for (int i = 0; i < param.ColumnMapping[sheetIndex].Count; i++)
                {
                    // 建立欄位
                    var cell = sheet.GetRow(line).CreateCell(i);

                    // 綁定欄位Style
                    cell.CellStyle = cellStyleArr[i];

                    // 給欄位值 reflection取值
                    string value = GetValue(item, param.ColumnMapping[sheetIndex], i);
                    // 幫cell填值
                    SetCellValue(value, ref cell, param.ColumnMapping[sheetIndex][i].DataType);
                }

                line++;
            }

            if (param.IsAutoFit)
                for (int i = 0; i < param.ColumnMapping[sheetIndex].Count; i++)
                    sheet.AutoSizeColumn(i);
        }



        /// <summary>
        /// 設定資料欄位中的Style
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="columnMapping"></param>
        /// <param name="fontStyle"></param>
        /// <returns></returns>
        private static ICellStyle[] InitialColumnStyle(IWorkbook workbook, List<ColumnMapping>[] columnMapping, FontStyle fontStyle,int sheetIndex)
        {
            int cellCount = columnMapping[sheetIndex].Count;

            ICellStyle[] styleArr = new ICellStyle[cellCount];

            for (int i = 0; i < cellCount; i++)
            {
                //取通用格式
                ICellStyle cellStyle = GetBaseCellStyle(workbook, fontStyle);

                //處理格式輸出
                if (!String.IsNullOrWhiteSpace(columnMapping[sheetIndex][i].Format))
                    cellStyle.DataFormat = GetCellFormat(workbook, columnMapping[sheetIndex][i].Format);

                styleArr[i] = cellStyle;
            }

            return styleArr;
        }

        /// <summary>
        /// 取得欄位格式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        private static short GetCellFormat(IWorkbook workbook, string format)
        {
            IDataFormat dataFormat = workbook.CreateDataFormat();
            return dataFormat.GetFormat(format);
        }

        /// <summary>
        ///  取得值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="list"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        private static string GetValue<T>(T data, List<ColumnMapping> columnMapping, int i)
        {
            var fieldName = columnMapping[i].ModelFieldName;
            var prop = typeof(T).GetProperties().Where(q => q.Name == fieldName).First();
            var value = prop.GetValue(data, null);

            return (value == null) ? "" : value.ToString();
        }

        /// <summary>
        /// 設定欄位值、屬性
        /// </summary>
        /// <param name="value"></param>
        /// <param name="cell"></param>
        /// <param name="dataType"></param>
        private static void SetCellValue(string value, ref ICell cell, NpoiDataType dataType)
        {
            switch (dataType)
            {
                case NpoiDataType.String:
                    if (!string.IsNullOrWhiteSpace(value)) 
                        cell.SetCellValue(value);
                    break;
                case NpoiDataType.Int:
                    if (!string.IsNullOrWhiteSpace(value)) 
                        cell.SetCellValue(Convert.ToDouble(value));
                    break;
                case NpoiDataType.Double:
                    if (!string.IsNullOrWhiteSpace(value))
                        cell.SetCellValue(Convert.ToDouble(value));
                    break;
                case NpoiDataType.DateTime:
                    if (!string.IsNullOrWhiteSpace(value)) 
                        cell.SetCellValue(Convert.ToDateTime(value));
                    break;
                case NpoiDataType.Bool:
                    if (!string.IsNullOrWhiteSpace(value))
                        cell.SetCellValue(Convert.ToBoolean(value));
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 將資料寫入電腦
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fileFullName"></param>
        private static void Export(IWorkbook workbook, string fileFullName)
        {
            //若檔案已存在先刪除
            if (File.Exists(fileFullName))
                File.Delete(fileFullName);

            using (FileStream targetFs = File.Create(fileFullName))
            {
                workbook.Write(targetFs);
                workbook = null;
            }
        }

        #region ListToExcelBinary

        /// <summary>
        /// 取得Excel二進位資料
        /// </summary>
        /// <typeparam name="T">資料型別</typeparam>
        /// <param name="queryData">資料</param>
        /// <param name="sheetName">sheet名稱</param>
        /// <param name="TitleName">Head欄位名稱</param>
        /// <param name="columWidth">欄位寬度</param>
        /// <returns></returns>
        public static byte[] GetExcelBinary<T>(IEnumerable<T> queryData,string sheetName, string[] TitleName, int[] columWidth)
        {
            XSSFWorkbook excel = new XSSFWorkbook();

            ISheet sheet;

            if (!string.IsNullOrEmpty(sheetName))
                sheet = excel.CreateSheet(sheetName);
            else
                sheet = excel.CreateSheet("sheetName");

            List<T> resultData = new List<T>();
            resultData = queryData.ToList<T>();

            sheet.CreateRow(0);

            for (int i = 0; i < columWidth.Count<int>(); i++)
                sheet.SetColumnWidth(i, columWidth[i]);

            for (int i = 0; i < TitleName.Count<string>(); i++)
                sheet.GetRow(0).CreateCell(i).SetCellValue(TitleName[i]);

            for (int i = 1; i <= resultData.Count<T>(); i++)
            {
                sheet.CreateRow(i);
                for (int j = 0; j < typeof(T).GetProperties().Count<PropertyInfo>(); j++)
                    sheet.GetRow(i).CreateCell(j).SetCellValue(typeof(T).GetProperty(typeof(T).GetProperties()[j].Name).GetValue(resultData[i - 1]).ToString());
            }
            MemoryStream MS = new MemoryStream();
            excel.Write(MS);
            return MS.ToArray();
        }

        #endregion

        #endregion

        #region DataTableToExcel

        /// <summary>
        /// 透過DataTable 匯出Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sheetName"></param>
        /// <param name="saveFilePath"></param>
        public static void DataTableToExcelFile(DataTable dt, string sheetName, string saveFilePath)
        {
            //建立Excel 2003檔案
            //IWorkbook wb = new HSSFWorkbook();
            //ISheet ws;

            ////建立Excel 2007檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            if (!string.IsNullOrEmpty(sheetName))
                ws = wb.CreateSheet(sheetName);
            else if (dt.TableName != string.Empty)
                ws = wb.CreateSheet(dt.TableName);
            else
                ws = wb.CreateSheet("Sheet1");

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ws.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                    ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
            }

            FileStream file = new FileStream(saveFilePath, FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();
        }

        #endregion

        #endregion

        #region 匯入Excel

        #region ExcelToList

        /// <summary>
        /// 匯入Excel
        /// 取得List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">取得檔案路徑</param>
        /// <param name="sheetIndex">取得Sheet的Index</param>
        /// <returns></returns>
        public static List<T> ImportExcel<T>(string filePath, ref string errorMsg, int sheetIndex = 0)
        {
            var result = new List<T>();

            errorMsg = string.Empty;

            IWorkbook wookbook = null;

            string extension = Path.GetExtension(filePath);

            extension = string.IsNullOrWhiteSpace(extension) ? string.Empty : extension.ToLower();

            try
            {
                FileStream fs = File.OpenRead(filePath);
                if (extension == ".xls" || extension == ".xlsx") // 2003
                    wookbook = WorkbookFactory.Create(fs);
                else
                {
                    errorMsg = "不支援的檔案格式!";
                    return null;
                }

                fs.Close();

                ISheet sheet = wookbook.GetSheetAt(sheetIndex);

                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);

                    if (row != null)
                    {
                        T t = Activator.CreateInstance<T>();
                        for (int j = 0; j < row.LastCellNum; j++)
                            typeof(T).GetProperties()[j].SetValue(t, GetCellValue(row.GetCell(j)));

                        result.Add(t);
                    }
                }

            }
            catch (Exception ex)
            {
                errorMsg = GetCustomErrorCodeDescription(ex, "匯入失敗!", true);
                return null;
            }

            return result;
        }

        /// <summary>
        /// 將取得的值轉為excell中的欄位型別
        /// 並將值轉為string，除了DateTime之外(為了方便在外部存取時調整格式)
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static object GetCellValue(ICell cell)
        {
            object value = null;
            try
            {
                if (cell.CellType != CellType.Blank)
                {
                    //格式最後皆為字串 除了日期不用
                    switch (cell.CellType)
                    {
                        case CellType.Numeric:
                            // Date comes here
                            if (DateUtil.IsCellDateFormatted(cell))
                                value = cell.DateCellValue;//.ToString("yyyy/MM/dd/HH:mm:ss");
                            else // Numeric(double) type 數字一律轉為字串，必要時在處理
                                value = cell.NumericCellValue.ToString();
                            break;
                        case CellType.Boolean:
                            // Boolean type
                            value = cell.BooleanCellValue.ToString();
                            break;
                        case CellType.Formula:
                            value = cell.CellFormula;
                            break;
                        default:
                            // String type
                            value = cell.StringCellValue;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                value = "";
                GetCustomErrorCodeDescription(ex, "轉型失敗", true);
            }
            return value;
        }

        #endregion

        #region ExcelToDataTable

        /// <summary>
        /// 匯入Excel
        /// 取得DataTable
        /// </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <param name="errorMsg">錯誤資訊</param>
        /// <param name="headerRowIndex">從第幾行開始取資料</param>
        /// <returns></returns>
        public static DataTable ImportExcelToDataTable(string filePath, ref string errorMsg, int headerRowIndex = 0)
        {
            errorMsg = string.Empty;

            var fileExtension = Path.GetExtension(filePath);

			fileExtension = string.IsNullOrWhiteSpace(fileExtension) ? string.Empty : fileExtension.ToLower();

            try
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    DataTable excelTable = null;
                    byte[] bytes = new byte[file.Length];
                    file.Read(bytes, 0, (int)file.Length);
                    file.Position = 0;

                    if (fileExtension == ".xls" || fileExtension == ".xlsx")
                    {
                        //註記: RenderDataTableFromExcel無法讀取CSV檔案
                        excelTable = RenderDataTableFromExcel(file, 0, headerRowIndex);
                    }
                    else
                    {
                        errorMsg = "不支援的檔案格式!";
                        return null;
                    }
                    return excelTable;
                }
            }
            catch (Exception ex)
            {
                errorMsg = GetCustomErrorCodeDescription(ex, "匯入失敗!", true);
                return null;
            }
        }

        /// <summary>
        /// 轉出Excel資料至DataTable
        /// </summary>
        /// <param name="excelFileStream"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="headerRowIndex">標題列位置，-1表示無標題列</param>
        /// <param name="cellCount">欄位數，若有指定標題列，則以標題列的欄位數為主，-1表示不指定欄位數</param>
        /// <returns></returns>
        public static DataTable RenderDataTableFromExcel(Stream excelFileStream, int sheetIndex = 0, int headerRowIndex = 0, int cellCount = -1)
        {
            var workbook = WorkbookFactory.Create(excelFileStream);
            //指定的Sheet
            var sheet = workbook.GetSheetAt(sheetIndex);
            //指定為Header的Row

            var table = ReadDataTableFromSheet(sheet, headerRowIndex, cellCount);
            //建議使用using開啟檔案，或是外部呼叫結束時控制關閉，不要在單純讀取資料的地方關閉
            //excelFileStream.Close(); //hint:
            workbook = null;
            sheet = null;
            return table;
        }

        /// <summary>
		/// 讀取Excel的Sheet轉換為DataTable
		/// </summary>
		private static DataTable ReadDataTableFromSheet(ISheet sheet, int headerRowIndex, int cellCount = -1)
        {
            var table = new DataTable();

            IRow headerRow = null;
            if (headerRowIndex != -1)
                headerRow = sheet.GetRow(headerRowIndex);

            var firstRow = sheet.FirstRowNum + headerRowIndex + 1;
            if (headerRow == null)
            {
                headerRow = sheet.GetRow(0);
                firstRow = sheet.FirstRowNum;
            }
            else
                cellCount = headerRow.LastCellNum;

            if (cellCount == -1)
                cellCount = headerRow.LastCellNum;
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                if (firstRow == 0)
                {
                    var column = new DataColumn("Col" + i.ToString());
                    table.Columns.Add(column);
                }
                else
                {
                    var column = new DataColumn(headerRow.GetCell(i).ToString());
                    table.Columns.Add(column);
                }
            }
            for (int i = firstRow; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null || row.Cells.Count == 0)
                    continue;

                // 是否為空白Row
                bool isEmptyRow = true;
                var dataRow = table.NewRow();
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    var cell = row.GetCell(j);
                    if (cell != null)
                    {
                        var content = cell.ToString();
                        dataRow[j] = content;

                        if (!string.IsNullOrWhiteSpace(content))
                            isEmptyRow = false;
                    }
                }

                if (!isEmptyRow)
                    table.Rows.Add(dataRow);
            }
            return table;
        }

        #endregion

        #endregion

        #region 錯誤處理

        /// <summary>
        /// 取得自訂錯誤描述
        /// </summary>
        /// <param name="ex">Exception ex</param>
        /// <param name="defineMsg">
        /// 預設錯誤資訊(當錯誤不在其中，則顯示[預設錯誤資訊+ex.Message]，
        /// 例如:匯入失敗 + [Message])
        /// </param>
        /// <param name="replace">
        /// 若選是，找不到代碼時，則不會帶出ex.Message，會直接回傳defineMsg
        /// (因為有些UI介面不要顯示詳細資訊)
        /// </param>
        /// <returns></returns>
        public static string GetCustomErrorCodeDescription(Exception ex, string defineMsg = "", bool replace = false)
        {
            var errorHeader = (string.IsNullOrWhiteSpace(defineMsg)
                ? string.Empty
                : defineMsg);
            var errorMsg = errorHeader + (replace ? string.Empty : ex.Message);
            var errorCode = GetSystemErrorCode(ex);
            switch (errorCode)
            {
                //Ref: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681382%28v=vs.85%29.aspx
                // MSDN System Error Codes(錯誤代碼表)
                // ERROR_SHARING_VIOLATION(32): The process cannot access the file because it is being used by another process.
                // ERROR_LOCK_VIOLATION(33): The process cannot access the file because another process has locked a portion of the file.
                case 32:
                case 33:
                    errorMsg = errorHeader + "檔案已被開啟或鎖定!";
                    break;
                case 87:
                    errorMsg = errorHeader + "檔案路徑不得空白!";
                    break;
                case 6434:
                    errorMsg = errorHeader + "Excel欄位名稱重複!";
                    break;
                case 5378:
                    errorMsg = errorHeader + "欄位格式有問題!";
                    break;
                case 16387:
                    errorMsg = errorHeader + "請勿插入空白欄位!";
                    break;

            }
            if (!string.IsNullOrEmpty(errorMsg))
                ExceptionPolicy.HandleException(ex, "Default Policy");
            else
                throw new Exception("Import File Error", ex);
            return errorMsg;
        }

        public static int GetSystemErrorCode(Exception ex)
        {
            //Ref: https://msdn.microsoft.com/zh-tw/library/windows/desktop/ms690088%28v=vs.85%29.aspx
            // MSDN HRESULT Define
            //HRESULT的格式定義，後面16碼為Error Code
            //65535轉2進位=>1111111111111111再and，可以取出所需代碼
            return Marshal.GetHRForException(ex) & 65535;
        }

        #endregion
    }
}
