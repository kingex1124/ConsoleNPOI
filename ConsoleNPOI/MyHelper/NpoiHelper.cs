using ConsoleNPOI.MyHelper.Model;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.MyHelper
{
    public static class NpoiHelper
    {
        /// <summary>
        /// 匯出Excel
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

                throw;
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
    }
}
