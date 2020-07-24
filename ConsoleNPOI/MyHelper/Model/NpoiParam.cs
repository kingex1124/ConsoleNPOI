using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.MyHelper.Model
{
    public class NpoiParam<T>
    {
        private string[] _sheetName = new string[1];
        private bool? _showHeader;
        private bool? _isAutoFit;
        private FontStyle _headerFontStyle;
        private FontStyle _dataFontStyle;

        /// <summary>
        /// 請用 HSSFWorkbook 或 XSSFWorkbook 實體化 IWorkbook
        /// 必填
        /// </summary>
        public IWorkbook Workbook { get; set; }

        /// <summary>
        /// 最後excel檔要被寫出到哪裡
        /// 必填
        /// </summary>
        public string FileFullName { get; set; }

        /// <summary>
        /// 資料
        /// 必填
        /// </summary>
        public List<T>[] Data { get; set; }
        
        /// <summary>
        /// 欲新增(或已存在)的 Sheet Name
        /// 預設為Sheet1
        /// </summary>
        public string[] SheetName 
        { 
            get 
            {
                for (int i = 0; i < _sheetName.Length; i++)
                    _sheetName[i] = string.IsNullOrWhiteSpace(_sheetName[i]) ? "Sheet" + i : _sheetName[i];
                return _sheetName;
            }
            set { _sheetName = value; } 
        }

        /// <summary>
        /// 與 Excel 檔間的欄位對應
        /// 必填
        /// </summary>
        public List<ColumnMapping>[] ColumnMapping { get; set; }
        
        /// <summary>
        /// 是否excel要畫表頭 (預設畫表頭 = true)
        /// </summary>
        public bool ShowHeader
        {
            get { return _showHeader ?? true; }
            set { _showHeader = value; }
        }

        /// <summary>
        /// 是否自動調整欄寬 (預設不自動調整欄寬 = false)
        /// </summary>
        public bool IsAutoFit
        {
            get { return _isAutoFit ?? false; }
            set { _isAutoFit = value; }
        }

        /// <summary>
        /// 決定表頭文字預設格式
        /// </summary>
        public FontStyle HeaderFontStyle
        {
            get { return _headerFontStyle ?? new FontStyle(); }
            set { _headerFontStyle = value; }
        }

        /// <summary>
        /// 決定資料文字預設格式
        /// </summary>
        public FontStyle DataFontStyle
        {
            get { return _dataFontStyle ?? new FontStyle(); }
            set { _dataFontStyle = value; }
        }
    }
}
