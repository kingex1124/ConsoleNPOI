using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.MyHelper.Model
{
    public class ExcelExportReportSource
    {

		/// <summary>
		/// 新增資料表頭上方資訊 如果沒有請設定為null 如果要加入一行空白列請新增null至清單裡
		/// </summary>
		public List<List<NameValuePair<string>>> DataHeaderList { get; set; }

		/// <summary>
		/// 資料表來源
		/// </summary>
		public DataTable Data { get; set; }

		/// <summary>
		/// 設定欄位格式
		/// </summary>
		public List<NameValuePair<string>> DataFormatList { get; set; }
	}
}
