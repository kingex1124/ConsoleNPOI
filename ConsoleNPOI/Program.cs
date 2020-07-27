using ConsoleNPOI.MyHelper;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ConsoleNPOI.MyHelper.SampleTest2;

namespace ConsoleNPOI
{
	class Program
	{
		static void Main(string[] args)
		{
			//Program pg = new Program();
			//pg.CreateExcelFile();

			SampleTest2 s = new SampleTest2();
			string errMessage = string.Empty;
			var re = NpoiHelper.ImportExcel<Sample>(@"C:\Users\011714\Desktop\result.xlsx", ref errMessage);
		
		}

		#region 寫入Excel

		#region 參考一

		//範例一，簡單產生Excel檔案的方法
		private void CreateExcelFile()
		{
			//建立Excel 2003檔案
			IWorkbook wb = new HSSFWorkbook();

			// Class 為 Sheet Name
			ISheet ws = wb.CreateSheet("Class");

			////建立Excel 2007檔案
			//IWorkbook wb = new XSSFWorkbook();
			//ISheet ws = wb.CreateSheet("Class");

			// header
			ws.CreateRow(0);//第一行為欄位名稱
			ws.GetRow(0).CreateCell(0).SetCellValue("name");
			ws.GetRow(0).CreateCell(1).SetCellValue("score");
		

			// data
			ws.CreateRow(1);//第二行之後為資料
			ws.GetRow(1).CreateCell(0).SetCellValue("abey");
			ws.AutoSizeColumn(0);
			ws.GetRow(1).CreateCell(1).SetCellValue(85);
			ws.AutoSizeColumn(1);

			ws.CreateRow(2);
			ws.GetRow(2).CreateCell(0).SetCellValue("tina111111111111111111111");
			ws.AutoSizeColumn(0);
			ws.GetRow(2).CreateCell(1).SetCellValue(82);
			ws.AutoSizeColumn(1);

			ws.CreateRow(3);
			ws.GetRow(3).CreateCell(0).SetCellValue("boi");
			ws.AutoSizeColumn(0);
			ws.GetRow(3).CreateCell(1).SetCellValue(84);
			ws.AutoSizeColumn(1);
			ws.CreateRow(4);

			ws.GetRow(4).CreateCell(0).SetCellValue("hebe22222");
			ws.AutoSizeColumn(0);
			ws.GetRow(4).CreateCell(1).SetCellValue(86);
			ws.AutoSizeColumn(1);

			ws.CreateRow(5);
			ws.GetRow(5).CreateCell(0).SetCellValue("paul");
			ws.AutoSizeColumn(0);
			ws.GetRow(5).CreateCell(1).SetCellValue(82);
			ws.AutoSizeColumn(1);


			FileStream file = new FileStream(@"C:\Users\011714\Desktop\TEST\npoi.xls", FileMode.Create);//產生檔案
			wb.Write(file);
			file.Close();
		}

		//範例二，DataTable轉成Excel檔案的方法
		private void DataTableToExcelFile(DataTable dt)
		{
			//建立Excel 2003檔案
			IWorkbook wb = new HSSFWorkbook();
			ISheet ws;

			////建立Excel 2007檔案
			//IWorkbook wb = new XSSFWorkbook();
			//ISheet ws;

			if (dt.TableName != string.Empty)
			{
				ws = wb.CreateSheet(dt.TableName);
			}
			else
			{
				ws = wb.CreateSheet("Sheet1");
			}

			ws.CreateRow(0);//第一行為欄位名稱
			for (int i = 0; i < dt.Columns.Count; i++)
			{
				ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
			}

			for (int i = 0; i < dt.Rows.Count; i++)
			{
				ws.CreateRow(i + 1);
				for (int j = 0; j < dt.Columns.Count; j++)
					ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
			}

			FileStream file = new FileStream(@"C:\Users\011714\Desktop\TEST\npoi.xls", FileMode.Create);//產生檔案
			wb.Write(file);
			file.Close();
		}

		#endregion

		#region 參考二

		public void WriteToExcel(string filePath)
		{
			//創建工作薄  
			IWorkbook wb;
			string extension = System.IO.Path.GetExtension(filePath);

			//根據指定的文件格式創建對應的類
			if (extension.Equals(".xls"))
			{
				// 2003
				wb = new HSSFWorkbook();
			}
			else
			{
				// 2007
				wb = new XSSFWorkbook();
			}
			ICellStyle style1 = wb.CreateCellStyle();//樣式
			style1.Alignment = HorizontalAlignment.Left;//文字水平對齊方式
			style1.VerticalAlignment = VerticalAlignment.Center;//文字垂直對齊方式
																				  //設置邊框
			style1.BorderBottom = BorderStyle.Thin;
			style1.BorderLeft = BorderStyle.Thin;
			style1.BorderRight = BorderStyle.Thin;
			style1.BorderTop = BorderStyle.Thin;

			style1.WrapText = true;//自動換行

			ICellStyle style2 = wb.CreateCellStyle();//樣式
			IFont font1 = wb.CreateFont();//字體
			font1.FontHeightInPoints = 12; // 字體尺寸
			font1.FontName = "楷體";
			font1.Color = HSSFColor.Red.Index;//字體顏色
			font1.Boldweight = (short)FontBoldWeight.Normal;//字體加粗樣式
			style2.SetFont(font1);//樣式裏的字體設置具體的字體樣式

			//設置背景色
			style2.FillForegroundColor = HSSFColor.Yellow.Index;
			style2.FillPattern = FillPattern.SolidForeground;
			style2.FillBackgroundColor = HSSFColor.Yellow.Index;
			style2.Alignment = HorizontalAlignment.Left;//文字水平對齊方式
			style2.VerticalAlignment = VerticalAlignment.Center;//文字垂直對齊方式

			ICellStyle dateStyle = wb.CreateCellStyle();//樣式
			dateStyle.Alignment = HorizontalAlignment.Left;//文字水平對齊方式
			dateStyle.VerticalAlignment = VerticalAlignment.Center;//文字垂直對齊方式

			//設置數據顯示格式
			IDataFormat dataFormatCustom = wb.CreateDataFormat();

			dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");

			//創建一個表單
			ISheet sheet = wb.CreateSheet("Sheet0");
			//設置列寬
			int[] columnWidth = { 10, 10, 20, 10 };

			for (int i = 0; i < columnWidth.Length; i++)
			{
				//設置列寬度，256*字符數，因為單位是1/256個字符
				sheet.SetColumnWidth(i, 256 * columnWidth[i]);
			}

			//測試數據
			int rowCount = 3, columnCount = 4;
			object[,] data = {
			{"列0", "列1", "列2", "列3"},
			{"", 400, 5.2, 6.01},
			{"", true, "2014-07-02", DateTime.Now}
			//日期可以直接傳字符串，NPOI會自動識別
			//如果是DateTime類型，則要設置CellStyle.DataFormat，否則會顯示為數字
			};

			IRow row;
			ICell cell;

			for (int i = 0; i < rowCount; i++)
			{
				row = sheet.CreateRow(i);//創建第i行
				for (int j = 0; j < columnCount; j++)
				{
					cell = row.CreateCell(j);//創建第j列
					cell.CellStyle = j % 2 == 0 ? style1 : style2;
					//根據數據類型設置不同類型的cell
					object obj = data[i, j];
					SetCellValue(cell, data[i, j]);
					//如果是日期，則設置日期顯示的格式
					if (obj.GetType() == typeof(DateTime))
					{
						cell.CellStyle = dateStyle;
					}
					//如果要根據內容自動調整列寬，需要先setCellValue再調用
					//sheet.AutoSizeColumn(j);
				}
			}

			//合並單元格，如果要合並的單元格中都有數據，只會保留左上角的
			//CellRangeAddress(0, 2, 0, 0)，合並0-2行，0-0列的單元格
			CellRangeAddress region = new CellRangeAddress(0, 2, 0, 0);
			sheet.AddMergedRegion(region);
			try
			{
				FileStream fs = File.OpenWrite(filePath);
				wb.Write(fs);//向打開的這個Excel文件中寫入表單並保存。  
				fs.Close();
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
			}
		}

		//根據數據類型設置不同類型的cell
		public static void SetCellValue(ICell cell, object obj)
		{
			if (obj.GetType() == typeof(int))
			{
				cell.SetCellValue((int)obj);
			}
			else if (obj.GetType() == typeof(double))
			{
				cell.SetCellValue((double)obj);
			}
			else if (obj.GetType() == typeof(IRichTextString))
			{
				cell.SetCellValue((IRichTextString)obj);
			}
			else if (obj.GetType() == typeof(string))
			{
				cell.SetCellValue(obj.ToString());
			}
			else if (obj.GetType() == typeof(DateTime))
			{
				cell.SetCellValue((DateTime)obj);
			}
			else if (obj.GetType() == typeof(bool))
			{
				cell.SetCellValue((bool)obj);
			}
			else
			{
				cell.SetCellValue(obj.ToString());
			}
		}

		#endregion

		#endregion

		#region 讀取Excel

		#region 參考二

		/// <summary>
		/// 讀取excel文件
		/// </summary>
		/// <param name="filePath">文件路徑</param>
		public void ReadFromExcelFile(string filePath)
		{
			IWorkbook wk = null;
			string extension = Path.GetExtension(filePath);
			try
			{
				FileStream fs = File.OpenRead(filePath);
				if (extension.Equals(".xls"))
				{
					// 2003
					//把xls文件中的數據寫入wk中
					wk = new HSSFWorkbook(fs);
				}
				else
				{
					// 2007
					//把xlsx文件中的數據寫入wk中
					wk = new XSSFWorkbook(fs);
				}

				fs.Close();
				//讀取當前表數據 (取第一個Sheet)
				ISheet sheet = wk.GetSheetAt(0);

				IRow row = sheet.GetRow(0);  //讀取當前行數據
											 //LastRowNum 是當前表的總行數-1（註意）

				string text = string.Empty;
				for (int i = 0; i <= sheet.LastRowNum; i++)
				{
					row = sheet.GetRow(i);  //讀取當前行數據
					if (row != null)
					{
						//LastCellNum 是當前行的總列數
						for (int j = 0; j < row.LastCellNum; j++)
						{
							//讀取該行的第j列數據
							string value = row.GetCell(j).ToString();
							//Console.Write(value.ToString() + " ");
							text = text + value.ToString() + "\r\n";
						}
						//Console.WriteLine("\n");
					}
				}
			}
			catch (Exception e)
			{
				//只在Debug模式下才輸出
				Console.WriteLine(e.Message);
			}
		}

		//獲取cell的數據，並設置為對應的數據類型
		public object GetCellValue(ICell cell)
		{
			object value = null;
			try
			{
				if (cell.CellType != CellType.Blank)
				{
					switch (cell.CellType)
					{
						case CellType.Numeric:
							// Date comes here
							if (DateUtil.IsCellDateFormatted(cell))
								value = cell.DateCellValue;
							else // Numeric type
								value = cell.NumericCellValue;
							break;
						case CellType.Boolean:
							// Boolean type
							value = cell.BooleanCellValue;
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
			}
			return value;
		}

		#endregion

		#endregion

		#region 自動產生圖表

		public void GenaretionChart()
		{
			FileStream RfileStream = new FileStream("D:\\test.xlsx", FileMode.Open, FileAccess.Read);
			//建立讀取資料的FileStream
			XSSFWorkbook wb = new XSSFWorkbook(RfileStream);
			//讀取檔案內的Workbook物件
			ISheet Wsheet = wb.GetSheetAt(1);
			//選擇圖表存放的sheet
			ISheet Rsheet = wb.GetSheetAt(0);
			//選擇資料來源的sheet
			IDrawing drawing = Wsheet.CreateDrawingPatriarch();
			//sheet產生drawing物件
			IClientAnchor clientAnchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, 5, 10);
			//設定圖表位置
			IChart chart = drawing.CreateChart(clientAnchor);
			//產生chart物件
			IChartLegend legend = chart.GetOrCreateLegend();
			//還沒研究出這行在做甚麼
			legend.Position = LegendPosition.TopRight;
			ILineChartData<double, double> data = chart.ChartDataFactory.CreateLineChartData<double, double>();
			//產生存放資料的物件(資料型態為double)
			IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
			//設定X軸
			IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
			//設定Y軸
			bottomAxis.Crosses = AxisCrosses.AutoZero;
			//設定X軸數值開始為0
			leftAxis.Crosses = AxisCrosses.AutoZero;
			//設定Y軸數值開始為0
			IChartDataSource<double> xs = DataSources.FromNumericCellRange(Rsheet, new CellRangeAddress(0, 4, 0, 0));
			//取得要讀取sheet的資料位置(CellRangeAddress(first_row,end_row, first_column, end_column))
			//x軸資料
			IChartDataSource<double> ys1 = DataSources.FromNumericCellRange(Rsheet, new CellRangeAddress(0, 4, 1, 1));
			//第一條y軸資料
			IChartDataSource<double> ys2 = DataSources.FromNumericCellRange(Rsheet, new CellRangeAddress(0, 4, 2, 2));
			//第二條y軸資料
			data.AddSeries(xs, ys1);
			data.AddSeries(xs, ys2);
			//加入到data
			chart.Plot(data, bottomAxis, leftAxis);
			//加入到chart
			FileStream WfileStream = new FileStream("D:\\test.xlsx", FileMode.Create, FileAccess.Write);
			//建立寫入資料的FileStream
			wb.Write(WfileStream);
			//將workbook寫入資料
			RfileStream.Close();
			//關閉FileStream
			WfileStream.Close();
			//關閉FileStream
		}

        #endregion
    }
}
