using ConsoleNPOI.MyHelper.Model;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.MyHelper
{
    public class SampleTest2
    {
        public class Sample
        {
            public string Id { get; set; }   //故意在這裡定義為string 但輸出到Excel中需轉為數字
            public string Name { get; set; }
            public DateTime Birthday { get; set; }
        }

        public SampleTest2()
        {
            #region ListExportExcel使用範例

            var data = new List<Sample>
            {
                new Sample { Id = "1", Name = "A", Birthday = DateTime.Parse("2001/1/1 12:00") },
                new Sample { Id = "2", Name = "B", Birthday = DateTime.Parse("2001/1/2 12:00") },
                new Sample { Id = "3", Name = "C", Birthday = DateTime.Parse("2001/1/3 12:00") },
            };

            var param = GetParam(data);
            NpoiHelper.ListExportExcel(param);

            #endregion

            #region DataSetExportExcel 使用範例

            var dataSet = new DataSet();

            //table
            var dataTable = new DataTable();

            dataTable.Columns.Add("Col1");
            dataTable.Columns.Add("Col2");
            dataTable.Columns.Add("Col3");

            var dataRow = dataTable.NewRow();
            dataRow[0] = "1";
            dataRow[1] = "A";
            dataRow[2] = DateTime.Parse("2001/1/1 12:00");

            dataTable.Rows.Add(dataRow);

            var dataRow1 = dataTable.NewRow();
            dataRow1[0] = "2";
            dataRow1[1] = "B";
            dataRow1[2] = DateTime.Parse("2001/1/2 12:00");

            dataTable.Rows.Add(dataRow1);

            dataSet.Tables.Add(dataTable);

            //table1
            var dataTable1 = new DataTable();

            dataTable1.Columns.Add("Col1");
            dataTable1.Columns.Add("Col2");
            dataTable1.Columns.Add("Col3");

            var dataRow2 = dataTable1.NewRow();
            dataRow2[0] = "3";
            dataRow2[1] = "C";
            dataRow2[2] = DateTime.Parse("2001/1/3 12:00");

            dataTable1.Rows.Add(dataRow2);

            var dataRow3 = dataTable1.NewRow();
            dataRow3[0] = "4";
            dataRow3[1] = "D";
            dataRow3[2] = DateTime.Parse("2001/1/4 12:00");

            dataTable1.Rows.Add(dataRow3);

            dataSet.Tables.Add(dataTable1);

            var dataSetParam = GetDataSetParam(dataSet);
            NpoiHelper.DataSetExportExcel(dataSetParam);

            #endregion

            #region DataTableExportExcel 使用範例

            //table2
            var dataTable2 = new DataTable();

            dataTable2.Columns.Add("Col1");
            dataTable2.Columns.Add("Col2");
            dataTable2.Columns.Add("Col3");

            var dataRow4 = dataTable2.NewRow();
            dataRow4[0] = "5";
            dataRow4[1] = "E";
            dataRow4[2] = DateTime.Parse("2001/1/5 12:00");

            dataTable2.Rows.Add(dataRow4);

            var dataRow5 = dataTable2.NewRow();
            dataRow5[0] = "6";
            dataRow5[1] = "F";
            dataRow5[2] = DateTime.Parse("2001/1/6 12:00");

            dataTable2.Rows.Add(dataRow5);

            var dataTableParam = GetDataTableParam(dataTable2);
            NpoiHelper.DataTableExportExcel(dataTableParam);

            #endregion
        }

        /// <summary>
        /// ListExportExcel 使用參數設定
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static NpoiParam<Sample> GetParam(List<Sample> data)
        {
            var param = new NpoiParam<Sample>
            {
                Workbook = new XSSFWorkbook(),
                //Workbook = new HSSFWorkbook(),            //要用一個新的或是用你自己的範本
                DataList = new List<Sample>[] { data, data },                              //資料
                FileFullName = @"C:\Users\011714\Desktop\result.xlsx",         //Excel檔要存在哪
                //FileFullName = @"C:\Users\011714\Desktop\result.xls",
                SheetName = new string[] { "data", "data2" },                       //Sheet要叫什麼名子
                ColumnMapping = new List<ColumnMapping>[]   //欄位對應 (處理Excel欄名、格式轉換)
                {
                    new List<ColumnMapping>()
                    {
                    new ColumnMapping { ModelFieldName = "Id", ExcelColumnName = "流水號", DataType = NpoiDataType.Int, Format = "0.00"},
                    new ColumnMapping { ModelFieldName = "Name", ExcelColumnName = "名子", DataType = NpoiDataType.String},
                    new ColumnMapping { ModelFieldName = "Birthday", ExcelColumnName = "生日", DataType = NpoiDataType.DateTime, Format="yyyy-MM-dd"}
                    },
                    new List<ColumnMapping>()
                    {
                    new ColumnMapping { ModelFieldName = "Id", ExcelColumnName = "流水號", DataType = NpoiDataType.Int, Format = "0.00"},
                    new ColumnMapping { ModelFieldName = "Name", ExcelColumnName = "名子", DataType = NpoiDataType.String},
                    new ColumnMapping { ModelFieldName = "Birthday", ExcelColumnName = "生日", DataType = NpoiDataType.DateTime, Format="yyyy-MM-dd"}
                    }
                },
                HeaderFontStyle = new FontStyle                //是否需自定Excel字型大小
                {
                    FontName = "Calibri",
                    FontHeightInPoints = 11,
                },
                DataFontStyle = new FontStyle
                {
                    FontName = "新細明體",
                    FontHeightInPoints = 10,
                },
                ShowHeader = true,                      //是否畫表頭
                IsAutoFit = true,                       //是否啟用自動欄寬
            };

            return param;
        }

        /// <summary>
        /// DataSetExportExcel 使用參數設定
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static NpoiParam<Sample> GetDataSetParam(DataSet data)
        {
            var param = new NpoiParam<Sample>
            {
                Workbook = new XSSFWorkbook(),
                //Workbook = new HSSFWorkbook(),            //要用一個新的或是用你自己的範本
                 DataSet = data,                              //資料
                FileFullName = @"C:\Users\011714\Desktop\result1.xlsx",         //Excel檔要存在哪
                //FileFullName = @"C:\Users\011714\Desktop\result.xls",
                SheetName = new string[] { "data", "data1" },                       //Sheet要叫什麼名子
                ColumnMapping = new List<ColumnMapping>[]   //欄位對應 (處理Excel欄名、格式轉換)
                {
                    new List<ColumnMapping>()
                    {
                    new ColumnMapping { ModelFieldName = "Id", ExcelColumnName = "流水號", DataType = NpoiDataType.String},
                    new ColumnMapping { ModelFieldName = "Name", ExcelColumnName = "名子", DataType = NpoiDataType.String},
                    new ColumnMapping { ModelFieldName = "Birthday", ExcelColumnName = "生日", DataType = NpoiDataType.DateTime, Format="yyyy-MM-dd"}
                    },
                    new List<ColumnMapping>()
                    {
                    new ColumnMapping { ModelFieldName = "Id", ExcelColumnName = "流水號", DataType = NpoiDataType.String},
                    new ColumnMapping { ModelFieldName = "Name", ExcelColumnName = "名子", DataType = NpoiDataType.String},
                    new ColumnMapping { ModelFieldName = "Birthday", ExcelColumnName = "生日", DataType = NpoiDataType.DateTime, Format="yyyy-MM-dd"}
                    }
                },
                HeaderFontStyle = new FontStyle                //是否需自定Excel字型大小
                {
                    FontName = "Calibri",
                    FontHeightInPoints = 11,
                },
                DataFontStyle = new FontStyle
                {
                    FontName = "新細明體",
                    FontHeightInPoints = 10,
                },
                ShowHeader = true,                      //是否畫表頭
                IsAutoFit = true,                       //是否啟用自動欄寬
            };

            return param;
        }

        /// <summary>
        /// DataTableExportExcel 使用參數設定
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static NpoiParam<Sample> GetDataTableParam(DataTable data)
        {
            var param = new NpoiParam<Sample>
            {
                Workbook = new XSSFWorkbook(),
                //Workbook = new HSSFWorkbook(),            //要用一個新的或是用你自己的範本
                DataTable = data,                              //資料
                FileFullName = @"C:\Users\011714\Desktop\result2.xlsx",         //Excel檔要存在哪
                //FileFullName = @"C:\Users\011714\Desktop\result.xls",
                SheetName = new string[] { "data" },                       //Sheet要叫什麼名子
                ColumnMapping = new List<ColumnMapping>[]   //欄位對應 (處理Excel欄名、格式轉換)
                {
                    new List<ColumnMapping>()
                    {
                    new ColumnMapping { ModelFieldName = "Id", ExcelColumnName = "流水號", DataType = NpoiDataType.String},
                    new ColumnMapping { ModelFieldName = "Name", ExcelColumnName = "名子", DataType = NpoiDataType.String},
                    new ColumnMapping { ModelFieldName = "Birthday", ExcelColumnName = "生日", DataType = NpoiDataType.DateTime, Format="yyyy-MM-dd"}
                    }
                },
                HeaderFontStyle = new FontStyle                //是否需自定Excel字型大小
                {
                    FontName = "Calibri",
                    FontHeightInPoints = 11,
                },
                DataFontStyle = new FontStyle
                {
                    FontName = "新細明體",
                    FontHeightInPoints = 10,
                },
                ShowHeader = true,                      //是否畫表頭
                IsAutoFit = true,                       //是否啟用自動欄寬
            };

            return param;
        }
    }
}
