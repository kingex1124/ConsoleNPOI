using ConsoleNPOI.MyHelper.Model;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.MyHelper
{
    public class SampleTest2
    {
        public class Sample
        {
            public double Id { get; set; }   //故意在這裡定義為string 但輸出到Excel中需轉為數字
            public string Name { get; set; }
            public DateTime Birthday { get; set; }
        }

        public SampleTest2()
        {
            var data = new List<Sample>
            {
                new Sample { Id = 1, Name = "A", Birthday = DateTime.Parse("2001/1/1 12:00") },
                new Sample { Id = 2, Name = "B", Birthday = DateTime.Parse("2001/1/2 12:00") },
                new Sample { Id = 3, Name = "C", Birthday = DateTime.Parse("2001/1/3 12:00") },
            };

            var param = GetParam(data);
            NpoiHelper.ExportExcel(param);
        }

        public static NpoiParam<Sample> GetParam(List<Sample> data)
        {
            var param = new NpoiParam<Sample>
            {
                Workbook = new XSSFWorkbook(),
                //Workbook = new HSSFWorkbook(),            //要用一個新的或是用你自己的範本
                Data = new List<Sample>[] { data ,data} ,                              //資料
                FileFullName = @"C:\Users\011714\Desktop\result.xlsx",         //Excel檔要存在哪
                //FileFullName = @"C:\Users\011714\Desktop\result.xls",
                SheetName = new string[] { "data","data2" },                       //Sheet要叫什麼名子
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
    }
}
