using ConsoleNPOI.Helper.Model;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.Helper
{
    public class SampleTest
    {
        class Sample
        {
            public string Id { get; set; }   //故意在這裡定義為string 但輸出到Excel中需轉為數字
            public string Name { get; set; }
            public DateTime Birthday { get; set; }
        }

        public SampleTest()
        {
            var data = new List<Sample>
            {
                new Sample { Id = "1", Name = "A", Birthday = DateTime.Parse("2001/1/1 12:00") },
                new Sample { Id = "2", Name = "B", Birthday = DateTime.Parse("2001/1/2 12:00") },
                new Sample { Id = "3", Name = "C", Birthday = DateTime.Parse("2001/1/3 12:00") },
            };

            var param = GetParam(data);
            NpoiHelper.ExportExcel(param);
        }

        private static NpoiParam<Sample> GetParam(List<Sample> data)
        {
            var param = new NpoiParam<Sample>
            {
                Workbook = new XSSFWorkbook(),            //要用一個新的或是用你自己的範本
                Data = data,                              //資料
                FileFullName = @"D:\result.xlsx",         //Excel檔要存在哪
                SheetName = "Data",                       //Sheet要叫什麼名子
                ColumnMapping = new List<ColumnMapping>   //欄位對應 (處理Excel欄名、格式轉換)
                {
                    new ColumnMapping { ModelFieldName = "Id", ExcelColumnName = "流水號", DataType = NpoiDataType.Number, Format = "0.00"},
                    new ColumnMapping { ModelFieldName = "Name", ExcelColumnName = "名子", DataType = NpoiDataType.String},
                    new ColumnMapping { ModelFieldName = "Birthday", ExcelColumnName = "生日", DataType = NpoiDataType.Date, Format="yyyy-MM-dd"},
                },
                FontStyle = new FontStyle                //是否需自定Excel字型大小
                {
                    FontName = "Calibri",
                    FontHeightInPoints = 11,
                },
                ShowHeader = true,                      //是否畫表頭
                IsAutoFit = true,                       //是否啟用自動欄寬
            };

            return param;
        }
    }
}
