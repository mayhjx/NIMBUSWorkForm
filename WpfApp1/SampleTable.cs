using System;
using System.Collections.Generic;
using System.Text;

namespace NIMBUSWorkForm
{
    public class SampleTable
    {
        public List<string> PlateNumber { get; set; } // 板号集合
        public List<Sample> Samples { get; set; }
    }

    public class Sample
    {
        private List<string> Rows = new List<string> { "A", "B", "C", "D", "E", "F", "G", "H" };

        public Sample(string number, string barCode, string plate, string position, string warmInfo)
        {
            Number = number;
            if (barCode == "----------")
            {
                barCode = "";
            }
            BarCode = barCode;
            Plate = plate;
            Position = position;
            string rowIndex = Position[0].ToString();
            int colIndex = int.Parse(Position.Substring(1));
            Order = $"{Rows.IndexOf(rowIndex) * 12 + colIndex}";
            WarmInfo = warmInfo;
        }
        public string Number { get; set; } //实验号
        public string BarCode { get; set; } // 条形码
        public string Plate { get; set; } // 板号
        public string Position { get; set; } // 孔位（A1，A2...H12）
        public string Order { get; set; } // 序号（1-96）
        public string WarmInfo { get; set; } //警告信息
    }
}
