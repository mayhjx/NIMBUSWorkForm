using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;

namespace NIMBUSWorkForm
{
    public class SampleTable
    {
        public List<string> PlateNumber { get; set; } // 板号集合
        public List<Sample> Samples { get; set; }

        public List<Sample> GetClinicSamplesByPlateNumber(string plateNumber)
        {
            return Samples.Where(i => i.IsClinicSample() && i.InPlate(plateNumber)).ToList();
        }

        public List<Sample> GetSTDByPlateNumber(string plateNumber)
        {
            return Samples.Where(i => i.IsSTD() && i.InPlate(plateNumber)).ToList();
        }

        public List<Sample> GetGroupOneQCByPlateNumber(string plateNumber)
        {
            return Samples.Where(i => i.IsGroupOneQC() && i.InPlate(plateNumber)).ToList();
        }

        public List<Sample> GetGroupTwoQCByPlateNumber(string plateNumber)
        {
            return Samples.Where(i => i.IsGroupTwoQC() && i.InPlate(plateNumber)).ToList();
        }
    }

    public class Sample
    {
        private readonly List<string> Rows = new List<string> { "A", "B", "C", "D", "E", "F", "G", "H" };

        public Sample(string number, string barCode, string plate, string position,string warnLevel, string warnInfo)
        {
            Number = number;
            BarCode = barCode;
            // 为了在生成上机列表时不管是生成条码还是实验号都显示定位孔
            if (warnLevel.StartsWith('X'))
            {
                Number = warnLevel;
                BarCode = warnLevel;
            }
            Plate = plate;
            Position = position;
            string rowIndex = Position[0].ToString();
            int colIndex = int.Parse(Position.Substring(1));
            Order = $"{Rows.IndexOf(rowIndex) * 12 + colIndex}";
            WarnLevel = warnLevel;
            WarnInfo = warnInfo;
        }

        public bool IsWarn() => !IsNormal() && IsClinicSample();
        public bool IsClinicSample()=> !IsSTD() && !IsQC() && !IsBlank(); // 需包括定位孔
        public bool IsNormal() => WarnLevel == "0" || IsX();

        public bool IsSTD() => Number.StartsWith("STD");
        public bool IsQC() => IsGroupOneQC() || IsGroupTwoQC();
        public bool IsGroupOneQC() => Number.StartsWith("QC") && int.Parse(Order) < 40;
        public bool IsGroupTwoQC() => Number.StartsWith("QC") && int.Parse(Order) > 40;
        public bool IsBlank() => Number.StartsWith("Blank");
        public bool IsX() => WarnLevel.StartsWith("X");

        public bool InPlate(string plate) => Plate == plate;

        public string Number { get; set; } //实验号
        public string BarCode { get; set; } // 条形码
        public string Plate { get; set; } // 板号
        public string Position { get; set; } // 孔位（A1，A2...H12）
        public string Order { get; set; } // 序号（1-96）
        public string WarnLevel { get; set; } // 警告级别 0:正常
        public string WarnInfo { get; set; } //警告信息
    }
}
