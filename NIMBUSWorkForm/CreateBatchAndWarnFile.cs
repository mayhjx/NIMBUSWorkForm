using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NIMBUSWorkForm
{
    public static class BatchAndWarnFile
    {
        public static void Create(SampleTable sampleTable, string filePath)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet;

            // 生成上机列表
            foreach (var plate in sampleTable.PlateNumber)
            {
                sheet = workbook.CreateSheet(plate);
                AddBatch(sampleTable, plate, sheet);
            }

            // 生成警告信息
            sheet = workbook.CreateSheet("警告信息");
            AddWarnInfoData(sampleTable, sheet);

            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }
        }

        private static void AddBatch(SampleTable sampleTable, string plate, ISheet sh)
        {
            int rowIndex = 0;

            // 标题
            SetBatchCell(sh.CreateRow(rowIndex++), "Sample", "Vial", "Plate");
            // 插入两个Blank和三个Test
            SetBatchCell(sh.CreateRow(rowIndex++), "Blank-1", "20010");
            SetBatchCell(sh.CreateRow(rowIndex++), "Test-1", "20009");
            SetBatchCell(sh.CreateRow(rowIndex++), "Test-2", "20009");
            SetBatchCell(sh.CreateRow(rowIndex++), "Test-3", "20009");
            SetBatchCell(sh.CreateRow(rowIndex++), "Blank-2", "20010");
            SetBatchCell(sh.CreateRow(rowIndex++), "KB" + plate[1..], "1");

            // 插入曲线
            var std = sampleTable.GetSTDByPlateNumber(plate);
            foreach (var sample in std)
            {
                SetBatchCell(sh.CreateRow(rowIndex++), sample.Number, sample.Order);
            }

            SetBatchCell(sh.CreateRow(rowIndex++), "Blank-3", "20010");

            // 插入第一组质控
            var groupOneQCs = sampleTable.GetGroupOneQCByPlateNumber(plate);
            foreach (var qc in groupOneQCs)
            {
                SetBatchCell(sh.CreateRow(rowIndex++), GetQCStringFromSettings(qc.Number), qc.Order);
            }

            SetBatchCell(sh.CreateRow(rowIndex++), "Blank-4", "20010");

            // 插入临床样品
            var clinicSamples = sampleTable.GetClinicSamplesByPlateNumber(plate);
            foreach (var sample in clinicSamples)
            {
                if (Settings.Default.ShowBarCode)
                {
                    SetBatchCell(sh.CreateRow(rowIndex++), sample.BarCode, sample.Order);
                }
                else
                {
                    SetBatchCell(sh.CreateRow(rowIndex++), sample.Number, sample.Order);
                }

                // 第90孔后面插入一个blank
                if (sample.Order == "90")
                {
                    SetBatchCell(sh.CreateRow(rowIndex++), "Blank-5", "20010");
                }
            }

            SetBatchCell(sh.CreateRow(rowIndex++), "Blank-6", "20010");

            // 插入第二组质控
            var groupTwoQCs = sampleTable.GetGroupTwoQCByPlateNumber(plate);
            foreach (var qc in groupTwoQCs)
            {
                SetBatchCell(sh.CreateRow(rowIndex++), GetQCStringFromSettings(qc.Number), qc.Order);
            }

            SetBatchCell(sh.CreateRow(rowIndex++), "Blank-7", "20010");
        }

        private static void SetBatchCell(IRow row, string name, string order, string plate = "1")
        {
            row.CreateCell(0).SetCellValue(name);
            row.CreateCell(1).SetCellValue(plate);
            row.CreateCell(2).SetCellValue(order);
        }

        private static void AddWarnInfoData(SampleTable sampleTable, ISheet sh)
        {
            int rowIndex = 0;
            var row = sh.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue("实验号");
            row.CreateCell(1).SetCellValue("条码");
            row.CreateCell(2).SetCellValue("板号");
            row.CreateCell(3).SetCellValue("孔位");
            row.CreateCell(4).SetCellValue("警告级别");
            row.CreateCell(5).SetCellValue("警告信息");

            foreach (var sample in sampleTable.Samples)
            {
                if (sample.IsWarn())
                {
                    row = sh.CreateRow(rowIndex++);
                    row.CreateCell(0).SetCellValue(sample.Number);
                    row.CreateCell(1).SetCellValue(sample.BarCode);
                    row.CreateCell(2).SetCellValue(sample.Plate);
                    row.CreateCell(3).SetCellValue(sample.Position);
                    row.CreateCell(4).SetCellValue(sample.WarnLevel);
                    row.CreateCell(5).SetCellValue(sample.WarnInfo);
                }
            }
        }

        private static string GetQCStringFromSettings(string qc)
        {
            switch (qc)
            {
                case "QC1":
                    return Settings.Default.QC1;
                case "QC2":
                    return Settings.Default.QC2;
                case "QC3":
                    return Settings.Default.QC3;
                default:
                    return qc;
            }
        }
    }
}
