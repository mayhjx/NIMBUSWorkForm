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
            var clinicSamples = sampleTable.GetClinicSamplesByPlateNumber(plate);
            int rowIndex = 0;
            var row = sh.CreateRow(rowIndex++);

            row.CreateCell(0).SetCellValue("Sample");
            row.CreateCell(1).SetCellValue("Plate");
            row.CreateCell(2).SetCellValue("Vial");

            // 插入两个Blank和三个Test
            row = sh.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue("Blank-1");
            row.CreateCell(1).SetCellValue("1");
            row.CreateCell(2).SetCellValue("20010");

            row = sh.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue("Test-1");
            row.CreateCell(1).SetCellValue("1");
            row.CreateCell(2).SetCellValue("20009");

            row = sh.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue("Test-2");
            row.CreateCell(1).SetCellValue("1");
            row.CreateCell(2).SetCellValue("20009");

            row = sh.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue("Test-3");
            row.CreateCell(1).SetCellValue("1");
            row.CreateCell(2).SetCellValue("20009");

            row = sh.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue("Blank-2");
            row.CreateCell(1).SetCellValue("1");
            row.CreateCell(2).SetCellValue("20010");

            // 插入曲线
            var std = sampleTable.GetSTDByPlateNumber(plate);
            foreach (var sample in std)
            {
                row = sh.CreateRow(rowIndex++);
                row.CreateCell(0).SetCellValue(sample.Number);
                row.CreateCell(1).SetCellValue("1");
                row.CreateCell(2).SetCellValue(sample.Order);
            }

            var groupOneQCs = sampleTable.GetGroupOneQCByPlateNumber(plate);
            var groupTwoQCs = sampleTable.GetGroupTwoQCByPlateNumber(plate);

            // 前面插入第一组质控
            foreach (var qc in groupOneQCs)
            {
                row = sh.CreateRow(rowIndex++);
                row.CreateCell(0).SetCellValue(qc.Number);
                row.CreateCell(1).SetCellValue("1");
                row.CreateCell(2).SetCellValue(qc.Order);
            }

            bool hasGroupTowQC = false;
            int clinicSampleCount = 0;
            foreach (var sample in clinicSamples)
            {
                if (clinicSampleCount == 40 && !hasGroupTowQC)
                {
                    // 第40个样后面插入第一组质控
                    foreach (var qc in groupOneQCs)
                    {
                        row = sh.CreateRow(rowIndex++);
                        row.CreateCell(0).SetCellValue(qc.Number);
                        row.CreateCell(1).SetCellValue("1");
                        row.CreateCell(2).SetCellValue(qc.Order);
                    }
                    hasGroupTowQC = true;
                }
                else
                {
                    // 插入临床样品
                    row = sh.CreateRow(rowIndex++);
                    row.CreateCell(0).SetCellValue(sample.BarCode);
                    row.CreateCell(1).SetCellValue("1");
                    row.CreateCell(2).SetCellValue(sample.Order);
                    clinicSampleCount++;
                }
            }

            // 最后插入第二组质控
            foreach (var qc in groupTwoQCs)
            {
                row = sh.CreateRow(rowIndex++);
                row.CreateCell(0).SetCellValue(qc.Number);
                row.CreateCell(1).SetCellValue("1");
                row.CreateCell(2).SetCellValue(qc.Order);
                clinicSampleCount++;
            }
        }

        private static void AddWarnInfoData(SampleTable sampleTable, ISheet sh)
        {
            int rowIndex = 0;
            var row = sh.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue("实验号");
            row.CreateCell(1).SetCellValue("板号");
            row.CreateCell(2).SetCellValue("孔位");
            row.CreateCell(3).SetCellValue("条码");
            row.CreateCell(4).SetCellValue("警告级别");
            row.CreateCell(5).SetCellValue("警告信息");

            foreach (var sample in sampleTable.Samples)
            {
                // 0：打液正常，X?：定位孔
                if (sample.IsWarn())
                {
                    row = sh.CreateRow(rowIndex++);
                    row.CreateCell(0).SetCellValue(sample.Number);
                    row.CreateCell(1).SetCellValue(sample.Plate);
                    row.CreateCell(2).SetCellValue(sample.Position);
                    row.CreateCell(3).SetCellValue(sample.BarCode);
                    row.CreateCell(4).SetCellValue(sample.WarnLevel);
                    row.CreateCell(5).SetCellValue(sample.WarnInfo);
                }
            }
        }
    }
}
