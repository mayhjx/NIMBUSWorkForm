using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace NIMBUSWorkForm
{
    public static class ReadUploadedExcel
    {
        public static SampleTable ReadNIMBUSWorkBook(string filePath)
        {
            // 读取NIMBUS工作清单中的板号，孔号，条码和警告信息
            var sampleTable = new SampleTable();
            var samples = new List<Sample>();
            var plates = new List<string>();
            sampleTable.PlateNumber = plates;
            sampleTable.Samples = samples;

            try
            {
                IWorkbook workbook = ReadExecl(filePath);
                if (workbook == null)
                {
                    throw new IOException($"无法打开{filePath}");
                }

                ISheet worksheet = workbook.GetSheetAt(0);

                int idCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "TPositionId").ColumnIndex;
                int bcCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "SPositionBC").ColumnIndex;
                int warmCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "Warm").ColumnIndex;
                // TODO can not find the column

                // 获取板号，最多四块板
                for (int i = 1; i <= worksheet.LastRowNum; i++)
                {
                    var row = worksheet.GetRow(i);
                    var warm = row.GetCell(warmCol).ToString();
                    if (warm.StartsWith('X'))
                    {
                        plates.Add(warm);
                    }
                }

                if (plates.Count == 0)
                {
                    throw new ArgumentException("未识别到板号");
                }

                int p = -1;
                for (int i = 1; i <= worksheet.LastRowNum; i++)
                {
                    var row = worksheet.GetRow(i);
                    if (row.Cells.Count == 0) { continue; }

                    string bc = row.GetCell(bcCol).ToString();
                    string pos = row.GetCell(idCol).ToString();
                    string warm = row.GetCell(warmCol).ToString();

                    if (pos == "A1")
                    {
                        p++;
                    }
                    string plate = plates[p];

                    var sample = new Sample(null, bc, plate, pos, warm);
                    samples.Add(sample);
                }
            }
            catch (Exception)
            {
                throw;
            }

            return sampleTable;
        }

        public static SampleTable ReadOperateWorkBook(SampleTable sampleTable, string filePath)
        {
            // 读取每日操作清单，补充实验号到samples中
            try
            {
                IWorkbook workbook = ReadExecl(filePath);
                if (workbook == null)
                {
                    throw new IOException($"无法打开{filePath}");
                }

                ISheet worksheet = workbook.GetSheetAt(0);

                int bcCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "主条码").ColumnIndex;
                int numCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "实验号").ColumnIndex;
                // TODO can not find the column

                for (int i = 1; i <= worksheet.LastRowNum; i++)
                {
                    var row = worksheet.GetRow(i);
                    if (row.Cells.Count == 0) { continue; }

                    string barCode = row.GetCell(bcCol).ToString();
                    string number = row.GetCell(numCol).ToString();

                    // 每日操作清单中的条码不包含“-”
                    var sample = sampleTable.Samples.Find(i => i.BarCode.Split('-')[0] == barCode);
                    if (sample != null)
                    {
                        sample.Number = number;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return sampleTable;
        }


        private static IWorkbook ReadExecl(string filePath)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    try
                    {
                        workbook = new HSSFWorkbook(fileStream);
                    }
                    catch
                    {
                        workbook = null;
                    }
                }
                if (workbook == null)
                {
                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        workbook = new XSSFWorkbook(fileStream);
                    }
                }
            }
            catch
            {
                throw;
            }
            return workbook;
        }
    }
}
