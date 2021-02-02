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

            // 曲线质控位置
            var STDQC = new Dictionary<string, string>
            {
                { "A1", "Blank" },
                { "B1", "STD0"},
                { "C1", "STD1"},
                { "D1", "STD2"},
                { "E1", "STD3"},
                { "F1", "STD4"},
                { "G1", "STD5"},
                { "H1", "STD6"},
                { "B2", "QC1"},
                { "C2", "QC2"},
                { "D2", "QC3"},
                { "E2", "QC1"},
                { "F2", "QC2"},
                { "G2", "QC3"}
            };

            try
            {
                IWorkbook workbook = ReadExecl(filePath);
                if (workbook == null)
                {
                    throw new IOException($"无法打开{filePath}");
                }

                ISheet worksheet = workbook.GetSheetAt(0);

                int posCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "TPositionId").ColumnIndex;
                int bcCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "SPositionBC").ColumnIndex;
                int warnInfoCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "TSumStateDescription").ColumnIndex;
                int warnLevelCol = worksheet.GetRow(0).Cells.FirstOrDefault(i => i.ToString() == "Warm").ColumnIndex;
                // TODO can not find the column

                // 获取板号
                for (int i = 1; i <= worksheet.LastRowNum; i++)
                {
                    var row = worksheet.GetRow(i);
                    var warnLevel = row.GetCell(warnLevelCol).ToString();
                    if (warnLevel.StartsWith('X'))
                    {
                        plates.Add(warnLevel);
                    }
                }

                if (plates.Count == 0)
                {
                    throw new ArgumentException("未识别到板号");
                }

                int p = -1; // 板号下标
                for (int i = 1; i <= worksheet.LastRowNum; i++)
                {
                    var row = worksheet.GetRow(i);
                    if (row.Cells.Count == 0) { continue; }

                    string num = "";
                    string bc = row.GetCell(bcCol).ToString();
                    string pos = row.GetCell(posCol).ToString();
                    string warnLevel = row.GetCell(warnLevelCol).ToString();
                    string warnInfo = row.GetCell(warnInfoCol).ToString();

                    if (bc == "----------")
                    {
                        bc = "";
                    }
                    if (pos == "A1")
                    {
                        p++;
                    }
                    if (STDQC.ContainsKey(pos))
                    {
                        num = STDQC.GetValueOrDefault(pos);
                    }

                    samples.Add(new Sample(num, bc, plates[p], pos, warnLevel, warnInfo));
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
