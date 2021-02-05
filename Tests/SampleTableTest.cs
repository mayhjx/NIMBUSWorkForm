using System;
using System.Collections.Generic;
using System.Text;
using NIMBUSWorkForm;
using Xunit;

namespace Tests
{
    public class SampleTableTest
    {
        [Fact]
        public void GetClinicSamplesByPlateNumber_ShouldBeReturnThree()
        {
            var plate = "X1";
            var sampleTable = new SampleTable
            {
                PlateNumber = new List<string>
                {
                    plate,
                },
                Samples = new List<Sample>
                {
                    new Sample("AD001","123",plate,"C3","0","Correct pipetting"),
                    new Sample("AD002","456",plate,"D3","0","Correct pipetting"),
                    new Sample("AD003","789",plate,"E3","0","Correct pipetting"),
                }
            };

            var result = sampleTable.GetClinicSamplesByPlateNumber(plate);
            Assert.Equal(3, result.Count);
        }

        [Fact]
        public void GetClinicSamplesByPlateNumber_ShouldBeEmpty()
        {
            var plate = "X1";
            var sampleTable = new SampleTable
            {
                PlateNumber = new List<string>
                {
                    plate,
                },
                Samples = new List<Sample>
                {
                    new Sample("STD0","",plate,"B1","2","Not Used"),
                    new Sample("STD1","",plate,"C1","2","Not Used"),
                    new Sample("STD2","",plate,"D1","2","Not Used"),
                    new Sample("STD3","",plate,"E1","2","Not Used"),
                    new Sample("STD4","",plate,"F1","2","Not Used"),
                    new Sample("STD5","",plate,"G1","2","Not Used"),
                    new Sample("STD6","",plate,"H1","2","Not Used"),
                }
            };

            var result = sampleTable.GetClinicSamplesByPlateNumber(plate);
            Assert.Empty(result);
        }

        [Fact]
        public void GetSTDByPlateNumber_ShouldReturnSeven()
        {
            var plate = "X1";
            var sampleTable = new SampleTable
            {
                PlateNumber = new List<string>
                {
                    plate,
                },
                Samples = new List<Sample>
                {
                    new Sample("STD0","",plate,"B1","2","Not Used"),
                    new Sample("STD1","",plate,"C1","2","Not Used"),
                    new Sample("STD2","",plate,"D1","2","Not Used"),
                    new Sample("STD3","",plate,"E1","2","Not Used"),
                    new Sample("STD4","",plate,"F1","2","Not Used"),
                    new Sample("STD5","",plate,"G1","2","Not Used"),
                    new Sample("STD6","",plate,"H1","2","Not Used"),
                }
            };

            var result = sampleTable.GetSTDByPlateNumber(plate);
            Assert.Equal(7, result.Count);
        }

        [Fact]
        public void GetSTDByPlateNumber_ShouldBeReturnEmpty()
        {
            var plate = "X1";
            var sampleTable = new SampleTable
            {
                PlateNumber = new List<string>
                {
                    plate,
                },
                Samples = new List<Sample>
                {
                    new Sample("AD001","123",plate,"C3","0","Correct pipetting"),
                    new Sample("AD002","456",plate,"D3","0","Correct pipetting"),
                    new Sample("AD003","789",plate,"E3","0","Correct pipetting"),
                }
            };

            var result = sampleTable.GetSTDByPlateNumber(plate);
            Assert.Empty(result);
        }

        [Fact]
        public void GetGroupOneQCByPlateNumber_ShouldBeReturnThree()
        {
            var plate = "X1";
            var sampleTable = new SampleTable
            {
                PlateNumber = new List<string>
                {
                    plate,
                },
                Samples = new List<Sample>
                {
                    new Sample("QC1","",plate,"B2","0","Correct pipetting"),
                    new Sample("QC2","",plate,"C2","0","Correct pipetting"),
                    new Sample("QC3","",plate,"D2","0","Correct pipetting"),
                }
            };

            var result = sampleTable.GetGroupOneQCByPlateNumber(plate);
            Assert.Equal(3, result.Count);
        }

        [Fact]
        public void GetGroupOneQCByPlateNumber_ShouldBeEmpty()
        {
            var plate = "X1";
            var sampleTable = new SampleTable
            {
                PlateNumber = new List<string>
                {
                    plate,
                },
                Samples = new List<Sample>
                {
                    new Sample("QC1","",plate,"E2","0","Correct pipetting"),
                    new Sample("QC2","",plate,"F2","0","Correct pipetting"),
                    new Sample("QC3","",plate,"G2","0","Correct pipetting"),
                }
            };

            var result = sampleTable.GetGroupOneQCByPlateNumber(plate);
            Assert.Empty(result);
        }

        [Fact]
        public void GetGroupTwoQCByPlateNumber_ShouldBeReturnThree()
        {
            var plate = "X1";
            var sampleTable = new SampleTable
            {
                PlateNumber = new List<string>
                {
                    plate,
                },
                Samples = new List<Sample>
                {
                    new Sample("QC1","",plate,"E2","0","Correct pipetting"),
                    new Sample("QC2","",plate,"F2","0","Correct pipetting"),
                    new Sample("QC3","",plate,"G2","0","Correct pipetting"),
                }
            };

            var result = sampleTable.GetGroupTwoQCByPlateNumber(plate);
            Assert.Equal(3, result.Count);
        }

        [Fact]
        public void GetGroupTwoQCByPlateNumber_ShouldBeReturnEmpty()
        {
            var plate = "X1";
            var sampleTable = new SampleTable
            {
                PlateNumber = new List<string>
                {
                    plate,
                },
                Samples = new List<Sample>
                {
                    new Sample("QC1","",plate,"B2","0","Correct pipetting"),
                    new Sample("QC2","",plate,"C2","0","Correct pipetting"),
                    new Sample("QC3","",plate,"D2","0","Correct pipetting"),
                }
            };

            var result = sampleTable.GetGroupTwoQCByPlateNumber(plate);
            Assert.Empty(result);
        }
    }
}
