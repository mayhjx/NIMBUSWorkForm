using NIMBUSWorkForm;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace Tests
{
    public class SampleTest
    {
        [Fact]
        public void Sample_IsX_ShouldBeTrue()
        {
            var number = "";
            var barCode = "";
            var plate = "X1";
            var position = "C1";
            var warnLevel = "X1";
            var warnInfo = "Not used";

            var sample = new Sample(number, barCode,plate,position,warnLevel,warnInfo);

            Assert.True(sample.IsX());
        }

        [Fact]
        public void Sample_NotIsX_ShouldBeFalse()
        {
            var number = "";
            var barCode = "";
            var plate = "X1";
            var position = "D1";
            var warnLevel = "0";
            var warnInfo = "Correct pipetting";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.False(sample.IsX());
        }

        [Fact]
        public void SampleWarnLevelIsZero_IsNormal_ShouldBeTrue()
        {
            var number = "";
            var barCode = "";
            var plate = "X1";
            var position = "D1";
            var warnLevel = "0";
            var warnInfo = "Correct pipetting";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.True(sample.IsNormal());
        }

        [Fact]
        public void SampleIsX_IsNormal_ShouldBeTrue()
        {
            var number = "";
            var barCode = "";
            var plate = "X1";
            var position = "D1";
            var warnLevel = "X1";
            var warnInfo = "Not Used";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.True(sample.IsNormal());
        }

        [Fact]
        public void Sample_WarnLevelIsX_NumberAndBarCode_ShouldBeX()
        {
            var number = "";
            var barCode = "";
            var plate = "X1";
            var position = "D1";
            var warnLevel = "X1";
            var warnInfo = "Not Used";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.True(sample.Number == warnLevel);
            Assert.True(sample.BarCode == warnLevel);
        }

        [Fact]
        public void Sample_PositionIsB1_Order_ShouldBe13()
        {
            var number = "";
            var barCode = "";
            var plate = "X1";
            var position = "B1";
            var warnLevel = "X1";
            var warnInfo = "Not Used";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.True(sample.Order == "13");
        }

        [Fact]
        public void SampleIsSTD_IsClinicSample_ShouldBeFalse()
        {
            var number = "STD1";
            var barCode = "";
            var plate = "X1";
            var position = "B1";
            var warnLevel = "2";
            var warnInfo = "Not Used";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.False(sample.IsClinicSample());
        }

        [Fact]
        public void SampleIsQC_IsClinicSample_ShouldBeFalse()
        {
            var number = "QC1";
            var barCode = "";
            var plate = "X1";
            var position = "C2";
            var warnLevel = "2";
            var warnInfo = "Not Used";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.False(sample.IsClinicSample());
        }

        [Fact]
        public void SampleIsBlank_IsClinicSample_ShouldBeFalse()
        {
            var number = "Blank";
            var barCode = "";
            var plate = "X1";
            var position = "A1";
            var warnLevel = "2";
            var warnInfo = "Not Used";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.False(sample.IsClinicSample());
        }

        [Fact]
        public void Sample_IsClinicSample_ShouldBeTrue()
        {
            var number = "A001";
            var barCode = "12345678";
            var plate = "X1";
            var position = "C2";
            var warnLevel = "0";
            var warnInfo = "Correct pipetting";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.True(sample.IsClinicSample());
        }

        [Fact]
        public void SampleIsNormal_IsWarn_ShouldBeFalse()
        {
            var number = "A001";
            var barCode = "12345678";
            var plate = "X1";
            var position = "C2";
            var warnLevel = "0";
            var warnInfo = "Correct pipetting";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.False(sample.IsWarn());
        }

        [Fact]
        public void SampleNotIsClinicSample_IsWarn_ShouldBeFalse()
        {
            var number = "STD1";
            var barCode = "";
            var plate = "X1";
            var position = "B1";
            var warnLevel = "2";
            var warnInfo = "Not Used";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.False(sample.IsWarn());
        }

        [Fact]
        public void SampleIsEmpty_IsWarn_ShouldBeFalse()
        {
            var number = "";
            var barCode = "";
            var plate = "X1";
            var position = "A1";
            var warnLevel = "2";
            var warnInfo = "Not Used";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.False(sample.IsWarn());
        }

        [Fact]
        public void SampleIsClinicSample_ButUnnormal_IsWarn_ShouldBeTrue()
        {
            var number = "A001";
            var barCode = "12345678";
            var plate = "X1";
            var position = "C1";
            var warnLevel = "4";
            var warnInfo = "Pipetting Warn";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);

            Assert.True(sample.IsWarn());
        }

        [Fact]
        public void Sample_HasMultiNumber_ShouldBeTrue()
        {
            var number = "";
            var barCode = "12345678";
            var plate = "X1";
            var position = "C1";
            var warnLevel = "4";
            var warnInfo = "Pipetting Warn";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);
            
            sample.UpdateNumbers("A001");
            sample.UpdateNumbers("B001");

            Assert.True(sample.HasMultiNumber());
        }

        [Fact]
        public void Sample_SingleNumber_ShouldBeFalse()
        {
            var number = "";
            var barCode = "12345678";
            var plate = "X1";
            var position = "C1";
            var warnLevel = "4";
            var warnInfo = "Pipetting Warn";

            var sample = new Sample(number, barCode, plate, position, warnLevel, warnInfo);
            sample.UpdateNumbers("A001");

            Assert.False(sample.HasMultiNumber());
        }

    }
}
