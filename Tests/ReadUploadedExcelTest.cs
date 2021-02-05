using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Text;
using NIMBUSWorkForm;
using Xunit;

namespace Tests
{
    public class ReadUploadedExcelTest
    {
        [Fact]
        public void IsTargetSampleName_ShouldBeTrue()
        {
            var targetPreFix = new StringCollection { "AD" };
            var sampleNames = "AD001";

            var result = ReadUploadedExcel.IsTargetSampleName(targetPreFix, sampleNames);
            Assert.True(result);
        }

        [Fact]
        public void IsTargetSampleName_ShouldBeFalse()
        {
            var targetPreFix = new StringCollection { "T" };
            var sampleNames = "TP001";

            var result = ReadUploadedExcel.IsTargetSampleName(targetPreFix, sampleNames);
            Assert.False(result);
        }

        [Fact]
        public void GetStringPart_ShouldBeEqual()
        {
            var input = "TP001";
            var result = ReadUploadedExcel.GetStringPart(input);
            Assert.Equal("TP", result);
        }

        [Fact]
        public void GetStringPart_ShouldBeNotEqual()
        {
            var input = "TP001";
            var result = ReadUploadedExcel.GetStringPart(input);
            Assert.NotEqual("T", result);
        }
    }
}
