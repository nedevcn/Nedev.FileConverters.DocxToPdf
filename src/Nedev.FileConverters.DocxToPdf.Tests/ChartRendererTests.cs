using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.Rendering;
using System.IO;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace Nedev.FileConverters.DocxToPdf.Tests
{
    public class ChartRendererTests
    {
        [Fact]
        public void ChartData_Properties_WorkCorrectly()
        {
            var data = new ChartData
            {
                Title = "Test Chart",
                ChartType = ChartType.Bar,
                CategoryAxisTitle = "Categories",
                ValueAxisTitle = "Values"
            };

            var series = new ChartSeries
            {
                Name = "Series 1"
            };
            series.Categories.AddRange(new[] { "A", "B", "C" });
            series.Values.AddRange(new[] { 10.0, 20.0, 30.0 });
            data.Series.Add(series);

            Assert.Equal("Test Chart", data.Title);
            Assert.Equal(ChartType.Bar, data.ChartType);
            Assert.Equal("Categories", data.CategoryAxisTitle);
            Assert.Equal("Values", data.ValueAxisTitle);
            Assert.Single(data.Series);
            Assert.Equal("Series 1", data.Series[0].Name);
            Assert.Equal(3, data.Series[0].Categories.Count);
            Assert.Equal(3, data.Series[0].Values.Count);
        }

        [Fact]
        public void ChartType_Enum_HasExpectedValues()
        {
            Assert.Equal(0, (int)ChartType.Bar);
            Assert.Equal(1, (int)ChartType.Column);
            Assert.Equal(2, (int)ChartType.Line);
            Assert.Equal(3, (int)ChartType.Pie);
            Assert.Equal(4, (int)ChartType.Area);
            Assert.Equal(5, (int)ChartType.Scatter);
            Assert.Equal(6, (int)ChartType.Radar);
        }

        [Fact]
        public void ChartSeries_DefaultValues_AreEmpty()
        {
            var series = new ChartSeries();
            
            Assert.Equal(string.Empty, series.Name);
            Assert.Empty(series.Categories);
            Assert.Empty(series.Values);
        }

        [Fact]
        public void ChartData_DefaultValues_AreCorrect()
        {
            var data = new ChartData();
            
            Assert.Equal(string.Empty, data.Title);
            Assert.Equal(ChartType.Bar, data.ChartType);
            Assert.Equal(string.Empty, data.CategoryAxisTitle);
            Assert.Equal(string.Empty, data.ValueAxisTitle);
            Assert.Empty(data.Series);
        }

        [Fact]
        public void ChartData_MultipleSeries_WorkCorrectly()
        {
            var data = new ChartData
            {
                Title = "Multi-Series Chart",
                ChartType = ChartType.Line
            };

            // 添加第一个系列
            var series1 = new ChartSeries
            {
                Name = "Q1"
            };
            series1.Categories.AddRange(new[] { "Jan", "Feb", "Mar" });
            series1.Values.AddRange(new[] { 100.0, 150.0, 200.0 });
            data.Series.Add(series1);

            // 添加第二个系列
            var series2 = new ChartSeries
            {
                Name = "Q2"
            };
            series2.Categories.AddRange(new[] { "Jan", "Feb", "Mar" });
            series2.Values.AddRange(new[] { 120.0, 180.0, 220.0 });
            data.Series.Add(series2);

            Assert.Equal(2, data.Series.Count);
            Assert.Equal("Q1", data.Series[0].Name);
            Assert.Equal("Q2", data.Series[1].Name);
            Assert.Equal(3, data.Series[0].Values.Count);
            Assert.Equal(3, data.Series[1].Values.Count);
        }

        [Fact]
        public void ChartRenderer_Constructor_InitializesCorrectly()
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();

                var options = new ConvertOptions();
                var renderer = new ChartRenderer(doc, options);

                Assert.NotNull(renderer);
            }
        }
    }
}
