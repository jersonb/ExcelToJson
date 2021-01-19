using System.IO;
using System.Text;
using Xunit;

namespace Converter.Test
{
    public class ConvertTest
    {
        [Fact]
        public void GenerateJsonTest()
        {
            var pathXlsx = "./Datas/people.xlsx";
            var pathJson = "./Datas/people-teste.json";

            if (File.Exists(pathJson))
            {
                File.Delete(pathJson);
            }

            var stream = File.OpenRead(pathXlsx);
            XlsxToJson xlsxToJson = stream;
            var json = xlsxToJson.Json;

            File.WriteAllText(pathJson, json.ToString());

            Assert.True(File.Exists(pathJson));
        }

        [Fact]
        public void GenerateFileTest()
        {
            var path = "./Datas/people.xlsx";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            var json = File.ReadAllText("./Datas/people.json", Encoding.UTF7);

            JsonToXlsx jsonToXlsx = json;

            var wb = jsonToXlsx.Workbook;

            wb.SaveAs(path);

            Assert.True(File.Exists(path));
        }

        [Fact]
        public void GenerateFileFromMemoryStreamTest()
        {
            var path = "./Datas/people.xlsx";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            var json = File.ReadAllText("./Datas/people.json", Encoding.UTF7);

            JsonToXlsx jsonToXlsx = json;

            using var ms = jsonToXlsx.MemoryStream;

            File.WriteAllBytes(path, ms.ToArray());

            Assert.True(File.Exists(path));
        }
    }
}