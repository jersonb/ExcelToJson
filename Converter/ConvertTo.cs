using ClosedXML.Excel;
using System.IO;

namespace Converter
{
    public static class ConvertTo
    {
        public static object Json(Stream stream)
        {
            XlsxToJson xlsxToJson = stream;
            return xlsxToJson.Json;
        }

        public static (MemoryStream MemoryStream, XLWorkbook Workbook) Xlsx(string json)
        {
            JsonToXlsx jsonToXlsx = json;
            return (jsonToXlsx.MemoryStream, jsonToXlsx.Workbook);
        }
    }
}