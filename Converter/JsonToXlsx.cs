using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using System.Data;
using System.IO;

namespace Converter
{
    public class JsonToXlsx
    {
        private static DataTable _dataTable;

        public JsonToXlsx(string json)
        {
            SetDataTable(json);
        }

        public static implicit operator JsonToXlsx(string json)
            => new JsonToXlsx(json);

        public XLWorkbook Workbook => GetWorkBook();
        public MemoryStream MemoryStream => GetMemoryStream();

        public XLWorkbook GetWorkBook()
        {
            var wb = new XLWorkbook();
            wb.Worksheets.Add(_dataTable);
            return wb;
        }

        public MemoryStream GetMemoryStream()
        {
            var xlsx = GetWorkBook();
            var fs = new MemoryStream();
            xlsx.SaveAs(fs);
            fs.Position = 0;
            return fs;
        }

        private static void SetDataTable(string json)
        {
            _dataTable = new DataTable("Planilha1");
            var jArray = JArray.Parse(json);
            SetColumns(jArray);
            SetRows(jArray);
        }

        private static void SetRows(JArray jArray)
        {
            foreach (var row in jArray)
            {
                var datarow = _dataTable.NewRow();
                foreach (var jToken in row)
                {
                    var jProperty = (JProperty)jToken;
                    datarow[jProperty.Name] = jProperty.Value.ToString();
                }
                _dataTable.Rows.Add(datarow);
            }
        }

        private static void SetColumns(JArray jArray)
        {
            foreach (var row in jArray)
            {
                foreach (var jToken in row)
                {
                    var jproperty = (JProperty)jToken;

                    if (_dataTable.Columns[jproperty.Name] is null)
                    {
                        _dataTable.Columns.Add(jproperty.Name, typeof(string));
                    }
                }
            }
        }
    }
}