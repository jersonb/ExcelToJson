using ClosedXML.Excel;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Converter
{
    public class XlsxToJson
    {
        private IEnumerable<string> _headers;
        private readonly IXLWorksheet _worksheet;
        private readonly DataTable _dataTable;

        public XlsxToJson(Stream stream)
        {
            var wb = new XLWorkbook(stream);
            _worksheet = wb.Worksheets.Worksheet(1);
            _dataTable = new DataTable();
            SetHeader();
            SetBody();
        }

        public static implicit operator XlsxToJson(Stream stream)
            => new XlsxToJson(stream);

        public object Json
            => JsonConvert.SerializeObject(new { table = _dataTable }, Formatting.Indented);

        private void SetBody()
        {
            var body = _worksheet.Rows(2, _worksheet.Rows().Count());
            foreach (var row in body)
            {
                int i = 0;
                var dataRow = _dataTable.NewRow();
                foreach (var cell in row.Cells())
                {
                    dataRow[_headers.ElementAt(i++)] = cell.Value.ToString();
                }
                _dataTable.Rows.Add(dataRow);
            }
        }

        private void SetHeader()
        {
            _headers = _worksheet.FirstRowUsed().Cells().Select(c => c.Value.ToString());
            foreach (var cell in _headers)
            {
                _dataTable.Columns.Add(cell);
            }
        }
    }
}