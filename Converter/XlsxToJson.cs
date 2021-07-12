using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
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
            => new(stream);

        public object Json
            => JsonConvert.SerializeObject(new { table = _dataTable }, Formatting.Indented);

        private void SetBody()
        {
            var body = _worksheet.Rows(2, _worksheet.Rows().Count()-1);
            foreach (var row in body)
            {
                int i = 0;
                var dataRow = _dataTable.NewRow();

                foreach (var cell in row.Cells(false).Take(_headers.Count()))
                {
                    dataRow[_headers.ElementAt(i++)] = cell.Value;
                }

                _dataTable.Rows.Add(dataRow);
            }
        }

        private void SetHeader()
        {
            _headers = _worksheet.FirstRowUsed().Cells().Select(c => c.Value.ToString());

            var types = _worksheet.Rows()
                                  .Skip(1)
                                  .FirstOrDefault(x => x.CellsUsed().Count() == _headers.Count())
                                  .Cells()
                                  .Select(x => x.DataType);

            var i = 0;
            foreach (var name in _headers)
            {
                var type = types.ElementAt(i++) switch
                {
                    XLDataType.Text => typeof(string),
                    XLDataType.Number => typeof(decimal),
                    XLDataType.Boolean => typeof(bool),
                    XLDataType.DateTime => typeof(DateTime),
                    XLDataType.TimeSpan => typeof(DateTime),
                    _ => typeof(string)
                };

                _dataTable.Columns.Add(name, type);
            }
        }
    }
}