using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAdapter
{
    public class SimpleExcelAdapter
    {
        private readonly string path;
        private _Application excel = new _Excel.Application();
        private Workbook wb;
        private Worksheet ws;

        public SimpleExcelAdapter(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            i += 1;
            j += 1;
            if (ws.Cells[i,j].Value2 != null)
            {
                return ws.Cells[i, j].Value2;
            }
            else
            {
                return "";
            }
        }
    }
}
