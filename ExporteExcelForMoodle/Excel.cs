using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExporteExcelForMoodle
{
    public class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;


        public Excel() { }

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        public string ReadCell(int i, int j)
        {
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2.ToString();
            else
                return "";
        }

        public void WriteStringToCell(int i, int j, string s)
        {
            ws.Cells[i, j].Value2 = s.ToString();
        }


        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        public void Close()
        {
            wb.Close();
            this.excel.Quit(); // need to checked
        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        }

        //public void CreateNewSheet()
        //{
        //    Worksheet tempSheet = wb.Worksheets.Add(After: ws);
        //}

        //public void SelectWorkSheet(int SheetNumber)
        //{
        //    this.ws = wb.Worksheets[SheetNumber];
        //}

        //public void DeleteWorkSheet(int SheetNumber)
        //{
        //    wb.Worksheets[SheetNumber].Delete();
        //}

        //public void ProtectSheet()
        //{
        //    ws.Protect();
        //}

        //public void ProtectSheet(string password)
        //{
        //    ws.Protect(password);
        //}

        //public void unProtectSheet()
        //{
        //    ws.Unprotect();
        //}

        public void leftToRight()
        {
            ws.DisplayRightToLeft = false;
        }

        public void unProtectSheet(string password)
        {
            ws.Unprotect(password);
        }



        public string[,] ReadRange(int starti, int startj, int endi, int endj)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, startj], ws.Cells[endi, endj]];
            object[,] holder = range.Value2;
            string[,] returnString = new string[endi - starti + 1, endj - startj + 1];
            for (int p = 0; p <= endi - starti; p++)
            {
                for (int q = 0; q <= endj - startj; q++)
                {
                    returnString[p, q] = holder[p + 1, q + 1].ToString();
                }
            }
            return returnString;
        }

        public List<string> ReadList(int starti, int startj, int endi, int endj)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, startj], ws.Cells[endi, endj]];
            object[,] holder = range.Value2;
            List<string> Read = new List<string>();
            for (int p = 0; p <= endi - starti; p++)
            {
                for (int q = 0; q <= endj - startj; q++)
                {
                    Read.Add(holder[p + 1, q + 1].ToString());
                }
            }
            return Read;
        }

        public void WriteRange(int starti, int startj, int endi, int endj, string[,] writeString)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, startj], ws.Cells[endi, endj]];
            range.Value2 = writeString;
        }

        public void WriteRange(int starti, int startj, int endi, int endj, string writeString)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, startj], ws.Cells[endi, endj]];
            range.Value2 = writeString;
        }


        public void WriteList(int starti, int startj, int endi, int endj, List<string> writeString)
        {
            for (int p = 0; p <= endi - starti;)
            {
                for (int q = 0; q <= endj - startj;)
                {
                    ws.Cells[starti, startj] = writeString.First<string>();
                    starti++;
                    writeString.RemoveAt(0);
                }
            }

        }
    }
}
