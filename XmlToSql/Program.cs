using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace XmlToSql
{
    class Program
    {
        static MCS db = new MCS();
        static void Main(string[] args)
        {
            ExcelPackage exp = new ExcelPackage();
            ExcelWorksheet worksheet = exp.Workbook.Worksheets.Add("List1");

            db.Area.ToList();

            int row = 2;
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Column(2).Width = 50;
            worksheet.Cells[1, 3].Value = "IP";
            worksheet.Column(3).Width = 11;

            foreach (Area area in db.Area)
            {
                worksheet.Cells[row, 1].Value = area.AreaId;
                worksheet.Cells[row, 2].Value = area.FullName;
                worksheet.Cells[row, 3].Value = area.IP;
                row++;
            }

            Dictionary<string, Area> dicIP = db.Area.Where(w => !string.IsNullOrEmpty(w.IP) && w.ParentId != 0).Select(s => new { s.IP }).Distinct().Select(s => new { ip = s.IP, area = db.Area.FirstOrDefault(f => f.IP == s.IP) }).ToDictionary(d => d.ip, d => d.area);

            ILookup<string, Area> lkp = db.Area.ToLookup(l => l.IP, l => l);

            ExcelWorksheet worksheet2 = exp.Workbook.Worksheets.Add("List2");
            row = 2;

            foreach (var item in dicIP)
            {
                worksheet2.Cells[row, 1].Value = item.Key;
                worksheet.Cells[row, 2].Value = item.Value.FullName;
                row++;
            }

            FileStream fs = File.Create("Excl.xlsx");
            fs.Close();
            exp.SaveAs(fs);
        }
    }
}
