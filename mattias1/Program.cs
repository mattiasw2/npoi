using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Net.Mime;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using Org.BouncyCastle.Utilities.Collections;


namespace mattias1
{
    class Program
    {
        private const string Samples = @"C:\data-oss\cs\npoi\mattias1\samples\";
        static void Main(string[] args)
        {
            Console.WriteLine("Started");
            var filename = Samples + @"clothes.xlsx";
            FileStream file = File.OpenRead(filename);
            var workbook = new XSSFWorkbook(file);

            ImmutableList<XlsxSheet> sheets = GetAllSheets(workbook);


#if true
            var pictures = GetAllPictures(workbook);
#else 
            var pictures = ImmutableList<XlsxPicture>.Empty;
#endif

            var xlsx = new XlsxFile()
            {
                FileName = filename,
                Sheets = sheets,
                Pictures = pictures
            };
            Console.WriteLine(xlsx.ToString());
        }

        private static ImmutableList<XlsxSheet> GetAllSheets(XSSFWorkbook workbook)
        {
            ImmutableList<XlsxSheet> res = ImmutableList<XlsxSheet>.Empty;
            for (int ii = 0; ii < workbook.NumberOfSheets; ii++)
            {
                var sheet = GetSheet(workbook.GetSheetAt(ii), ii);
                res = res.Add(sheet);
            }

            return res;
        }

        private static XlsxSheet GetSheet(ISheet sheet, int sheetno)
        {
            ImmutableList<XlsxRow> rows = ImmutableList<XlsxRow>.Empty;
            for (int rowno = sheet.FirstRowNum; rowno < sheet.LastRowNum; rowno++)
            {
                rows = rows.Add(GetRow(sheet.GetRow(rowno), sheetno, rowno));
            }

            return new XlsxSheet()
            {
                name = sheet.SheetName,
                sheet = sheetno,
                Rows = rows
            };
        }

        private static XlsxRow GetRow(IRow row, int sheetno, int rowno)
        {
            var cells = ImmutableList<XlsxCell>.Empty;

            for (int colno = row.FirstCellNum; colno < row.LastCellNum; colno++)
            {
                cells = cells.Add(GetCell(row.GetCell(colno), sheetno, rowno, colno));
            }

            return new XlsxRow()
            {
                sheet = sheetno,
                row = rowno,
                Cells = cells
            };
        }

        private static XlsxCell GetCell(ICell getCell, int sheetno, int rowno, int colno)
        {
            return new XlsxCell()
            {
                col = colno,
                row = rowno,
                sheet = sheetno
            };
        }


        static ImmutableList<XlsxPicture> GetAllPictures(XSSFWorkbook workbook)
        {
            IList pictures = workbook.GetAllPictures();
            ImmutableList<XlsxPicture> res = ImmutableList<XlsxPicture>.Empty;
            int i = 0;
            foreach (IPictureData pic in pictures)
            {
                string ext = pic.SuggestFileExtension();
                string? filename;
                if (ext.Equals("jpeg"))
                {
                    Image jpg = Image.FromStream(new MemoryStream(pic.Data));
                    filename = string.Format("pic{0}.jpg", i++);
                    Console.WriteLine("saving file " + filename);
                    jpg.Save(filename);
                }
                else if (ext.Equals("png"))
                {
                    Image png = Image.FromStream(new MemoryStream(pic.Data));
                    filename = string.Format("pic{0}.png", i++);
                    Console.WriteLine("saving file " + filename);
                    png.Save(filename);
                }
                else
                {
                    throw new NotImplementedException("Strange picture");
                }

                res = res.Add(new XlsxPicture()
                {
                    FileName = filename,
                    MimeType = pic.MimeType,
                    PictureType = pic.PictureType,
                    SuggestFileExtension = pic.SuggestFileExtension()
                });
            }

            return res;
        }
    }


}
