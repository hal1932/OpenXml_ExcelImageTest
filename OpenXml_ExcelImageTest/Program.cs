using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Drawing;
using System.IO;

namespace OpenXml_ExcelImageTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var filepath = System.IO.Path.GetFullPath("test.xlsx");
            using (var doc = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {
                var bookpart = doc.AddWorkbookPart();
                bookpart.Workbook = new Workbook();

                var sheets = bookpart.Workbook.AppendChild<Sheets>(new Sheets());

                var sheetpart = bookpart.AddNewPart<WorksheetPart>();
                sheetpart.Worksheet = new Worksheet(new SheetData());

                var sheet = new Sheet()
                {
                    Id = bookpart.GetIdOfPart(sheetpart),
                    SheetId = 1,
                    Name = "sheet1",
                };
                sheets.Append(sheet);

                InsertImageTest(sheet, sheetpart);

                doc.Close();
            }
        }

        static void InsertImageTest(Sheet sheet, WorksheetPart sheetPart)
        {
            var filepath = System.IO.Path.GetFullPath(@"..\..\test.jpg");
            int widthPx = 100, heightPx = 100;
            int rowIndex = 2, colomnIndex = 3;
            int rowOffsetPx = 5, columnOffsetPx = 10;

            var imageType = ImagePartType.Jpeg;
            var noChangeAspect = true;
            var noCrop = false;
            var noMove = false;
            var noResize = false;
            var noRotation = false;
            var noSelection = false;

            float imageResX, imageResY;
            using (var bmp = Image.FromFile(filepath) as Bitmap)
            {
                imageResX = bmp.HorizontalResolution;
                imageResY = bmp.VerticalResolution;
            }
            var widthEmu = CalcEmuScale(widthPx, imageResX);
            var heightEmu = CalcEmuScale(heightPx, imageResY);
            var columnOffsetEmu = CalcEmuScale(columnOffsetPx, imageResY);
            var rowOffsetEmu = CalcEmuScale(rowOffsetPx, imageResX);

            var drawingsPart = sheetPart.DrawingsPart ?? sheetPart.AddNewPart<DrawingsPart>();

            if (!sheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
            {
                sheetPart.Worksheet.Append(new Drawing() { Id = sheetPart.GetIdOfPart(drawingsPart) });
            }

            if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            }
            var sheetDrawing = drawingsPart.WorksheetDrawing;

            var imagePart = drawingsPart.AddImagePart(imageType);

            using (var stream = new FileStream(filepath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            var nvps = sheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = (nvps.Count() > 0) ? (UInt32Value)nvps.Max(prop => prop.Id.Value) + 1 : 1U;

            var pictureLocks = new A.PictureLocks()
            {
                NoChangeAspect = noChangeAspect,
                NoCrop = noCrop,
                NoMove = noMove,
                NoResize = noResize,
                NoRotation = noRotation,
                NoSelection = noSelection,
            };

            var oneCellAnchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker()
                {
                    ColumnId = new Xdr.ColumnId((colomnIndex - 1).ToString()),
                    RowId = new Xdr.RowId((rowIndex - 1).ToString()),
                    ColumnOffset = new Xdr.ColumnOffset(columnOffsetEmu.ToString()),
                    RowOffset = new Xdr.RowOffset(rowOffsetEmu.ToString()),
                },

                new Xdr.Extent() { Cx = widthEmu, Cy = heightEmu, },

                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties() { Id = nvpId, Name = Path.GetFileName(filepath), Description = filepath, },
                        new Xdr.NonVisualPictureDrawingProperties(pictureLocks)
                        ),
                    new Xdr.BlipFill(
                        new A.Blip() { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())
                        ),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset() { X = 0, Y = 0, },
                            new A.Extents() { Cx = widthEmu, Cy = heightEmu }
                            ),
                        new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle }
                        )
                    ),
                new Xdr.ClientData());

            sheetDrawing.Append(oneCellAnchor);

            var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(sheetPart);
            foreach (var err in errors)
            {
                Console.WriteLine(err.Description);
            }
            if(errors.Count() > 0)
            {
                Console.Read();
            }
        }

        static long CalcEmuScale(int value, float resolusion)
        {
            return (long)(value * 914400.0F / resolusion);
        }
    }
}
