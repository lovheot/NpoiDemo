using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NpoiDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "tpl", "1.xlsx");
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), Guid.NewGuid() + ".xlsx");
            var qrcodeData = File.ReadAllBytes("qrcode.png");
            using (var fs = File.OpenRead(tplPath))
            {
                var workbook = new XSSFWorkbook(fs);
                var sheet = workbook.GetSheetAt(0);
                for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;
                        //得到原来的值
                        string cellContent = cell.StringCellValue;
                        if (string.IsNullOrEmpty(cellContent)) continue;
                        //替换和字典对应的值
                        if (cellContent == "{{qrcode}}")
                        {
                            int pictureIndex = workbook.AddPicture(qrcodeData, PictureType.JPEG);
                            ICreationHelper helper = workbook.GetCreationHelper();
                            IDrawing drawing = sheet.CreateDrawingPatriarch();
                            IClientAnchor anchor = helper.CreateClientAnchor();
                            anchor.Col1 = j;
                            anchor.Row1 = i;
                            XSSFPicture picture = (XSSFPicture)drawing.CreatePicture(anchor, pictureIndex);
                            picture.LineStyle = LineStyle.DashDotGel;
                            picture.Resize();
                        }

                    }
                }
                using (FileStream localFs = File.OpenWrite(filePath))
                {
                    workbook.Write(localFs);
                    localFs.Close();
                }

                //一些列关闭释放操作
                fs.Close();
                fs.Dispose();
            }
        }
    }
}