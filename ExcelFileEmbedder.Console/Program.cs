namespace ExcelFileEmbedder.Console
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string embedPath = @"C:\test\openxml\to_be_embeded.txt";
            string excelFilePath = @"C:\test\openxml\Book1.xlsx";
            string outputExcelPath = @"C:\test\openxml\Book1-embeded.xlsx";
            using var fileEmbedStream = File.OpenRead(embedPath);

            var excelTools = new ExcelTools(
                new ExcelFileTools.EmbedFileOptions
                {
                    EmbedFileStream = fileEmbedStream,
                    ExcelFilePath = excelFilePath,
                    FileName = Path.GetFileName(embedPath),
                    OutputExcelPath = outputExcelPath,
                    PositionFrom = new ExcelFileTools.CellCoordinates
                    {
                        Column = "2",
                        Row = "2",
                    },
                    PositionTo = new ExcelFileTools.CellCoordinates
                    {
                        Column = "5",
                        Row = "5",
                    },
                }
                );
            excelTools.EmbedFile();
        }
    }
}