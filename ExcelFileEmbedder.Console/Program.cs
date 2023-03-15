using System.Runtime.InteropServices;

namespace ExcelFileEmbedder.Console
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string outputExcelPath;
            string fromColumn;
            string fromRow;
            string toColumn;
            string toRow;
            string excelFilePath;
            string embedPath;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                excelFilePath = @"/mnt/c/test/openxml/Book1.xlsx";
                embedPath = @"/mnt/c/test/openxml/to_be_embeded.txt";
                outputExcelPath = @"/mnt/c/test/openxml/Book1-embeded.xlsx";
                fromColumn = "2";
                fromRow = "3";
                toColumn = "5";
                toRow = "6";
            }
            else
            {
                System.Console.WriteLine("Enter excel file path:");
                excelFilePath = System.Console.ReadLine();
                if (excelFilePath == "")
                {
                    excelFilePath = @"/mnt/c/test/openxml/Book1.xlsx";
                }
                System.Console.WriteLine("Enter path to file to be embedded:");
                embedPath = System.Console.ReadLine();
                if (embedPath == "")
                {
                    embedPath = @"/mnt/c/test/openxml/to_be_embeded.txt";
                }

                System.Console.WriteLine("Enter output path for the excel with embedded file:");
                outputExcelPath = System.Console.ReadLine();
                if (outputExcelPath == "")
                {
                    outputExcelPath = @"/mnt/c/test/openxml/Book1-embeded.xlsx";
                }

                System.Console.WriteLine("Enter top left column index position:");
                fromColumn = System.Console.ReadLine();

                System.Console.WriteLine("Enter top left row index position:");
                fromRow = System.Console.ReadLine();

                System.Console.WriteLine("Enter bottom right column index position:");
                toColumn = System.Console.ReadLine();

                System.Console.WriteLine("Enter bottom right row index position:");
                toRow = System.Console.ReadLine();
            }



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
                        Column = fromColumn,
                        Row = fromRow,
                    },
                    PositionTo = new ExcelFileTools.CellCoordinates
                    {
                        Column = toColumn,
                        Row = toRow,
                    },
                }
                );
            excelTools.EmbedFile();


            System.Console.WriteLine("Embedding successful");
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                System.Console.ReadKey();
            }
        }
    }
}