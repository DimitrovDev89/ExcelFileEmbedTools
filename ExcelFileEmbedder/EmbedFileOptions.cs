using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileTools
{
    public class CellCoordinates
    {
        public string Row { get; set; }
        public string Column { get; set; }
    }
    public class EmbedFileOptions
    {
        public string ExcelFilePath { get; set; }
        public Stream EmbedFileStream { get; set; }
        public string FileName { get; set; }
        public CellCoordinates PositionFrom { get; set; }
        public CellCoordinates PositionTo { get; set; }
        public string OutputExcelPath { get; set; }
    }
}
